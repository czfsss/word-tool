# from collections.abc import Generator
# from typing import Any
# import tempfile
# import os
# import json
# from datetime import datetime
# from dify_plugin.file.file import File
# from dify_plugin import Tool
# from dify_plugin.entities.tool import ToolInvokeMessage
# from docx import Document
# import docx

# from tools.utils.logger_utils import get_logger
# from tools.utils.file_utils import get_meta_data, sanitize_filename


# class WordCommentTool(Tool):
#     # 获取当前模块的日志记录器
#     logger = get_logger(__name__)

#     def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
#         # 检查python-docx版本
#         self.logger.info(f"当前python-docx版本: {docx.__version__}")

#         # 获取上传的Word文件
#         word_content: File = tool_parameters.get("word_content")
#         comments_json: str = tool_parameters.get("comments_json", "{}")

#         if not word_content:
#             self.logger.error("未提供Word文件")
#             yield self.create_text_message("请提供Word文件")
#             return

#         # 检查文件类型
#         if not isinstance(word_content, File):
#             self.logger.error("无效的文件格式，期望File对象")
#             yield self.create_text_message("无效的文件格式，期望File对象")
#             return

#         # 解析批注JSON
#         try:
#             comments_dict = json.loads(comments_json)
#             if not isinstance(comments_dict, dict):
#                 self.logger.error("批注JSON格式错误，期望对象格式")
#                 yield self.create_text_message("批注JSON格式错误，期望对象格式")
#                 return
#         except json.JSONDecodeError as e:
#             self.logger.error(f"JSON解析失败: {str(e)}")
#             yield self.create_text_message(f"JSON解析失败: {str(e)}")
#             return

#         # 获取批注者信息和自定义文件名
#         author = tool_parameters.get("author", "批注者")
#         custom_filename = tool_parameters.get("output_filename", "").strip()

#         self.logger.info(
#             f"开始处理Word文档批注，文件名: {word_content.filename if word_content.filename else '未知'}，批注数量: {len(comments_dict)}，批注者: {author}，自定义文件名: {custom_filename if custom_filename else '未设置'}"
#         )

#         try:
#             # 创建临时文件保存上传的Word文件
#             with tempfile.NamedTemporaryFile(
#                 delete=False, suffix=".docx"
#             ) as temp_input_file:
#                 # 获取文件内容（字节）并写入临时文件
#                 temp_input_file.write(word_content.blob)
#                 temp_input_path = temp_input_file.name

#             # 创建临时文件保存处理后的Word文件
#             with tempfile.NamedTemporaryFile(
#                 delete=False, suffix=".docx"
#             ) as temp_output_file:
#                 temp_output_path = temp_output_file.name

#             # 调用批注添加函数
#             self.logger.info("开始添加批注到Word文档")
#             comment_count = self.add_native_comments_to_document(
#                 temp_input_path, temp_output_path, comments_dict, author
#             )
#             self.logger.info(f"成功添加了 {comment_count} 个批注")

#             # 读取处理后的Word文件内容
#             with open(temp_output_path, "rb") as docx_file:
#                 docx_blob = docx_file.read()

#             # 处理自定义文件名
#             processed_filename = None
#             if custom_filename:
#                 # 清理并处理文件名
#                 processed_filename = sanitize_filename(custom_filename)
#                 self.logger.info(f"使用自定义文件名: {processed_filename}")
#             else:
#                 # 使用原始文件名的基本名称（去掉扩展名）
#                 if word_content.filename:
#                     base_name = os.path.splitext(word_content.filename)[0]
#                     processed_filename = f"{base_name}_commented"
#                 else:
#                     processed_filename = "commented_document"
#                 self.logger.info(f"使用默认文件名: {processed_filename}")

#             # 使用create_blob_message返回处理后的Word文件
#             yield self.create_blob_message(
#                 blob=docx_blob,
#                 meta=get_meta_data(
#                     mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#                     output_filename=processed_filename,
#                 ),
#             )

#             # 清理临时文件
#             os.unlink(temp_input_path)
#             os.unlink(temp_output_path)
#             self.logger.info("Word批注处理完成，临时文件已清理")

#         except Exception as e:
#             self.logger.exception("处理Word文档批注时发生异常")
#             yield self.create_text_message(f"处理Word文档批注时出错: {str(e)}")

#     def add_native_comments_to_document(
#         self,
#         input_path: str,
#         output_path: str,
#         comments_dict: dict,
#         author: str = "批注者",
#     ) -> int:
#         """
#         向Word文档添加真正的批注（使用python-docx原生批注API）

#         Args:
#             input_path: 输入Word文档路径
#             output_path: 输出Word文档路径
#             comments_dict: 批注字典，key为摘要文本，value为批注内容
#             author: 批注者姓名

#         Returns:
#             int: 成功添加的批注数量
#         """
#         doc = Document(input_path)
#         comment_count = 0

#         # 获取作者缩写（取前两个字符）
#         initials = author[:2] if len(author) >= 2 else author

#         # 遍历文档中的所有段落
#         for paragraph in doc.paragraphs:
#             comment_count += self._process_paragraph_comments(
#                 doc, paragraph, comments_dict, author, initials
#             )

#         # 检查表格中的内容
#         for table in doc.tables:
#             comment_count += self._process_table_comments(
#                 doc, table, comments_dict, author, initials
#             )

#         # 保存文档
#         doc.save(output_path)
#         return comment_count

#     def _process_paragraph_comments(
#         self, doc, paragraph, comments_dict, author, initials
#     ):
#         """处理段落中的批注"""
#         comment_count = 0
#         paragraph_text = paragraph.text

#         for summary, comment_text in comments_dict.items():
#             if summary in paragraph_text:
#                 try:
#                     self.logger.debug(f"在段落中找到摘要 '{summary}'，开始添加批注")

#                     # 使用原生批注API添加批注
#                     success = self._add_native_comment_to_paragraph(
#                         doc, paragraph, summary, comment_text, author, initials
#                     )
#                     if success:
#                         comment_count += 1
#                         self.logger.debug(
#                             f"成功为摘要 '{summary}' 添加批注: {comment_text}"
#                         )
#                 except Exception as e:
#                     self.logger.warning(f"为摘要 '{summary}' 添加批注时出错: {str(e)}")
#                     continue

#         return comment_count

#     def _process_table_comments(self, doc, table, comments_dict, author, initials):
#         """处理表格中的批注"""
#         comment_count = 0

#         for row in table.rows:
#             for cell in row.cells:
#                 for paragraph in cell.paragraphs:
#                     comment_count += self._process_paragraph_comments(
#                         doc, paragraph, comments_dict, author, initials
#                     )

#         return comment_count

#     def _add_native_comment_to_paragraph(
#         self, doc, paragraph, target_text, comment_text, author, initials
#     ):
#         """
#         使用python-docx原生API向段落添加批注

#         Args:
#             doc: Word文档对象
#             paragraph: 段落对象
#             target_text: 要批注的目标文本
#             comment_text: 批注内容
#             author: 批注者
#             initials: 批注者缩写

#         Returns:
#             bool: 是否成功添加批注
#         """
#         try:
#             # 查找包含目标文本的Run
#             target_run = None
#             for run in paragraph.runs:
#                 if target_text in run.text:
#                     target_run = run
#                     break

#             if not target_run:
#                 self.logger.warning(f"未找到包含文本 '{target_text}' 的Run")
#                 return False

#             # 如果目标文本只是Run的一部分，需要分割Run
#             if target_run.text.strip() != target_text.strip():
#                 target_run = self._split_run_for_comment(
#                     paragraph, target_run, target_text
#                 )
#                 if not target_run:
#                     return False

#             # 使用python-docx 1.2.0的原生批注API
#             try:
#                 comment = doc.add_comment(
#                     runs=target_run, text=comment_text, author=author, initials=initials
#                 )
#                 self.logger.debug(f"成功使用原生API添加批注")
#                 return True
#             except AttributeError:
#                 # 如果不支持原生批注API，使用备用方案
#                 self.logger.warning("当前版本不支持原生批注API，使用备用方案")
#                 return self._add_fallback_comment(target_run, comment_text, author)

#         except Exception as e:
#             self.logger.error(f"添加原生批注时出错: {str(e)}")
#             return False

#     def _split_run_for_comment(self, paragraph, run, target_text):
#         """
#         分割Run以精确定位批注文本

#         Args:
#             paragraph: 段落对象
#             run: 要分割的Run对象
#             target_text: 目标文本

#         Returns:
#             Run: 包含目标文本的新Run对象
#         """
#         try:
#             # 使用partition将文本拆分为：前缀、目标、后缀
#             before, matched, after = run.text.partition(target_text)

#             if not matched:
#                 return None

#             # 保存原Run的样式
#             original_bold = run.bold
#             original_italic = run.italic
#             original_underline = run.underline
#             original_font_size = run.font.size
#             original_font_name = run.font.name

#             # 更新原Run为前缀文本
#             run.text = before

#             # 创建目标文本的新Run
#             target_run = paragraph.add_run(matched)
#             target_run.bold = original_bold
#             target_run.italic = original_italic
#             target_run.underline = original_underline
#             if original_font_size:
#                 target_run.font.size = original_font_size
#             if original_font_name:
#                 target_run.font.name = original_font_name

#             # 创建后缀文本的新Run
#             if after:
#                 after_run = paragraph.add_run(after)
#                 after_run.bold = original_bold
#                 after_run.italic = original_italic
#                 after_run.underline = original_underline
#                 if original_font_size:
#                     after_run.font.size = original_font_size
#                 if original_font_name:
#                     after_run.font.name = original_font_name

#             return target_run

#         except Exception as e:
#             self.logger.error(f"分割Run时出错: {str(e)}")
#             return None

#     def _add_fallback_comment(self, run, comment_text, author):
#         """
#         备用批注方案（当不支持原生API时）

#         Args:
#             run: Run对象
#             comment_text: 批注内容
#             author: 批注者

#         Returns:
#             bool: 是否成功添加
#         """
#         try:
#             # 高亮目标文本
#             run.font.highlight_color = 7  # 黄色高亮

#             # 在文本后添加批注标记
#             original_text = run.text
#             current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
#             run.text = (
#                 f"{original_text} [批注by {author} {current_time}: {comment_text}]"
#             )

#             self.logger.debug(f"使用备用方案添加批注")
#             return True

#         except Exception as e:
#             self.logger.error(f"添加备用批注时出错: {str(e)}")
#             return False


from collections.abc import Generator
from typing import Any
import tempfile
import os
import json
import re
from datetime import datetime
from difflib import SequenceMatcher
from dify_plugin.file.file import File
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from docx import Document
import docx

from tools.utils.logger_utils import get_logger
from tools.utils.file_utils import get_meta_data, sanitize_filename


class WordCommentTool(Tool):
    # 获取当前模块的日志记录器
    logger = get_logger(__name__)

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """
        计算两个文本的相似度

        Args:
            text1: 第一个文本
            text2: 第二个文本

        Returns:
            float: 相似度，范围0-1
        """
        # 清理文本，去除多余空格和标点符号对比
        clean_text1 = re.sub(r"\s+", "", text1.strip())
        clean_text2 = re.sub(r"\s+", "", text2.strip())

        # 使用SequenceMatcher计算相似度
        similarity = SequenceMatcher(None, clean_text1, clean_text2).ratio()
        return similarity

    def _find_fuzzy_match(
        self, target_text: str, paragraph_text: str, threshold: float = 0.8
    ) -> tuple:
        """
        在段落中寻找与目标文本模糊匹配的部分，支持多句话和多段落匹配

        Args:
            target_text: 目标文本（批注的key）
            paragraph_text: 段落文本
            threshold: 相似度阈值（默认0.8即80%）

        Returns:
            tuple: (是否找到匹配, 匹配的文本, 相似度)
        """
        # 首先尝试精确匹配
        if target_text in paragraph_text:
            return True, target_text, 1.0

        # 如果精确匹配失败，尝试模糊匹配
        best_match = None
        best_similarity = 0.0
        best_sentence = None
        
        # 1. 改进的句子分割策略，支持更多标点符号
        sentences = re.split(r"[。！？.!?；;,、\n\r]+", paragraph_text)
        
        # 2. 尝试单个句子匹配
        for sentence in sentences:
            sentence = sentence.strip()
            if len(sentence) < 5:  # 降低最小长度阈值
                continue

            similarity = self._calculate_similarity(target_text, sentence)
            if similarity > best_similarity and similarity >= threshold:
                best_similarity = similarity
                best_match = sentence
                best_sentence = sentence
        
        # 3. 如果目标文本包含多个句子，尝试多句子组合匹配
        if not best_match and len(target_text) > 30:
            # 尝试相邻句子组合
            for i in range(len(sentences) - 1):
                # 尝试两个相邻句子
                combined_sentence = (sentences[i].strip() + " " + sentences[i+1].strip()).strip()
                if len(combined_sentence) < 10:
                    continue
                    
                similarity = self._calculate_similarity(target_text, combined_sentence)
                if similarity > best_similarity and similarity >= threshold:
                    best_similarity = similarity
                    best_match = combined_sentence
                    best_sentence = combined_sentence
                
                # 尝试三个相邻句子
                if i < len(sentences) - 2:
                    combined_three = (combined_sentence + " " + sentences[i+2].strip()).strip()
                    if len(combined_three) < 15:
                        continue
                        
                    similarity = self._calculate_similarity(target_text, combined_three)
                    if similarity > best_similarity and similarity >= threshold:
                        best_similarity = similarity
                        best_match = combined_three
                        best_sentence = combined_three
        
        # 4. 如果没有找到句子级别的匹配，尝试滑动窗口匹配
        if not best_match and len(target_text) > 20:
            target_len = len(target_text)
            # 动态调整窗口大小
            window_size = max(target_len - 10, target_len // 2, 30)
            
            # 使用不同的窗口大小进行匹配
            for window_factor in [1.0, 0.8, 0.6, 1.2]:
                current_window_size = int(window_size * window_factor)
                
                for i in range(len(paragraph_text) - current_window_size + 1):
                    window_text = paragraph_text[i : i + current_window_size]
                    similarity = self._calculate_similarity(target_text, window_text)
                    
                    if similarity > best_similarity and similarity >= threshold:
                        best_similarity = similarity
                        best_match = window_text
                        
                # 如果已经找到匹配，可以提前退出
                if best_match:
                    break
        
        # 5. 如果仍然没有找到匹配，尝试关键词匹配
        if not best_match:
            # 提取目标文本中的关键词（去除常见停用词）
            target_words = re.findall(r'\b\w+\b', target_text.lower())
            # 简单过滤掉一些常见词
            stop_words = {'的', '了', '和', '是', '在', '我', '有', '这', '个', '那', '你', '会', '说', 'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'must', 'can', 'this', 'that', 'these', 'those'}
            key_words = [word for word in target_words if word not in stop_words and len(word) > 1]
            
            if key_words:
                # 计算段落中包含的关键词比例
                paragraph_words = re.findall(r'\b\w+\b', paragraph_text.lower())
                matched_words = [word for word in key_words if word in paragraph_words]
                
                if matched_words:
                    keyword_ratio = len(matched_words) / len(key_words)
                    # 如果关键词匹配比例较高，则认为找到匹配
                    if keyword_ratio >= 0.7:  # 70%的关键词匹配
                        # 尝试找到包含最多关键词的文本片段
                        best_match = self._find_best_keyword_match(paragraph_text, key_words)
                        best_similarity = max(threshold, keyword_ratio)  # 使用关键词比例作为相似度

        if best_match:
            return True, best_match, best_similarity
        else:
            return False, None, 0.0

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 检查python-docx版本
        self.logger.info(f"当前python-docx版本: {docx.__version__}")

        # 获取上传的Word文件
        word_content = tool_parameters.get("word_content")
        comments_json: str = tool_parameters.get("comments_json", "{}")

        if not word_content:
            self.logger.error("未提供Word文件")
            yield self.create_text_message("请提供Word文件")
            return

        # 检查文件类型
        if not isinstance(word_content, File):
            self.logger.error("无效的文件格式，期望File对象")
            yield self.create_text_message("无效的文件格式，期望File对象")
            return

        # 解析批注JSON
        try:
            comments_data = json.loads(comments_json)

            # 支持新的数组格式：[{"原文1":"批注1","原文2":"批注2"},{...}]
            if isinstance(comments_data, list):
                # 合并所有数组元素中的批注对
                comments_dict = {}
                processed_groups = 0
                total_comments = 0

                for i, comment_group in enumerate(comments_data):
                    if isinstance(comment_group, dict):
                        # 过滤掉空对象和无效的批注对
                        valid_comments = {}
                        for key, value in comment_group.items():
                            # 跳过空键、空值或分隔符
                            if (
                                key
                                and value
                                and str(key).strip() != ""
                                and str(value).strip() != ""
                                and not str(key).startswith("--")
                                and str(key) != "合同原文的问题句"
                            ):
                                valid_comments[key] = value

                        if valid_comments:
                            comments_dict.update(valid_comments)
                            processed_groups += 1
                            total_comments += len(valid_comments)
                            self.logger.info(
                                f"处理第{i+1}个对象组，包含{len(valid_comments)}个有效批注对"
                            )
                        else:
                            self.logger.debug(
                                f"第{i+1}个对象组为空或无有效批注，已跳过"
                            )
                    else:
                        self.logger.warning(
                            f"批注JSON数组中第{i+1}个元素不是对象格式，已跳过"
                        )

                self.logger.info(
                    f"解析数组格式批注JSON成功，处理了 {processed_groups} 个有效对象组，共合并 {len(comments_dict)} 个批注对"
                )

                if len(comments_dict) == 0:
                    self.logger.warning("JSON数组中未找到有效的批注对")
                    yield self.create_text_message(
                        "JSON数组中未找到有效的批注对，请检查数据格式"
                    )
                    return

            elif isinstance(comments_data, dict):
                # 兼容旧的对象格式：{"原文":"批注"}
                # 同样过滤无效的批注对
                comments_dict = {}
                for key, value in comments_data.items():
                    if (
                        key
                        and value
                        and str(key).strip() != ""
                        and str(value).strip() != ""
                        and not str(key).startswith("--")
                        and str(key) != "合同原文的问题句"
                    ):
                        comments_dict[key] = value

                self.logger.info(
                    f"解析对象格式批注JSON成功，共 {len(comments_dict)} 个有效批注对"
                )

                if len(comments_dict) == 0:
                    self.logger.warning("JSON对象中未找到有效的批注对")
                    yield self.create_text_message(
                        "JSON对象中未找到有效的批注对，请检查数据格式"
                    )
                    return
            else:
                self.logger.error("批注JSON格式错误，期望数组或对象格式")
                yield self.create_text_message(
                    "批注JSON格式错误，期望数组格式 [{},...] 或对象格式 {}"
                )
                return

        except json.JSONDecodeError as e:
            self.logger.error(f"JSON解析失败: {str(e)}")
            yield self.create_text_message(f"JSON解析失败: {str(e)}")
            return

        # 获取批注者信息、自定义文件名和相似度阈值
        author = tool_parameters.get("author", "批注者")
        custom_filename = tool_parameters.get("output_filename", "").strip()
        similarity_threshold = tool_parameters.get("similarity_threshold", 0.8)

        # 验证相似度阈值范围
        if not isinstance(similarity_threshold, (int, float)) or not (
            0.1 <= similarity_threshold <= 1.0
        ):
            similarity_threshold = 0.8
            self.logger.warning(f"相似度阈值无效，使用默认值 0.8")

        self.logger.info(
            f"开始处理Word文档批注，文件名: {word_content.filename if word_content.filename else '未知'}，批注数量: {len(comments_dict)}，批注者: {author}，自定义文件名: {custom_filename if custom_filename else '未设置'}，相似度阈值: {similarity_threshold:.2%}"
        )

        try:
            # 创建临时文件保存上传的Word文件
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".docx"
            ) as temp_input_file:
                # 获取文件内容（字节）并写入临时文件
                temp_input_file.write(word_content.blob)
                temp_input_path = temp_input_file.name

            # 创建临时文件保存处理后的Word文件
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".docx"
            ) as temp_output_file:
                temp_output_path = temp_output_file.name

            # 调用批注添加函数
            self.logger.info("开始添加批注到Word文档")
            comment_count = self.add_native_comments_to_document(
                temp_input_path,
                temp_output_path,
                comments_dict,
                author,
                similarity_threshold,
            )
            self.logger.info(f"成功添加了 {comment_count} 个批注")

            # 读取处理后的Word文件内容
            with open(temp_output_path, "rb") as docx_file:
                docx_blob = docx_file.read()

            # 处理自定义文件名
            processed_filename = None
            if custom_filename:
                # 清理并处理文件名
                processed_filename = sanitize_filename(custom_filename)
                self.logger.info(f"使用自定义文件名: {processed_filename}")
            else:
                # 使用原始文件名的基本名称（去掉扩展名）
                if word_content.filename:
                    base_name = os.path.splitext(word_content.filename)[0]
                    processed_filename = f"{base_name}_commented"
                else:
                    processed_filename = "commented_document"
                self.logger.info(f"使用默认文件名: {processed_filename}")

            # 使用create_blob_message返回处理后的Word文件
            yield self.create_blob_message(
                blob=docx_blob,
                meta=get_meta_data(
                    mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    output_filename=processed_filename,
                ),
            )

            # 清理临时文件
            os.unlink(temp_input_path)
            os.unlink(temp_output_path)
            self.logger.info("Word批注处理完成，临时文件已清理")

        except Exception as e:
            self.logger.exception("处理Word文档批注时发生异常")
            yield self.create_text_message(f"处理Word文档批注时出错: {str(e)}")

    def add_native_comments_to_document(
        self,
        input_path: str,
        output_path: str,
        comments_dict: dict,
        author: str = "批注者",
        similarity_threshold: float = 0.8,
    ) -> int:
        """
        向Word文档添加真正的批注（使用python-docx原生批注API，支持模糊匹配）

        Args:
            input_path: 输入Word文档路径
            output_path: 输出Word文档路径
            comments_dict: 批注字典，key为摘要文本，value为批注内容
            author: 批注者姓名
            similarity_threshold: 模糊匹配的相似度阈值（0.1-1.0）

        Returns:
            int: 成功添加的批注数量
        """
        doc = Document(input_path)
        comment_count = 0

        # 获取作者缩写（取前两个字符）
        initials = author[:2] if len(author) >= 2 else author

        # 记录所有可用的批注键
        self.logger.info(
            f"开始处理批注，共有 {len(comments_dict)} 个批注对，相似度阈值: {similarity_threshold:.2%}:"
        )
        for i, (key, value) in enumerate(comments_dict.items(), 1):
            self.logger.debug(f"批注{i}: '{key[:50]}...' -> '{value[:50]}...'")

        # 遍历文档中的所有段落
        total_paragraphs = len(doc.paragraphs)
        processed_paragraphs = 0

        self.logger.info(f"开始处理 {total_paragraphs} 个段落")

        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():  # 只处理非空段落
                processed_paragraphs += 1
                paragraph_comments = self._process_paragraph_comments(
                    doc,
                    paragraph,
                    comments_dict,
                    author,
                    initials,
                    similarity_threshold,
                )
                comment_count += paragraph_comments

                if paragraph_comments > 0:
                    self.logger.debug(f"段落 {i+1} 添加了 {paragraph_comments} 个批注")

        # 检查表格中的内容
        total_tables = len(doc.tables)
        if total_tables > 0:
            self.logger.info(f"开始处理 {total_tables} 个表格")

            for i, table in enumerate(doc.tables):
                table_comments = self._process_table_comments(
                    doc, table, comments_dict, author, initials, similarity_threshold
                )
                comment_count += table_comments

                if table_comments > 0:
                    self.logger.debug(f"表格 {i+1} 添加了 {table_comments} 个批注")

        self.logger.info(
            f"文档处理完成：处理了 {processed_paragraphs} 个段落和 {total_tables} 个表格，成功添加 {comment_count} 个批注"
        )

        # 保存文档
        doc.save(output_path)
        return comment_count

    def _process_paragraph_comments(
        self, doc, paragraph, comments_dict, author, initials, similarity_threshold=0.8
    ):
        """处理段落中的批注（支持模糊匹配）"""
        comment_count = 0
        paragraph_text = paragraph.text

        # 跳过空段落
        if not paragraph_text.strip():
            return comment_count

        self.logger.debug(f"正在处理段落：{paragraph_text[:100]}...")

        # 统计在当前段落中找到的批注数量
        found_comments = []

        for summary, comment_text in comments_dict.items():
            # 跳过空的或无效的批注
            if not summary or not comment_text:
                continue

            # 使用模糊匹配查找相似的文本
            found, matched_text, similarity = self._find_fuzzy_match(
                summary, paragraph_text, threshold=similarity_threshold
            )

            if found:
                found_comments.append((summary, comment_text, matched_text, similarity))

                if similarity == 1.0:
                    self.logger.debug(f"在段落中精确匹配到摘要: '{summary[:50]}...'")
                else:
                    self.logger.info(
                        f"在段落中模糊匹配到摘要 (相似度: {similarity:.2%}):"
                    )
                    self.logger.info(f"  原摘要: '{summary[:50]}...'")
                    self.logger.info(f"  匹配文本: '{matched_text[:50]}...'")

                try:
                    # 使用匹配到的文本添加批注
                    success = self._add_native_comment_to_paragraph(
                        doc, paragraph, matched_text, comment_text, author, initials
                    )
                    if success:
                        comment_count += 1
                        self.logger.info(
                            f"成功为摘要 '{summary[:30]}...' 添加批注 (相似度: {similarity:.2%})"
                        )
                    else:
                        self.logger.warning(f"为摘要 '{summary[:30]}...' 添加批注失败")
                except Exception as e:
                    self.logger.warning(
                        f"为摘要 '{summary[:30]}...' 添加批注时出错: {str(e)}"
                    )
                    continue
            else:
                # 记录未找到匹配的情况
                self.logger.debug(f"未找到匹配 (相似度 < 80%): '{summary[:50]}...'")

        if found_comments:
            self.logger.debug(
                f"在当前段落中找到 {len(found_comments)} 个匹配的批注，成功添加 {comment_count} 个"
            )

        return comment_count

    def _process_table_comments(
        self, doc, table, comments_dict, author, initials, similarity_threshold=0.8
    ):
        """处理表格中的批注"""
        comment_count = 0

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    comment_count += self._process_paragraph_comments(
                        doc,
                        paragraph,
                        comments_dict,
                        author,
                        initials,
                        similarity_threshold,
                    )

        return comment_count

    def _add_native_comment_to_paragraph(
        self, doc, paragraph, target_text, comment_text, author, initials
    ):
        """
        使用python-docx原生API向段落添加批注

        Args:
            doc: Word文档对象
            paragraph: 段落对象
            target_text: 要批注的目标文本
            comment_text: 批注内容
            author: 批注者
            initials: 批注者缩写

        Returns:
            bool: 是否成功添加批注
        """
        try:
            # 查找包含目标文本的Run
            target_run = None
            for run in paragraph.runs:
                if target_text in run.text:
                    target_run = run
                    break

            if not target_run:
                self.logger.warning(f"未找到包含文本 '{target_text}' 的Run")
                return False

            # 如果目标文本只是Run的一部分，需要分割Run
            if target_run.text.strip() != target_text.strip():
                target_run = self._split_run_for_comment(
                    paragraph, target_run, target_text
                )
                if not target_run:
                    return False

            # 使用python-docx 1.2.0的原生批注API
            try:
                comment = doc.add_comment(
                    runs=target_run, text=comment_text, author=author, initials=initials
                )
                self.logger.debug(f"成功使用原生API添加批注")
                return True
            except AttributeError:
                # 如果不支持原生批注API，使用备用方案
                self.logger.warning("当前版本不支持原生批注API，使用备用方案")
                return self._add_fallback_comment(target_run, comment_text, author)

        except Exception as e:
            self.logger.error(f"添加原生批注时出错: {str(e)}")
            return False

    def _split_run_for_comment(self, paragraph, run, target_text):
        """
        分割Run以精确定位批注文本

        Args:
            paragraph: 段落对象
            run: 要分割的Run对象
            target_text: 目标文本

        Returns:
            Run: 包含目标文本的新Run对象
        """
        try:
            # 使用partition将文本拆分为：前缀、目标、后缀
            before, matched, after = run.text.partition(target_text)

            if not matched:
                return None

            # 保存原Run的样式
            original_bold = run.bold
            original_italic = run.italic
            original_underline = run.underline
            original_font_size = run.font.size
            original_font_name = run.font.name

            # 更新原Run为前缀文本
            run.text = before

            # 创建目标文本的新Run
            target_run = paragraph.add_run(matched)
            target_run.bold = original_bold
            target_run.italic = original_italic
            target_run.underline = original_underline
            if original_font_size:
                target_run.font.size = original_font_size
            if original_font_name:
                target_run.font.name = original_font_name

            # 创建后缀文本的新Run
            if after:
                after_run = paragraph.add_run(after)
                after_run.bold = original_bold
                after_run.italic = original_italic
                after_run.underline = original_underline
                if original_font_size:
                    after_run.font.size = original_font_size
                if original_font_name:
                    after_run.font.name = original_font_name

            return target_run

        except Exception as e:
            self.logger.error(f"分割Run时出错: {str(e)}")
            return None

    def _find_best_keyword_match(self, paragraph_text: str, key_words: list) -> str:
        """
        在段落中找到包含最多关键词的文本片段

        Args:
            paragraph_text: 段落文本
            key_words: 关键词列表

        Returns:
            str: 包含最多关键词的文本片段
        """
        # 将段落按句子分割
        sentences = re.split(r"[。！？.!?；;,、\n\r]+", paragraph_text)
        
        best_sentence = ""
        best_score = 0
        
        for sentence in sentences:
            sentence = sentence.strip()
            if len(sentence) < 5:  # 跳过太短的句子
                continue
                
            # 计算句子中包含的关键词数量
            sentence_words = re.findall(r'\b\w+\b', sentence.lower())
            matched_count = sum(1 for word in key_words if word in sentence_words)
            
            # 计算得分：考虑关键词数量和句子长度
            if len(sentence_words) > 0:
                score = matched_count / len(sentence_words)  # 关键词密度
                if matched_count > best_score or (matched_count == best_score and len(sentence) < len(best_sentence)):
                    best_score = matched_count
                    best_sentence = sentence
        
        # 如果单个句子匹配效果不佳，尝试相邻句子组合
        if best_score < len(key_words) * 0.5 and len(sentences) > 1:  # 如果匹配的关键词少于50%
            for i in range(len(sentences) - 1):
                combined = (sentences[i].strip() + " " + sentences[i+1].strip()).strip()
                if len(combined) < 10:
                    continue
                    
                combined_words = re.findall(r'\b\w+\b', combined.lower())
                matched_count = sum(1 for word in key_words if word in combined_words)
                
                if len(combined_words) > 0:
                    score = matched_count / len(combined_words)
                    if score > best_score:
                        best_score = score
                        best_sentence = combined
        
        return best_sentence if best_sentence else paragraph_text[:100]  # 如果没有找到合适的句子，返回段落前100个字符
    
    def _add_fallback_comment(self, run, comment_text, author):
        """
        备用批注方案（当不支持原生API时）

        Args:
            run: Run对象
            comment_text: 批注内容
            author: 批注者

        Returns:
            bool: 是否成功添加
        """
        try:
            # 高亮目标文本
            run.font.highlight_color = 7  # 黄色高亮

            # 在文本后添加批注标记
            original_text = run.text
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
            run.text = (
                f"{original_text} [批注by {author} {current_time}: {comment_text}]"
            )

            self.logger.debug(f"使用备用方案添加批注")
            return True

        except Exception as e:
            self.logger.error(f"添加备用批注时出错: {str(e)}")
            return False
