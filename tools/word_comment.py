from collections.abc import Generator
from typing import Any
import tempfile
import os
import json
from datetime import datetime
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

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 检查python-docx版本
        self.logger.info(f"当前python-docx版本: {docx.__version__}")

        # 获取上传的Word文件
        word_content: File = tool_parameters.get("word_content")
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
            comments_dict = json.loads(comments_json)
            if not isinstance(comments_dict, dict):
                self.logger.error("批注JSON格式错误，期望对象格式")
                yield self.create_text_message("批注JSON格式错误，期望对象格式")
                return
        except json.JSONDecodeError as e:
            self.logger.error(f"JSON解析失败: {str(e)}")
            yield self.create_text_message(f"JSON解析失败: {str(e)}")
            return

        # 获取批注者信息和自定义文件名
        author = tool_parameters.get("author", "批注者")
        custom_filename = tool_parameters.get("output_filename", "").strip()

        self.logger.info(
            f"开始处理Word文档批注，文件名: {word_content.filename if word_content.filename else '未知'}，批注数量: {len(comments_dict)}，批注者: {author}，自定义文件名: {custom_filename if custom_filename else '未设置'}"
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
                temp_input_path, temp_output_path, comments_dict, author
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
    ) -> int:
        """
        向Word文档添加真正的批注（使用python-docx原生批注API）

        Args:
            input_path: 输入Word文档路径
            output_path: 输出Word文档路径
            comments_dict: 批注字典，key为摘要文本，value为批注内容
            author: 批注者姓名

        Returns:
            int: 成功添加的批注数量
        """
        doc = Document(input_path)
        comment_count = 0

        # 获取作者缩写（取前两个字符）
        initials = author[:2] if len(author) >= 2 else author

        # 遍历文档中的所有段落
        for paragraph in doc.paragraphs:
            comment_count += self._process_paragraph_comments(
                doc, paragraph, comments_dict, author, initials
            )

        # 检查表格中的内容
        for table in doc.tables:
            comment_count += self._process_table_comments(
                doc, table, comments_dict, author, initials
            )

        # 保存文档
        doc.save(output_path)
        return comment_count

    def _process_paragraph_comments(
        self, doc, paragraph, comments_dict, author, initials
    ):
        """处理段落中的批注"""
        comment_count = 0
        paragraph_text = paragraph.text

        for summary, comment_text in comments_dict.items():
            if summary in paragraph_text:
                try:
                    self.logger.debug(f"在段落中找到摘要 '{summary}'，开始添加批注")

                    # 使用原生批注API添加批注
                    success = self._add_native_comment_to_paragraph(
                        doc, paragraph, summary, comment_text, author, initials
                    )
                    if success:
                        comment_count += 1
                        self.logger.debug(
                            f"成功为摘要 '{summary}' 添加批注: {comment_text}"
                        )
                except Exception as e:
                    self.logger.warning(f"为摘要 '{summary}' 添加批注时出错: {str(e)}")
                    continue

        return comment_count

    def _process_table_comments(self, doc, table, comments_dict, author, initials):
        """处理表格中的批注"""
        comment_count = 0

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    comment_count += self._process_paragraph_comments(
                        doc, paragraph, comments_dict, author, initials
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
