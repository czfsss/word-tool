from collections.abc import Generator
from typing import Any
import tempfile
import os
import re
import json
from dify_plugin.file.file import File
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from docx import Document


class WordChunkTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 获取上传的word文件
        word_content: File = tool_parameters.get("word_content")
        chunk_num: int = tool_parameters.get("chunk_num")

        if not word_content:
            yield self.create_text_message("请提供word文件")
            return
        # 检查文件类型
        if not isinstance(word_content, File):
            yield self.create_text_message("无效的文件格式，期望File对象")
            return

        try:
            # 创建临时文件保存上传的word文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                # 获取文件内容（字节）并写入临时文件
                temp_file.write(word_content.blob)
                temp_file_path = temp_file.name

            # 调用分段函数
            chunks = self.smart_chunk_paragraphs(
                temp_file_path
            )
            
            # 限制分段个数不超过30个
            chunks = self.limit_chunks_to_max(chunks, max_chunks=chunk_num)
            
            # 清理临时文件
            os.unlink(temp_file_path)

            # 返回分段结果
            
            result = {str(i + 1): chunk for i, chunk in enumerate(chunks)}
            # result = [{str(i + 1): chunk} for i, chunk in enumerate(chunks)]
            # result["chunk_num"] = len(chunks)
            
            yield self.create_json_message(result)
    
        except Exception as e:
            yield self.create_text_message(f"处理word文件时出错: {str(e)}")

    def smart_chunk_paragraphs(self, doc_path, min_length=1000):
        """
        智能合并短段落，生成有意义的文本块。
        参数:
            doc_path: Word文档路径
            min_length: 被认为是有独立意义的最小段落长度（字符数）
        返回:
            一个包含合并后文本块的列表
        """
        doc = Document(doc_path)
        chunks = []  # 最终返回的块列表
        current_chunk = []  # 当前正在构建的块（由多个段落组成）
        consecutive_title_count = 0  # 连续标题计数器

        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()

            # 跳过完全空的段落（通常是格式性的换行）
            if not text:
                continue

            # 检查是否是明显的分块标志（如章节标题）
            # 使用统一的标题判断函数
            is_heading = self.is_title(paragraph)

            # 情况1：遇到一个标题
            if is_heading:
                consecutive_title_count += 1  # 增加连续标题计数

                # 如果当前块不为空，且不是连续标题的情况，则保存当前块
                if current_chunk and consecutive_title_count == 1:
                    chunks.append("\n".join(current_chunk))
                    current_chunk = [text]  # 新块以标题开始
                # 如果是连续标题，则添加到当前块
                else:
                    current_chunk.append(text)
            # 情况2：遇到非标题段落
            else:
                # 如果之前有连续标题，则将非标题段落与之前的标题合并
                if consecutive_title_count > 0:
                    current_chunk.append(text)
                    consecutive_title_count = 0  # 重置计数器
                # 情况3：当前段落很短，且当前块不为空 -> 合并到当前块
                elif len(text) < min_length and current_chunk:
                    current_chunk.append(text)
                # 情况4：当前段落很长 -> 它自己可以成为一个有意义的块
                elif len(text) >= min_length and not current_chunk:
                    chunks.append(text)
                # 情况5：当前段落很长，但当前块已有内容 -> 结束当前块，并以此段落开始新块
                elif len(text) >= min_length and current_chunk:
                    chunks.append("\n".join(current_chunk))
                    current_chunk = [text]
                # 情况6：当前段落很短，且当前块为空 -> 开始一个新块（希望后续段落能合并进来）
                else:
                    current_chunk.append(text)

        # 循环结束后，不要忘记最后一个块
        if current_chunk:
            chunks.append("\n".join(current_chunk))

        # 如果设置了min_chunk_size参数，则进行最小chunk大小处理
        # if min_chunk_size is not None:
        #     final_chunks = []
        #     i = 0
        #     while i < len(chunks):
        #         current_chunk = chunks[i]

        #         # 如果当前chunk长度小于min_chunk_size，并且不是最后一个chunk，则与下一个chunk合并
        #         while i < len(chunks) - 1 and len(current_chunk) < min_chunk_size:
        #             i += 1
        #             current_chunk += "\n" + chunks[i]

        #         final_chunks.append(current_chunk)
        #         i += 1

        #     chunks = final_chunks

        return chunks

    def limit_chunks_to_max(self, chunks, max_chunks=30):
        """
        限制分段个数不超过指定数量，如果超过则按数学方法合并相邻段落。
        
        合并策略：
        1. 计算 商 = 段落总数 // max_chunks，余数 = 段落总数 % max_chunks
        2. 如果余数为0，每「商」个段落合并成1个
        3. 如果余数不为0，前「余数」组每「商+1」个段落合并成1个，剩余的每「商」个段落合并成1个
        
        参数:
            chunks: 原始分段列表
            max_chunks: 最大分段个数，默认30
        返回:
            合并后的分段列表（最多 max_chunks 个）
        """
        if len(chunks) <= max_chunks:
            return chunks
        
        total_chunks = len(chunks)
        quotient = total_chunks // max_chunks  # 商
        remainder = total_chunks % max_chunks  # 余数
        
        result_chunks = []
        chunk_index = 0
        
        # 判断是否能整除
        if total_chunks % max_chunks == 0:
            # 能整除：每 quotient 个段落合并成1个
            for i in range(max_chunks):
                merged_chunk_parts = []
                for j in range(quotient):
                    if chunk_index < total_chunks:
                        merged_chunk_parts.append(chunks[chunk_index])
                        chunk_index += 1
                
                if merged_chunk_parts:
                    result_chunks.append("\n".join(merged_chunk_parts))
        else:
            # 不能整除：前 remainder 组每组 quotient+1 个段落，剩余的每组 quotient 个段落
            
            # 处理前 remainder 组，每组 quotient+1 个段落
            for i in range(remainder):
                # 合并 quotient+1 个段落
                merged_chunk_parts = []
                for j in range(quotient + 1):
                    if chunk_index < total_chunks:
                        merged_chunk_parts.append(chunks[chunk_index])
                        chunk_index += 1
                
                if merged_chunk_parts:
                    result_chunks.append("\n".join(merged_chunk_parts))
            
            # 处理剩余的组，每组 quotient 个段落
            remaining_groups = max_chunks - remainder
            for i in range(remaining_groups):
                # 合并 quotient 个段落
                merged_chunk_parts = []
                for j in range(quotient):
                    if chunk_index < total_chunks:
                        merged_chunk_parts.append(chunks[chunk_index])
                        chunk_index += 1
                
                if merged_chunk_parts:
                    result_chunks.append("\n".join(merged_chunk_parts))
        
        return result_chunks


    # 统一的标题判断函数
    def is_title(self, paragraph, median_size=10):
        text = paragraph.text.strip()

        # 排除短文本干扰项
        if any(k in text for k in ["图", "表", "注："]):
            return False

        # 1. 样式名称检测
        if hasattr(paragraph, "style") and paragraph.style:
            style_name = paragraph.style.name.lower()
            if (
                "heading" in style_name
                or "title" in style_name
                or "标题" in style_name
                or "chapter" in style_name
            ):
                return True

        # 2. 正则匹配增强
        patterns = [
            r"^\d+\.\d+",
            r"^\d+\.[\s\t]*",
            r"^[IVXLCDM]+\.",
            r"^第[\u4e00-\u9fa5\d]+[章节]",
            r"^[一二三四五六七八九十百千万\d]+[\.、]",
            r"^[A-Z][A-Z\s]+\b",  # 全大写短语
            r"^第[一二三四五六七八九十百千万\d]+[条]",
        ]
        if any(re.match(p, text) for p in patterns):
            return True

        # 3. 格式特征（多run检测）
        if hasattr(paragraph, "runs") and paragraph.runs:
            # 加粗比例检测
            bold_runs = sum(
                1 for run in paragraph.runs if run.font and run.bold is True
            )
            if bold_runs / len(paragraph.runs) > 0.7:
                return True

            # 字体检测（抽样三个run）
            for run in paragraph.runs[:3]:
                if run.font:
                    # 字体名称检测
                    if run.font.name and (
                        "黑体" in run.font.name or "Hei" in run.font.name
                    ):
                        return True
                    # 动态字体大小检测
                    if run.font.size and run.font.size.pt > median_size * 1.6:
                        return True

        # 4. 文本特征（放宽长度限制）
        if 2 < len(text) < 30:  # 扩展长度上限
            # 增强大写检测（避免单个字母误判）
            if len(text) > 3 and text.isupper():
                return True
            # 章节词检测
            if text.startswith(("摘要", "附录", "参考文献")):
                return True

        return False
