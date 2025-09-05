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
from pdf2docx import Converter
from tools.utils.logger_utils import get_logger
from tools.utils.file_utils import get_meta_data, sanitize_filename


class PdfToWordTool(Tool):
    # 获取当前模块的日志记录器
    logger = get_logger(__name__)

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 获取上传的PDF文件和自定义文件名
        pdf_content: File = tool_parameters.get("pdf_content")
        custom_filename = tool_parameters.get("output_filename", "").strip()

        if not pdf_content:
            yield self.create_text_message("请提供PDF文件")
            return
        # 检查文件类型
        if not isinstance(pdf_content, File):
            yield self.create_text_message("无效的文件格式，期望File对象")
            return

        try:
            self.logger.info(
                f"开始处理PDF转Word，文件名: {pdf_content.filename if pdf_content.filename else '未知'}，自定义文件名: {custom_filename if custom_filename else '未设置'}"
            )
            # 创建临时文件保存上传的PDF文件
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".pdf"
            ) as temp_pdf_file:
                # 获取文件内容（字节）并写入临时文件
                temp_pdf_file.write(pdf_content.blob)
                temp_pdf_path = temp_pdf_file.name

            # 创建临时文件保存转换后的Word文件
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".docx"
            ) as temp_docx_file:
                temp_docx_path = temp_docx_file.name

            # 调用PDF转Word函数
            self.logger.info("开始执行PDF到DOCX的转换")
            self.pdf_to_docx(temp_pdf_path, temp_docx_path)
            self.logger.info("PDF转换完成")

            # 读取转换后的Word文件内容
            with open(temp_docx_path, "rb") as docx_file:
                docx_blob = docx_file.read()

            # 处理输出文件名
            if custom_filename:
                # 清理并处理自定义文件名
                output_filename = sanitize_filename(custom_filename)
                self.logger.info(f"使用自定义文件名: {output_filename}")
            else:
                # 使用原始PDF文件名的基本名称（去掉扩展名）
                if pdf_content.filename:
                    base_name = os.path.splitext(pdf_content.filename)[0]
                    output_filename = sanitize_filename(base_name)
                else:
                    output_filename = "converted"
                self.logger.info(f"使用默认文件名: {output_filename}")

            # 使用create_blob_message返回转换后的Word文件
            yield self.create_blob_message(
                blob=docx_blob,
                meta=get_meta_data(
                    mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    output_filename=output_filename,
                ),
            )

            # 清理临时文件
            os.unlink(temp_pdf_path)
            os.unlink(temp_docx_path)
            self.logger.info("PDF转Word处理完成，临时文件已清理")

        except Exception as e:
            self.logger.exception("处理PDF文件时发生异常")
            yield self.create_text_message(f"处理PDF文件时出错: {str(e)}")

    def pdf_to_docx(self, pdf_path, docx_path):
        # 创建转换器对象
        cv = Converter(pdf_path)
        # 转换整个PDF（start=起始页，end=结束页，None表示全部）
        cv.convert(docx_path, start=0, end=None)
        # 关闭转换器释放资源
        cv.close()
