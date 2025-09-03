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


class PdfToWordTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 获取上传的PDF文件
        pdf_content: File = tool_parameters.get("pdf_content")

        if not pdf_content:
            yield self.create_text_message("请提供PDF文件")
            return
        # 检查文件类型
        if not isinstance(pdf_content, File):
            yield self.create_text_message("无效的文件格式，期望File对象")
            return

        try:
            # 创建临时文件保存上传的PDF文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf_file:
                # 获取文件内容（字节）并写入临时文件
                temp_pdf_file.write(pdf_content.blob)
                temp_pdf_path = temp_pdf_file.name
            
            # 创建临时文件保存转换后的Word文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx_file:
                temp_docx_path = temp_docx_file.name
            
            # 调用PDF转Word函数
            self.pdf_to_docx(temp_pdf_path, temp_docx_path)
            
            # 读取转换后的Word文件内容
            with open(temp_docx_path, 'rb') as docx_file:
                docx_blob = docx_file.read()
            
            # 使用create_blob_message返回转换后的Word文件
            yield self.create_blob_message(
                blob=docx_blob,
                meta={
                    "file_name": f"{os.path.splitext(pdf_content.filename)[0]}.docx" if pdf_content.filename else "converted.docx",
                    "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
            )
            
            # 清理临时文件
            os.unlink(temp_pdf_path)
            os.unlink(temp_docx_path)
            
        except Exception as e:
            yield self.create_text_message(f"处理PDF文件时出错: {str(e)}")
    
    def pdf_to_docx(self, pdf_path, docx_path):
        # 创建转换器对象
        cv = Converter(pdf_path)
        # 转换整个PDF（start=起始页，end=结束页，None表示全部）
        cv.convert(docx_path, start=0, end=None)
        # 关闭转换器释放资源
        cv.close()





