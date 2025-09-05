from typing import Optional, Dict, Union


def get_meta_data(
    mime_type: str, output_filename: Optional[str]
) -> Dict[str, Union[str, None]]:
    """
    生成文件元数据

    Args:
        mime_type: 文件MIME类型
        output_filename: 输出文件名（可选）

    Returns:
        dict: 包含文件元数据的字典
    """
    if not mime_type:
        raise ValueError("Failed to generate meta data, mime_type is not defined")

    # 规范化文件名
    result_filename: Optional[str] = None
    temp_filename = output_filename.strip() if output_filename else None
    if temp_filename:
        # 确保有扩展名
        if (
            mime_type
            == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ):
            extension = ".docx"
            if not temp_filename.lower().endswith(extension):
                temp_filename = f"{temp_filename}{extension}"
        result_filename = temp_filename

    return {
        "mime_type": mime_type,
        "filename": result_filename,
    }


def sanitize_filename(filename: str) -> str:
    """
    清理文件名，移除不允许的字符

    Args:
        filename: 原始文件名

    Returns:
        str: 清理后的文件名
    """
    # Windows不允许的字符
    forbidden_chars = ["<", ">", ":", '"', "/", "\\", "|", "?", "*"]

    clean_filename = filename
    for char in forbidden_chars:
        clean_filename = clean_filename.replace(char, "_")

    # 移除首尾空格和点
    clean_filename = clean_filename.strip(" .")

    # 确保文件名不为空
    if not clean_filename:
        clean_filename = "document"

    # 限制文件名长度（Windows最大255字符）
    if len(clean_filename) > 250:  # 留一些空间给扩展名
        clean_filename = clean_filename[:250]

    return clean_filename
