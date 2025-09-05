import logging

from dify_plugin.config.logger_format import plugin_logger_handler


def get_logger(name: str) -> logging.Logger:
    """
    获取配置好的日志记录器

    Args:
        name: 日志记录器名称，通常使用 __name__

    Returns:
        logging.Logger: 配置好的日志记录器
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(plugin_logger_handler)
    return logger
