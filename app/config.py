"""应用配置

从环境变量读取配置，避免硬编码在源码中。
使用方法：复制 .env.example 为 .env，按需修改变量值。
"""
import os


class Config:
    """基础配置类"""

    # 服务端口（默认 8765，与原 testcase_viewer.py 保持一致）
    PORT = int(os.environ.get('PORT', 8765))

    # 数据目录（存放 Excel 测试用例文件）
    DATA_DIR = os.environ.get('DATA_DIR', 'testcases')

    # Flask 调试模式
    DEBUG = os.environ.get('FLASK_DEBUG', '0') == '1'


class DevelopmentConfig(Config):
    """开发环境配置"""

    DEBUG = True


class ProductionConfig(Config):
    """生产环境配置"""

    DEBUG = False


# 配置名称映射，供 create_app() 使用
config = {
    'default': Config,
    'development': DevelopmentConfig,
    'production': ProductionConfig,
}
