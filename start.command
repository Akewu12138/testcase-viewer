#!/bin/bash
# ============================================
#  测试用例记录表 - Mac 启动脚本
#  双击此文件即可运行（首次需要右键→打开）
# ============================================

cd "$(dirname "$0")"

echo "=========================================="
echo "  🔬 测试用例记录表  v1.1"
echo "=========================================="
echo ""

# 检查 Python 3
if ! command -v python3 &> /dev/null; then
    echo "❌ 未检测到 Python 3，请先安装"
    echo "   下载地址: https://www.python.org/downloads/"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

echo "📦 Python: $(python3 --version)"
echo "📂 工作目录: $(pwd)"
echo ""

# 安装依赖
echo "🔍 检查依赖..."
python3 -m pip install --quiet flask openpyxl 2>/dev/null

# 启动主程序
python3 testcase_viewer.py

# 退出后等待
echo ""
read -p "按回车键关闭..."
