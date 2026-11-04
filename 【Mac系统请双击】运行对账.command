#!/bin/bash
# 账单退款对账 - Mac 一键运行
# 双击此文件即可自动扫描 Excel、生成 HTML 报告并打开

cd "$(dirname "$0")"
echo "=========================================="
echo "📊 账单退款对账工具 (Mac)"
echo "=========================================="
echo ""

# 检查 python3
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 python3，请先安装 Python 3"
    echo "下载地址：https://www.python.org/downloads/"
    read -p "按回车键退出..."
    exit 1
fi

# 检查依赖
python3 -c "import pandas, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "📦 首次运行，正在安装依赖..."
    python3 -m pip install pandas openpyxl -q
fi

echo "🚀 开始处理..."
echo ""

python3 process.py

echo ""
read -p "按回车键关闭窗口..."
