#!/bin/bash

# 检查Python是否已安装
echo "检查Python是否已安装..."
if ! command -v python3 &> /dev/null; then
    echo "未找到Python解释器，请先安装Python 3.9或更高版本"
    echo "您可以从以下链接下载Python：https://www.python.org/downloads/"
    read -p "按任意键退出..."
    exit 1
fi

# 使用python3而不是python
PYTHON_CMD="python3"

# 检查是否已安装必要的依赖项
echo "检查必要的依赖项..."
$PYTHON_CMD -m pip list | grep -i -E "streamlit|pandas|plotly|matplotlib|numpy|openpyxl|xlsxwriter|seaborn" > /dev/null 2>&1

if [ $? -ne 0 ]; then
    echo "正在安装必要的依赖项..."
    $PYTHON_CMD -m pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "依赖项安装失败"
        read -p "按任意键退出..."
        exit 1
    fi
fi

# 启动Streamlit应用
echo "正在启动可视化数据应用..."
$PYTHON_CMD -m streamlit run app.py

read -p "按任意键退出..."
