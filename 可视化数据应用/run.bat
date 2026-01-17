@echo off

REM 检查Python是否已安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 未找到Python解释器，请先安装Python 3.9或更高版本
    echo 您可以从以下链接下载Python：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 检查是否已安装必要的依赖项
echo 检查必要的依赖项...
python -m pip list | findstr /i "streamlit pandas plotly matplotlib numpy openpyxl xlsxwriter seaborn" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装必要的依赖项...
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo 依赖项安装失败
        pause
        exit /b 1
    )
)

REM 启动Streamlit应用
echo 正在启动可视化数据应用...
python -m streamlit run app.py --browser.gatherUsageStats false

pause
