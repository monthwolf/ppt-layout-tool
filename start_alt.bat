@echo off
echo 正在检查Conda环境...

REM 检查conda是否已安装
call conda --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo 未检测到Conda环境，请先安装Anaconda或Miniconda
    pause
    exit /b
)

REM 环境名称
set ENV_NAME=ppt_tool_alt

REM 检查环境是否存在
call conda env list | findstr /C:"%ENV_NAME%" >nul
if %ERRORLEVEL% NEQ 0 (
    echo 正在创建新的Conda环境: %ENV_NAME%...
    call conda create -n %ENV_NAME% python=3.9 -y
) else (
    echo Conda环境 %ENV_NAME% 已存在
)

REM 激活环境
call conda activate %ENV_NAME%
if %ERRORLEVEL% NEQ 0 (
    echo 环境激活失败，请检查Conda安装
    pause
    exit /b
)

REM 确保pip是最新版本
echo 更新pip...
call python -m pip install --upgrade pip

REM 安装基本依赖
echo 安装基本依赖...
call pip install python-pptx==0.6.21 Pillow==9.5.0 reportlab==3.6.12

REM 安装PySide6而不是PyQt6
echo 安装PySide6库...
call pip install PySide6==6.4.2

REM 启动备用版应用程序
echo 启动PPT处理工具(兼容版)...
python src/main_pyside.py
if %ERRORLEVEL% NEQ 0 (
    echo 程序运行出错，错误代码: %ERRORLEVEL%
    echo 请检查日志获取更多信息
    pause
    exit /b
)

REM 在应用程序结束后保持窗口
pause 