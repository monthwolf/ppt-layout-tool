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
set ENV_NAME=ppt_tool

REM 检查环境是否存在
call conda env list | findstr /C:"%ENV_NAME%" >nul
if %ERRORLEVEL% NEQ 0 (
    echo 正在创建新的Conda环境: %ENV_NAME%...
    call conda create -n %ENV_NAME% python=3.9 -y
) else (
    echo Conda环境 %ENV_NAME% 已存在
    choice /C YN /M "是否删除并重新创建环境？(Y/N)"
    if errorlevel 2 (
        echo 继续使用现有环境...
    ) else (
        echo 正在删除并重新创建环境 %ENV_NAME%...
        call conda env remove -n %ENV_NAME% -y
        call conda create -n %ENV_NAME% python=3.9 -y
    )
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

REM 安装依赖
echo 正在安装必要依赖...
call pip install -r requirements.txt

REM 检查安装是否成功
call python -c "from PyQt6.QtWidgets import QApplication" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo PyQt6安装可能存在问题，尝试修复...
    echo 安装Visual C++ Redistributable可能有助于解决这个问题
    call pip uninstall -y PyQt6 PyQt6-Qt6 PyQt6-sip pyqt6-tools
    echo 重新安装PyQt6组件...
    call pip install PyQt6==6.4.2 pyqt6-tools==6.4.2.3.3
)

REM 检查PowerPoint是否可用（用于处理旧版PPT）
echo 检查PowerPoint支持...
call python -c "import comtypes.client; comtypes.client.CreateObject('PowerPoint.Application')" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo 注意: 未检测到PowerPoint，将使用模拟方式处理PPT文件
    echo 如需完整功能，请确保已安装Microsoft PowerPoint
    echo.
)

echo 支持的PPT文件格式:
echo - PPTX格式: 使用python-pptx直接读取
echo - PPT格式(旧版): 将尝试使用PowerPoint COM接口处理
echo.

REM 启动应用程序
echo 启动PPT处理工具...
python src/main.py
if %ERRORLEVEL% NEQ 0 (
    echo 程序运行出错，错误代码: %ERRORLEVEL%
    echo 请检查日志获取更多信息
    echo 尝试启动备用版本...
    python src/main_pyside.py
    if %ERRORLEVEL% NEQ 0 (
        echo 备用版本也无法启动，请参考README中的故障排除说明
        pause
        exit /b
    )
)

REM 在应用程序结束后保持窗口
pause 