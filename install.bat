@echo off
chcp 65001 >nul
REM Word 格式修改器 — Windows 依赖安装脚本
REM 双击此文件安装所需 Python 包（需已安装 Python 3.9+）

echo ==================================================
echo   Word 格式修改器 — 安装依赖
echo ==================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo 正在安装依赖...
python -m pip install flask python-docx fonttools -i https://pypi.tuna.tsinghua.edu.cn/simple

if %errorlevel% equ 0 (
    echo.
    echo ==================================================
    echo   安装完成！请双击 start.bat 启动程序
    echo ==================================================
) else (
    echo.
    echo [错误] 安装失败，请检查网络连接后重试
)

pause
