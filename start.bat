@echo off
chcp 65001 >nul
REM Word 格式修改器 — Windows 双击启动脚本
REM 双击此文件，自动启动本地服务并打开浏览器

cd /d "%~dp0"

REM 杀死占用 8080 端口的旧进程（防止多开）
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8080 " ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a 2>nul
)
timeout /t 1 /nobreak >nul

echo ==================================================
echo   Word 格式修改器 启动中...
echo   浏览器将自动打开：http://localhost:8080
echo   关闭此窗口即退出程序
echo ==================================================

REM 启动服务
start http://localhost:8080
python server.py

pause
