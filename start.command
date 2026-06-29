#!/bin/bash
# Word 格式修改器 — Mac 双击启动脚本
# 双击此文件，自动启动本地服务并打开浏览器

cd "$(dirname "$0")"

# 杀死占用 8080 端口的旧进程（防止多开）
lsof -ti:8080 2>/dev/null | xargs kill -9 2>/dev/null
sleep 1

echo "=================================================="
echo "  Word 格式修改器 启动中..."
echo "  浏览器将自动打开：http://localhost:8080"
echo "  关闭此窗口即退出程序"
echo "=================================================="

# 启动服务（使用 venv Python，依赖已安装）
/Users/zhangyinan/.workbuddy/binaries/python/envs/default/bin/python3 server.py
