#!/bin/bash
# Word 格式修改器 — Mac 依赖安装脚本
# 双击此文件安装所需 Python 包

cd "$(dirname "$0")"

echo "=================================================="
echo "  Word 格式修改器 — 安装依赖"
echo "=================================================="
echo ""

# 优先使用项目自带的 venv Python，不存在则用系统 Python
VENV_PYTHON="/Users/zhangyinan/.workbuddy/binaries/python/envs/default/bin/python3"
if [ -f "$VENV_PYTHON" ]; then
    PYTHON="$VENV_PYTHON"
    echo "[检测] 使用虚拟环境 Python"
else
    PYTHON=$(which python3 2>/dev/null)
    if [ -z "$PYTHON" ]; then
        echo "[错误] 未检测到 Python 3，请先安装"
        echo "下载地址：https://www.python.org/downloads/"
        echo ""
        read -p "按回车键退出..."
        exit 1
    fi
    echo "[检测] 使用系统 Python：$PYTHON"
fi

echo ""
echo "正在安装依赖..."
$PYTHON -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

if [ $? -eq 0 ]; then
    echo ""
    echo "=================================================="
    echo "  安装完成！请双击 start.command 启动程序"
    echo "=================================================="
else
    echo ""
    echo "[错误] 安装失败，请检查网络连接后重试"
fi

echo ""
read -p "按回车键退出..."
