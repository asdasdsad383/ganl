#!/bin/bash
# 论文格式自动转换工具 - 启动脚本

echo "======================================"
echo "论文格式自动转换工具"
echo "======================================"
echo ""

# 检查Python版本
echo "检查Python环境..."
python3 --version

# 检查依赖
echo ""
echo "检查依赖包..."
pip3 show Flask > /dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Flask未安装，正在安装依赖..."
    pip3 install -r requirements.txt
fi

# 创建必要的目录
echo ""
echo "创建必要的目录..."
mkdir -p uploads outputs

# 启动应用
echo ""
echo "======================================"
echo "启动Web服务器..."
echo "======================================"
echo "请在浏览器中访问: http://localhost:5000"
echo "按 Ctrl+C 停止服务器"
echo ""

python3 app.py
