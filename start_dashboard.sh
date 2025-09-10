#!/bin/bash
# 启动香港疫情数据可视化大屏

echo "🏥 启动香港疫情数据可视化大屏..."
echo "=================================="

# 激活虚拟环境
source venv/bin/activate

# 启动Flask应用
echo "正在启动Flask应用..."
python app.py
