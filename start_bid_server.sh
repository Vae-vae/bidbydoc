#!/bin/bash
# 直接使用 conda 环境的 Python（无需 source conda）
PYTHON_PATH="/root/miniconda3/bin/python"   # ← 请修改为上面 which python 的结果

cd /home/bid
exec "$PYTHON_PATH" bid_api_server.py