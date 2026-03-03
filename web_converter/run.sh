#!/bin/bash
cd "$(dirname "$0")"
echo "表格转手册 - 启动中..."
pip3 install -q -r requirements.txt 2>/dev/null || true
python3 app.py
