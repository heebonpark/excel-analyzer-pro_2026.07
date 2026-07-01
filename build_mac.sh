#!/bin/bash
# Excel PRO v5.1 - Mac .app 빌드
echo ""
echo "================================================"
echo "  Excel PRO - Mac .app 빌드"
echo "================================================"

echo "[1/3] PyInstaller 설치..."
pip3 install pyinstaller --quiet

echo "[2/3] 패키지 확인..."
pip3 install pandas openpyxl xlrd matplotlib xlwings --quiet

echo "[3/3] .app 빌드 중 (3~8분 소요)..."

pyinstaller \
    --onefile \
    --windowed \
    --name "Excel분석기PRO" \
    --hidden-import pandas \
    --hidden-import openpyxl \
    --hidden-import xlrd \
    --hidden-import matplotlib \
    --hidden-import matplotlib.backends.backend_tkagg \
    --hidden-import requests \
    --hidden-import urllib3 \
    --hidden-import xlwings \
    --hidden-import numpy \
    --hidden-import tkinter \
    --hidden-import tkinter.ttk \
    --hidden-import tkinter.filedialog \
    --hidden-import tkinter.messagebox \
    --hidden-import tkinter.scrolledtext \
    --hidden-import tkinter.simpledialog \
    --collect-all matplotlib \
    --collect-all pandas \
    --noconfirm \
    Excel_Analyzer_PRO.py

if [ -f "dist/Excel분석기PRO" ]; then
    echo ""
    echo "================================================"
    echo "  빌드 성공!"
    echo "  위치: dist/Excel분석기PRO"
    echo "================================================"
    open dist/
else
    echo "[ERROR] 빌드 실패 — 오류 메시지를 확인하세요"
fi
