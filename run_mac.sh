#!/bin/bash
# ============================================
#   Excel 분석기 PRO v5.1 - Mac 실행 스크립트
# ============================================
cd "$(dirname "$0")"

echo ""
echo "============================================"
echo "  Excel 분석기 PRO v5.1"
echo "============================================"
echo ""

# ── Python 선택 전략 ─────────────────────────
# 1순위: /usr/local/bin/python3.11  (python.org 3.11.x, Tcl/Tk 8.6 번들 → 안정)
# 2순위: /usr/local/bin/python3.12  (python.org 3.12.x)
# 3순위: Homebrew python3.11 / python3.13
# 건너뜀: Homebrew python3.12 (Tcl/Tk 9.0 링크 → 크래시)
#          python3.14 (Tcl/Tk 9.0 → 크래시)
PY=""
_try_py() {
    local bin="$1"
    command -v "$bin" &>/dev/null || return 1
    local ver
    ver=$("$bin" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null)
    [ -z "$ver" ] && return 1
    # tkinter 동작 확인 (실패 시 skip)
    "$bin" -c "import tkinter; tkinter.Tk().destroy()" 2>/dev/null || return 1
    PY="$bin"; VER="$ver"
    return 0
}

# python.org 설치본 우선
_try_py /usr/local/bin/python3.11 ||
_try_py /usr/local/bin/python3.12 ||
_try_py /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 ||
_try_py /Library/Frameworks/Python.framework/Versions/3.12/bin/python3 ||
_try_py python3.11 || _try_py python3.13 || _try_py python3

if [ -z "$PY" ]; then
    echo "[오류] 사용 가능한 Python 없음"
    echo "  해결: https://www.python.org/downloads/macos/ 에서 Python 3.11 설치"
    exit 1
fi
echo "Python: $PY ($VER)"
echo "tkinter: OK"

# ── 패키지 설치 ──────────────────────────────
echo ""
echo "[1/2] 패키지 확인 중..."
for pkg in pandas openpyxl xlrd matplotlib xlwings requests urllib3; do
    "$PY" -c "import ${pkg//-/_}" 2>/dev/null && continue
    echo "  설치: $pkg"
    "$PY" -m pip install "$pkg" -q 2>/dev/null \
    || "$PY" -m pip install "$pkg" -q --break-system-packages 2>/dev/null \
    || "$PY" -m pip install "$pkg" -q --user 2>/dev/null \
    || echo "  [경고] $pkg 설치 실패 — 계속 시도"
done

echo "[2/2] 시작..."
echo ""
export TK_SILENCE_DEPRECATION=1

"$PY" "$(dirname "$0")/Excel_Analyzer_PRO.py"

if [ $? -ne 0 ]; then
    echo ""
    echo "=========================================="
    echo " 실행 실패 시 해결 방법:"
    echo ""
    echo " 권장: python.org에서 Python 3.11 설치"
    echo "   https://www.python.org/downloads/macos/"
    echo "   설치 후: /usr/local/bin/python3.11 Excel_Analyzer_PRO.py"
    echo "=========================================="
    read -p "종료하려면 Enter..."
fi
