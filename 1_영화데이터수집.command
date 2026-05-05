#!/bin/bash
cd "$(dirname "$0")"

echo "========================================"
echo "영화 데이터 수집을 시작합니다..."
echo "Starting movie data collection..."
echo "========================================"

# Try python3 first, then python
if command -v python3 &>/dev/null; then
    python3 "system/tmdb_fetch.py"
elif command -v python &>/dev/null; then
    python "system/tmdb_fetch.py"
else
    echo ""
    echo "[오류] 파이썬을 찾을 수 없습니다."
    echo "파이썬이 설치되어 있는지 확인해주세요."
fi

echo ""
echo "완료되었습니다. 창을 닫으려면 아무 키나 누르세요."
read -n 1 -s -r -p ""
