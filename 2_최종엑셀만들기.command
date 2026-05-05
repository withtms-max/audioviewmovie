#!/bin/bash
cd "$(dirname "$0")"

echo "========================================"
echo "데이터 병합 및 최종 엑셀을 생성합니다..."
echo "Merging data and creating final Excel..."
echo "========================================"

# Try python3 first, then python
if command -v python3 &>/dev/null; then
    python3 "system/final_format.py"
elif command -v python &>/dev/null; then
    python "system/final_format.py"
else
    echo ""
    echo "[오류] 파이썬을 찾을 수 없습니다."
    echo "파이썬이 설치되어 있는지 확인해주세요."
fi

echo ""
echo "완료되었습니다. 창을 닫으려면 아무 키나 누르세요."
read -n 1 -s -r -p ""
