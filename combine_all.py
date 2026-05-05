import pandas as pd
from openpyxl import load_workbook
import os

# Files
API_FILE = "영화정보_입력양식_최종_1773419337.xlsx" # This had 273 with info, 41 without
MANUAL_PROCESSED_FILE = "추가영화_정리_완료.xlsx" # This has the 41 manual entries
TEMPLATE_FILE = "영화정보_입력양식.xlsx"
FINAL_MASTER_FILE = "영화정보_최종_전체합본.xlsx"

print("최종 합본 파일 생성 중...")

# 1. Load data
df_api = pd.read_excel(API_FILE)
df_manual = pd.read_excel(MANUAL_PROCESSED_FILE)

# In API_FILE, the first 273 rows are usually the ones with info. 
# We'll filter out the "API 정보 없음" ones and replace them with manual ones.
df_api_clean = df_api[df_api['줄거리'] != "API 정보 없음"].copy()

# 2. Combine
df_final = pd.concat([df_api_clean, df_manual], ignore_index=True)

# 3. Save to template format to preserve styling/images (conceptually)
# Since pandas.to_excel doesn't preserve images well, we'll just inform the user 
# that both files are available, but I'll try to make a clean combined text-data file first.
df_final.to_excel(FINAL_MASTER_FILE, index=False)

print(f"✅ 합본 완료! 총 {len(df_final)}편의 데이터가 '{FINAL_MASTER_FILE}'에 저장되었습니다.")
