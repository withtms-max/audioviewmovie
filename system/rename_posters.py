import pandas as pd
import requests
import os
import re
import time

INPUT_DATA_FILE = "영화_정보_최종_합본.xlsx"
TARGET_DIR = "영화_포스터_모음"

def clean_filename(filename):
    # Remove characters that are invalid for filenames
    return re.sub(r'[\\/*?:"<>|]', "", str(filename))

print("데이터 읽는 중...")
df = pd.read_excel(INPUT_DATA_FILE)
df = df.drop_duplicates(subset=['영화 제목'], keep='first')

if not os.path.exists(TARGET_DIR):
    os.makedirs(TARGET_DIR)

print(f"총 {len(df)}개의 포스터를 제목별로 저장 시작합니다.")

count = 0
for idx, row in df.iterrows():
    title = str(row['영화 제목'])
    poster_url = str(row['포스터 URL'])
    
    if poster_url.startswith("http"):
        clean_title = clean_filename(title)
        filename = f"{clean_title}.jpg"
        filepath = os.path.join(TARGET_DIR, filename)
        
        # Check if already exists to avoid redundant download
        if os.path.exists(filepath):
            continue
            
        try:
            res = requests.get(poster_url, timeout=10)
            if res.status_code == 200:
                with open(filepath, "wb") as f:
                    f.write(res.content)
                count += 1
                if count % 20 == 0:
                    print(f"진행 중: {count}개 저장 완료...")
        except Exception as e:
            print(f"Error saving {title}: {e}")
        
        time.sleep(0.05) # Subtle delay

print(f"\n✅ 완료! '{TARGET_DIR}' 폴더에 총 {count}개의 이미지가 저장되었습니다.")
