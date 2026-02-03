# app.py
import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import sys
import traceback
import base64

app = Flask(__name__)

# SECRET_KEY 설정 (기존 그대로)
raw_secret_key_base64 = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_base64_default')
try:
    app.config['SECRET_KEY'] = base64.urlsafe_b64decode(raw_secret_key_base64)
except:
    app.config['SECRET_KEY'] = b'fallback_secret_key_if_decoding_fails'

excel_file_name = 'search.xlsx'  # 확인됨
image_file_name = 'search.png'
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']

ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', 'your_secret_password_default')
company_urls = {
    '두성': 'https://www.doosungpaper.co.kr/goods/goods_search.php?keyword=',
    '삼원': 'https://www.samwonpaper.com/product/paper/list?search.searchString=',
    '한국': 'https://www.hankukpaper.com/ko/product/listProductinfo.do',
    '무림': 'https://www.moorim.co.kr:13002/product/paper.php',
    '삼화': 'https://www.samwhapaper.com/product/samwhapaper/all?keyword=',
    '서경': 'https://wedesignpaper.com/search?type=shopping&sort=consensus_desc&keyword=',
    '한솔': 'https://www.hansolpaper.co.kr/product/insper',
    '전주': 'https://jeonjupaper.com/publicationpaper'
}

def load_data():
    data = []
    
    # Render 호환 절대경로
    excel_paths = [
        'search.xlsx',  # 루트
        os.path.join(os.path.dirname(__file__), 'search.xlsx'),  # app.py 옆
        os.path.join(os.path.dirname(__file__), 'data', 'search.xlsx')  # data 폴더
    ]
    
    excel_path = None
    for path in excel_paths:
        if os.path.exists(path):
            excel_path = path
            break
    
    if not excel_path:
        print("[ERROR] search.xlsx 파일을 찾을 수 없습니다. 루트 디렉토리에 배치해주세요.", file=sys.stderr)
        return pd.DataFrame()

    print(f"[DEBUG] Loading from: {excel_path}", file=sys.stderr)

    for sheet in sheets:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet, engine='openpyxl')
            df.columns = df.columns.str.strip()
            
            # 컬럼 정리 (기존 로직 유지)
            if '품목' not in df.columns:
                for col in df.columns:
                    if '품목' in str(col):
                        df.rename(columns={col: '품목'}, inplace=True)
                        break
            
            # ✅ pandas 호환 수정
            if '품목' in df.columns:
                df['품목'] = df['품목'].ffill()
            if '사이즈' in df.columns:
                df['사이즈'] = df['사이즈'].ffill()
            if '평량' in df.columns:
                df['평량'] = df['평량'].ffill()
            if '색상 및 패턴' in df.columns:
                df['색상 및 패턴'] = df['색상 및 패턴'].ffill()
            
            df['시트명'] = sheet
            data.append(df)
            print(f"[SUCCESS] Sheet '{sheet}' loaded: {len(df)} rows", file=sys.stderr)
            
        except Exception as e:
            print(f"[ERROR] Sheet '{sheet}': {e}", file=sys.stderr)
    
    if data:
        return pd.concat(data, ignore_index=True)
    return pd.DataFrame()

# 데이터 로드
df_all = load_data()

# calculate_seneca, index, API 라우트는 기존 그대로 유지...
# (나머지 코드는 동일하니 생략)

if __name__ == '__main__':
    app.run(debug=True)
