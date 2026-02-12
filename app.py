import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
import numpy as np
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'paper_system_final_v17' # 세션 보안 키

SITE_PASSWORD = "03877"
EXCEL_FILE = 'search.xlsx'
cached_data = []
last_updated = ""

def load_data():
    global cached_data, last_updated
    if not os.path.exists(EXCEL_FILE): 
        print("Excel file not found!")
        return
        
    try:
        # 파일 수정 시간으로 마지막 업데이트 일자 기록
        mtime = os.path.getmtime(EXCEL_FILE)
        last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
        
        # 엑셀의 모든 시트 읽기
        all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, engine='openpyxl')
        combined_list = []
        
        for sheet_name, df in all_sheets.items():
            # 1. 컬럼명 정리 (앞뒤 공백 제거)
            df.columns = [str(col).strip() for col in df.columns]
            
            # [중요 보완] 실제 빈칸이 아닌 '공백 문자'가 있을 경우 NaN으로 변환하여 ffill이 작동하게 함
            df = df.replace(r'^\s*$', np.nan, regex=True)
            
            # 2. 위에서 아래로 데이터 채우기 (ffill)
            df = df.ffill() 
            
            # 3. 필요한 컬럼 매핑 (유사한 이름 찾기)
            col_map = {
                '품목': next((c for c in df.columns if '품목' in c), '품목'),
                '색상': next((c for c in df.columns if '색상' in c or '패턴' in c), '색상'),
                '사이즈': next((c for c in df.columns if '사이즈' in c or '규격' in c), '사이즈'),
                '평량': next((c for c in df.columns if '평량' in c), '평량'),
                '고시가': next((c for c in df.columns if '고시가' in c), '고시가'),
                '두께': next((c for c in df.columns if '두께' in c), '두께')
            }
            
            temp_df = pd.DataFrame()
            
            # 4. 데이터 추출 및 'nan' 방지 처리
            temp_df['품목'] = df[col_map['품목']].fillna('Unknown').astype(str).str.strip() if col_map['품목'] in df.columns else "Unknown"
            temp_df['색상'] = df[col_map['색상']].fillna('-').astype(str).str.strip() if col_map['색상'] in df.columns else "-"
            temp_df['사이즈'] = df[col_map['사이즈']].fillna('-').astype(str).str.strip() if col_map['사이즈'] in df.columns else "-"
            
            # 평량: 숫자로 변환 후 정수형 문자열로 저장 (비어있으면 "0")
            if col_map['평량'] in df.columns:
                temp_df['평량'] = pd.to_numeric(df[col_map['평량']], errors='coerce').fillna(0).astype(int).astype(str)
            else:
                temp_df['평량'] = "0"
                
            # 두께: 숫자로 변환 후 정수형 문자열로 저장 (비어있으면 "0")
            if col_map['두께'] in df.columns:
                temp_df['두께'] = pd.to_numeric(df[col_map['두께']], errors='coerce').fillna(0).astype(int).astype(str)
            else:
                temp_df['두께'] = "0"
            
            # 고시가: 천단위 콤마 처리
            if col_map['고시가'] in df.columns:
                prices = pd.to_numeric(df[col_map['고시가']], errors='coerce').fillna(0)
                temp_df['고시가'] = prices.apply(lambda x: f"{int(x):,}" if x > 0 else "0")
            else:
                temp_df['고시가'] = "0"
                
            temp_df['시트명'] = str(sheet_name).strip()
            combined_list.append(temp_df)
            
        if combined_list:
            # 모든 시트 데이터 합치기
            cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
            
    except Exception as e:
        print(f"Excel Loading Error: {e}")

# 초기 데이터 로드
load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    # 로그인 처리
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == SITE_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            return render_template('index.html', authenticated=False, error="비밀번호가 틀렸습니다.")
            
    if not session.get('authenticated'):
        return render_template('index.html', authenticated=False)

    # 검색 처리
    keyword = request.form.get('keyword', '').strip() if request.method == 'POST' else ""
    results = []
    
    if keyword:
        low_keyword = keyword.lower()
        results = [
            item for item in cached_data 
            if low_keyword in str(item.get('품목', '')).lower() or 
               low_keyword in str(item.get('색상', '')).lower()
        ]
        
        # 제조사별 링크 생성 로직
        for item in results:
            p, s = item['품목'], item['시트명']
            if '두성' in s:
                item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
            elif '삼원' in s:
                item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
            else:
                item['url'] = f"https://www.google.com/search?q={s}+{p}"
    
    return render_template('index.html', results=results, keyword=keyword, authenticated=True, last_updated=last_updated)

if __name__ == '__main__':
    # Render 환경 등을 고려하여 포트 설정
    app.run(host='0.0.0.0', port=5000)