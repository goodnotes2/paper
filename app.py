import pandas as pd
import numpy as np
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import sys
import traceback
import base64

app = Flask(__name__)

# 세션 보안키 설정
raw_secret_key_base64 = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_base64_default')
try:
    app.config['SECRET_KEY'] = base64.urlsafe_b64decode(raw_secret_key_base64)
except Exception as e:
    app.config['SECRET_KEY'] = b'fallback_secret_key_if_decoding_fails'

# --- 설정 ---
excel_file_name = 'search.xlsx'
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

# --- 데이터 로드 함수 (평량 빈칸 해결 버전) ---
def load_data():
    data = []
    excel_file_path = os.path.join(app.root_path, excel_file_name)

    if not os.path.exists(excel_file_path):
        return pd.DataFrame()

    for sheet in sheets:
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet, engine='openpyxl')
            
            # 1. 컬럼명 공백 제거 및 표준화
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
            
            # 2. 필수 컬럼 이름 통일
            name_map = {'품목': '품목', '사이즈': '사이즈', '평량': '평량', '색상': '색상 및 패턴'}
            for col in df.columns:
                clean_col = col.replace(' ', '')
                for key, val in name_map.items():
                    if key in clean_col:
                        df.rename(columns={col: val}, inplace=True)

            # 3. 평량 컬럼 특수 처리: 숫자가 아닌 값(점선 등)을 빈칸(NaN)으로 변환
            if '평량' in df.columns:
                df['평량'] = pd.to_numeric(df['평량'], errors='coerce')
                df['평량'] = df['평량'].ffill()

            # 4. 나머지 컬럼 자동 채우기
            for target in ['품목', '사이즈', '색상 및 패턴']:
                if target in df.columns:
                    df[target] = df[target].ffill()

            df['시트명'] = sheet
            data.append(df)
        except Exception as e:
            print(f"Error loading sheet {sheet}: {e}", file=sys.stderr)
    
    return pd.concat(data, ignore_index=True) if data else pd.DataFrame()

df_all = load_data()

def calculate_seneca(page_count, thickness):
    try:
        pc, t = float(page_count), float(thickness)
        if t <= 0: return "두께 오류"
        return f"{(pc / 2) * t / 1000:,.1f}"
    except:
        return "입력 오류"

@app.route('/', methods=['GET', 'POST'])
def index():
    global df_all
    if df_all.empty: df_all = load_data()

    authenticated = session.get('authenticated', False)
    search_keyword = (request.form.get('keyword', '') if request.method == 'POST' else request.args.get('keyword', '')).strip()
    
    seneca_result = None
    search_results = []
    message = ""

    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == ACCESS_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            flash('비밀번호가 틀렸습니다.', 'danger')
            return render_template('index.html', authenticated=False)

    if authenticated:
        if request.method == 'POST' and 'calculate_seneca_btn' in request.form:
            pc = request.form.get('seneca_page_count', '').strip()
            tk = request.form.get('seneca_selected_thickness_hidden', '').strip()
            if pc and tk and tk != 'nan': seneca_result = calculate_seneca(pc, tk)

        if not df_all.empty:
            if not search_keyword:
                res_df = df_all.head(50).copy()
            elif search_keyword in sheets:
                res_df = df_all[df_all['시트명'] == search_keyword].copy()
            else:
                res_df = df_all[df_all['품목'].astype(str).str.contains(search_keyword, case=False, na=False)].copy()

            for _, row in res_df.iterrows():
                sheet_name = row.get('시트명')
                search_results.append({
                    '품목': row.get('품목', 'N/A'),
                    '사이즈': row.get('사이즈', 'N/A'),
                    '평량': row.get('평량', 'N/A'),
                    '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                    '고시가': f"{int(float(row['고시가'])):,}" if pd.notna(row.get('고시가')) and str(row.get('고시가')).replace('.','').isdigit() else 'N/A',
                    '두께': row.get('두께', 'N/A'),
                    '시트명': sheet_name,
                    'url': f"{company_urls.get(sheet_name, '#')}{search_keyword}" if sheet_name in ['두성','삼원','서경','삼화'] else company_urls.get(sheet_name, '#')
                })
        return render_template('index.html', authenticated=True, results=search_results, keyword=search_keyword, message=message, seneca_result=seneca_result)

    return render_template('index.html', authenticated=False)

if __name__ == '__main__':
    app.run(debug=True)