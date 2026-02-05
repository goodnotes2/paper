import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import sys
import traceback
import base64

app = Flask(__name__)

# 세션을 사용하기 위한 SECRET_KEY 설정.
raw_secret_key_base64 = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_base64_default')
try:
    app.config['SECRET_KEY'] = base64.urlsafe_b64decode(raw_secret_key_base64)
except Exception as e:
    print(f"[ERROR] Failed to decode FLASK_SECRET_KEY from Base64: {e}", file=sys.stderr)
    app.config['SECRET_KEY'] = b'fallback_secret_key_if_decoding_fails'

# --- 설정 (Configuration) ---
excel_file_name = 'search.xlsx'
image_file_name = 'search.png'
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']

# --- 접속 암호 (Environment Variable에서 가져옴) ---
ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', 'your_secret_password_default')

# 제지사별 홈페이지 URL
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

# --- 데이터 로드 함수 (빈 칸 자동 채우기 포함) ---
def load_data():
    data = []
    excel_file_path = os.path.join(app.root_path, excel_file_name)

    if not os.path.exists(excel_file_path):
        print(f"[ERROR] '{excel_file_path}' 파일을 찾을 수 없습니다.", file=sys.stderr)
        return pd.DataFrame()

    for sheet in sheets:
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet, engine='openpyxl')
            
            # 컬럼명 공백 제거
            df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            
            # '품목' 컬럼 찾기 및 이름 통일
            if '품목' not in df.columns:
                for col in df.columns:
                    if '품 목' in col.replace(' ', ''):
                        df.rename(columns={col: '품목'}, inplace=True)
                        break
            
            # --- 중요: 빈 칸(NaN) 위에서 아래로 자동 채우기 (.ffill) ---
            # 엑셀에서 첫 줄에만 값을 적고 아래는 비워둔 경우를 처리합니다.
            if '품목' in df.columns:
                df['품목'] = df['품목'].ffill()
            
            if '사이즈' in df.columns:
                df['사이즈'] = df['사이즈'].ffill()
            
            if '평량' in df.columns:
                df['평량'] = df['평량'].ffill()

            if '색상 및 패턴' in df.columns:
                df['색상 및 패턴'] = df['색상 및 패턴'].ffill()
            # ------------------------------------------------------

            df['시트명'] = sheet
            data.append(df)
            print(f"[DEBUG] '{sheet}' 시트 로드 성공.", file=sys.stderr)
        except Exception as e:
            print(f"[ERROR] '{sheet}' 시트 로드 중 오류: {e}", file=sys.stderr)
    
    if data:
        return pd.concat(data, ignore_index=True)
    return pd.DataFrame()

# 서버 시작 시 데이터 로드
df_all = load_data()

# --- 세네카 계산 로직 ---
def calculate_seneca(page_count, thickness):
    try:
        pc = float(page_count)
        t = float(thickness)
        if t <= 0: return "두께 오류"
        result = (pc / 2) * t / 1000
        return f"{result:,.1f}"
    except:
        return "숫자 입력 필요"

# --- 메인 페이지 루틴 ---
@app.route('/', methods=['GET', 'POST'])
def index():
    global df_all
    if df_all.empty: # 데이터가 없으면 다시 로드 시도
        df_all = load_data()

    authenticated = session.get('authenticated', False)
    search_keyword = (request.form.get('keyword', '') if request.method == 'POST' else request.args.get('keyword', '')).strip()
    
    seneca_result = None
    search_results = []
    message = ""

    # 1. 로그인 처리
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == ACCESS_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            flash('비밀번호가 틀렸습니다.', 'danger')
            return render_template('index.html', authenticated=False)

    # 2. 로그인된 상태에서의 검색 및 계산
    if authenticated:
        # 세네카 계산 요청인 경우
        if request.method == 'POST' and 'calculate_seneca_btn' in request.form:
            pc = request.form.get('seneca_page_count', '').strip()
            tk = request.form.get('seneca_selected_thickness_hidden', '').strip()
            if pc and tk and tk != 'nan' and tk != 'N/A':
                seneca_result = calculate_seneca(pc, tk)

        # 검색 처리
        if df_all.empty:
            message = "파일에 데이터가 없거나 불러올 수 없습니다."
        else:
            if not search_keyword:
                res_df = df_all.head(100).copy() # 너무 많으면 처음 100개만
            elif search_keyword in sheets:
                res_df = df_all[df_all['시트명'] == search_keyword].copy()
            else:
                res_df = df_all[df_all['품목'].astype(str).str.contains(search_keyword, case=False, na=False)].copy()

            if res_df.empty:
                message = f"'{search_keyword}'에 대한 결과가 없습니다."
            else:
                for _, row in res_df.iterrows():
                    sheet_name = row.get('시트명')
                    base_url = company_urls.get(sheet_name, '#')
                    # 검색 연동 URL 생성
                    final_url = f"{base_url}{search_keyword}" if sheet_name in ['두성', '삼원', '서경', '삼화'] else base_url

                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': f"{int(float(row['고시가'])):,}" if pd.notna(row.get('고시가')) and str(row.get('고시가')).replace('.','').isdigit() else row.get('고시가', 'N/A'),
                        '두께': row.get('두께', 'N/A'),
                        '시트명': sheet_name,
                        'url': final_url
                    })

        return render_template('index.html', authenticated=True, results=search_results, keyword=search_keyword, message=message, seneca_result=seneca_result)

    return render_template('index.html', authenticated=False)

# API: 비동기 계산용
@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    if not session.get('authenticated'): return jsonify({'error': 'Unauthorized'}), 401
    data = request.get_json()
    res = calculate_seneca(data.get('page_count', ''), data.get('thickness', ''))
    return jsonify({'result': res})

if __name__ == '__main__':
    app.run(debug=True)
