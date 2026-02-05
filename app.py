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

# --- Configuration ---
excel_file_name = 'search.xlsx'
image_file_name = 'search.png'
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']

# --- Access Password ---
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

# --- Data Loading Function ---
def load_data():
    data = []
    excel_file_path = os.path.join(app.root_path, excel_file_name)

    print(f"[DEBUG] Attempting to load Excel file from: {excel_file_path}", file=sys.stderr)

    if not os.path.exists(excel_file_path):
        print(f"[ERROR] Excel file '{excel_file_path}' not found.", file=sys.stderr)
        return pd.DataFrame()

    for sheet in sheets:
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet, engine='openpyxl')
            df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            
            if '품목' not in df.columns:
                found_spaced_품목 = False
                for col in df.columns:
                    if '품 목' in col.replace(' ', ''):
                        df.rename(columns={col: '품목'}, inplace=True)
                        found_spaced_품목 = True
                        break
                
                if not found_spaced_품목 and '품' in df.columns and '목' in df.columns:
                    df['품목'] = (df['품'].fillna('') + df['목'].fillna('')).replace('', '알 수 없음')
                elif not found_spaced_품목:
                    df['품목'] = "알 수 없음"
            
            # --- Pandas 최신 버전 호환성 수정 (중요!) ---
            if '품목' in df.columns:
                df['품목'] = df['품목'].ffill() # 수정 완료
            
            if '사이즈' in df.columns:
                df['사이즈'] = df['사이즈'].ffill() # 수정 완료
            
            if '평량' in df.columns:
                df['평량'] = df['평량'].ffill() # 수정 완료

            if '색상 및 패턴' in df.columns:
                df['색상 및 패턴'] = df['색상 및 패턴'].ffill() # 수정 완료
            # ------------------------------------------

            df['시트명'] = sheet
            data.append(df)
            print(f"[DEBUG] Sheet '{sheet}' loaded successfully.", file=sys.stderr)
        except Exception as e:
            print(f"[ERROR] Error loading sheet '{sheet}': {e}", file=sys.stderr)
    
    if data:
        df_combined = pd.concat(data, ignore_index=True)
        return df_combined
    else:
        return pd.DataFrame()

df_all = load_data()

# --- Seneca Calculation Function ---
def calculate_seneca(page_count, thickness):
    try:
        pc = float(page_count)
        t = float(thickness)
        if t == 0: return "두께는 0이 될 수 없습니다."
        seneca_result = (pc / 2) * t / 1000
        return f"{seneca_result:,.1f}"
    except:
        return "유효한 값을 입력해주세요."

# --- Web Routes ---
@app.route('/', methods=['GET', 'POST'])
def index():
    # 데이터 로드 확인 (앱 시작시 실패했을 경우 대비)
    global df_all
    if df_all.empty:
        df_all = load_data()

    authenticated = session.get('authenticated', False)
    seneca_result, seneca_page_count, seneca_selected_thickness, seneca_selected_product_info = None, "", "", ""
    search_results, message = [], ""
    
    search_keyword = request.form.get('keyword', '').strip() if request.method == 'POST' else request.args.get('keyword', '').strip()
    sort_by = request.args.get('sort_by')
    sort_order = request.args.get('sort_order', 'asc')

    if request.method == 'POST':
        if 'password' in request.form:
            if request.form.get('password') == ACCESS_PASSWORD:
                session['authenticated'] = True
                return redirect(url_for('index'))
            else:
                flash('비밀번호가 틀렸습니다.', 'danger')
                return render_template('index.html', authenticated=False)
        
        if not authenticated:
            return render_template('index.html', authenticated=False)

        if 'calculate_seneca_btn' in request.form:
            seneca_page_count = request.form.get('seneca_page_count', '').strip()
            seneca_selected_thickness = request.form.get('seneca_selected_thickness_hidden', '').strip()
            seneca_selected_product_info = request.form.get('seneca_selected_product_info_hidden', '').strip()
            if seneca_page_count and seneca_selected_thickness and seneca_selected_thickness != 'N/A':
                seneca_result = calculate_seneca(seneca_page_count, seneca_selected_thickness)
    
    if authenticated:
        if df_all.empty:
            message = "데이터를 불러오는 데 실패했습니다. 파일명을 확인하세요."
            result_df = pd.DataFrame()
        else:
            if not search_keyword:
                result_df = df_all.copy()
            else:
                if search_keyword in sheets:
                    result_df = df_all[df_all['시트명'] == search_keyword].copy()
                else:
                    result_df = df_all[df_all['품목'].astype(str).str.contains(search_keyword, case=False, na=False)].copy()

            if result_df.empty:
                message = f"'{search_keyword}'에 대한 결과가 없습니다."
            else:
                # 결과 포맷팅
                for _, row in result_df.iterrows():
                    sheet_name = row.get('시트명')
                    url_to_use = company_urls.get(sheet_name, '#')
                    if sheet_name in ['두성', '삼원', '서경', '삼화'] and search_keyword:
                        url_to_use = f"{company_urls[sheet_name]}{search_keyword}"

                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': row.get('고시가', 'N/A'),
                        '두께': row.get('두께', 'N/A'),
                        '시트명': sheet_name,
                        'url': url_to_use
                    })
            
        return render_template('index.html', authenticated=authenticated, results=search_results, keyword=search_keyword, message=message, logo_path=image_file_name, seneca_result=seneca_result)
    return render_template('index.html', authenticated=False)

@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    if not session.get('authenticated'): return jsonify({'error': 'Unauthorized'}), 401
    data = request.get_json()
    res = calculate_seneca(data.get('page_count', ''), data.get('thickness', ''))
    return jsonify({'result': res})

if __name__ == '__main__':
    app.run(debug=True)
