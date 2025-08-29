# app.py
import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import sys
import traceback
import base64

app = Flask(__name__)

# 세션을 사용하기 위한 SECRET_KEY 설정.
# 환경 변수에서 Base64 인코딩된 문자열을 가져와 디코딩합니다.
raw_secret_key_base64 = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_base64_default')
try:
    # Base64 문자열을 바이트로 디코딩합니다.
    app.config['SECRET_KEY'] = base64.urlsafe_b64decode(raw_secret_key_base64)
except Exception as e:
    print(f"[ERROR] Failed to decode FLASK_SECRET_KEY from Base64: {e}", file=sys.stderr)
    print("Please ensure FLASK_SECRET_KEY environment variable is a valid Base64 string.", file=sys.stderr)
    app.config['SECRET_KEY'] = b'fallback_secret_key_if_decoding_fails'

# --- Configuration ---
excel_file_name = 'search.xlsx'
image_file_name = 'search.png'
# 시트 목록: '전주' 시트가 새로 추가되었습니다.
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']

# --- Access Password ---
ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', 'your_secret_password_default')

# --- Data Loading Function ---
def load_data():
    data = []
    excel_file_path = os.path.join(app.root_path, excel_file_name)

    print(f"[DEBUG] Attempting to load Excel file from: {excel_file_path}", file=sys.stderr)

    if not os.path.exists(excel_file_path):
        print(f"[ERROR] Excel file '{excel_file_path}' not found. Please ensure it's in the same directory as app.py.", file=sys.stderr)
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
                    df['품목'] = (df['품'].fillna('') + df['품'].fillna('')).replace('', '알 수 없음')
                elif not found_spaced_품목:
                    df['품목'] = "알 수 없음"
            
            if '품목' in df.columns:
                df['품목'] = df['품목'].fillna(method='ffill')
            
            if '사이즈' in df.columns:
                df['사이즈'] = df['사이즈'].fillna(method='ffill')
            
            if '평량' in df.columns:
                df['평량'] = df['평량'].fillna(method='ffill')

            if '색상 및 패턴' in df.columns:
                df['색상 및 패턴'] = df['색상 및 패턴'].fillna(method='ffill')
            
            df['시트명'] = sheet
            data.append(df)
            print(f"[DEBUG] Sheet '{sheet}' loaded successfully.", file=sys.stderr)
        except Exception as e:
            print(f"[ERROR] Error loading sheet '{sheet}': {e}", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
    
    if data:
        df_combined = pd.concat(data, ignore_index=True)
        print(f"[DEBUG] Combined DataFrame loaded. Total rows: {len(df_combined)}", file=sys.stderr)
        return df_combined
    else:
        print("[ERROR] No Excel sheets loaded successfully.", file=sys.stderr)
        return pd.DataFrame()

# Load all data when the Flask application starts
df_all = load_data()

# --- Seneca Calculation Function ---
def calculate_seneca(page_count, thickness):
    """
    세네카 계산 로직: 페이지수 / 2 * 두께 / 1000
    결과값은 소수점 첫째 자리까지 반올림하여 표시합니다.
    """
    try:
        pc = float(page_count)
        t = float(thickness)
        
        if t == 0:
            return "두께는 0이 될 수 없습니다."

        seneca_result = (pc / 2) * t / 1000
        return f"{seneca_result:,.1f}"
    except ValueError:
        return "유효한 숫자 값을 입력해주세요."
    except Exception as e:
        return f"계산 오류: {e}"

# --- Web Routes ---
@app.route('/', methods=['GET', 'POST'])
def index():
    authenticated = session.get('authenticated', False)
    
    seneca_result = None
    seneca_page_count = ""
    seneca_selected_thickness = ""
    seneca_selected_product_info = ""

    search_results = []
    message = ""
    
    search_keyword = request.form.get('keyword', '').strip() if request.method == 'POST' else request.args.get('keyword', '').strip()
    
    sort_by = request.args.get('sort_by')
    sort_order = request.args.get('sort_order', 'asc')

    if request.method == 'POST':
        if 'password' in request.form:
            entered_password = request.form.get('password')
            if entered_password == ACCESS_PASSWORD:
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
            else:
                flash("세네카 계산을 위한 페이지 수와 유효한 품목 두께를 선택해주세요.", 'info')
                seneca_result = "값 부족 또는 품목 미선택"
            
        pass
    
    if authenticated:
        if df_all.empty:
            message = "로드된 데이터가 없습니다. 검색을 수행할 수 없습니다."
            result_df = pd.DataFrame()
        else:
            if not search_keyword:
                result_df = df_all.copy()
            else:
                if search_keyword in sheets:
                    result_df = df_all[df_all['시트명'].astype(str).str.lower() == search_keyword.lower()].copy()
                elif '품목' not in df_all.columns:
                    message = "'품목' 컬럼을 찾을 수 없습니다. Excel 파일 구조를 확인해주세요."
                    result_df = pd.DataFrame()
                else:
                    result_df = df_all[df_all['품목'].astype(str).str.contains(search_keyword, case=False, na=False)].copy()

            if result_df.empty and not message:
                message = f"'{search_keyword}'에 대한 검색 결과가 없습니다."
            
            if not result_df.empty and sort_by:
                if sort_by in ['평량', '고시가', '두께']:
                    result_df[f'{sort_by}_sortable'] = pd.to_numeric(result_df[sort_by], errors='coerce')
                    result_df = result_df.sort_values(
                        by=f'{sort_by}_sortable',
                        ascending=(sort_order == 'asc'),
                        na_position='last'
                    ).drop(columns=f'{sort_by}_sortable')
                else:
                    result_df = result_df.sort_values(
                        by=sort_by,
                        ascending=(sort_order == 'asc')
                    )

            if not result_df.empty:
                for _, row in result_df.iterrows():
                    formatted_고시가 = 'N/A'
                    original_고시가 = None
                    if '고시가' in row and pd.notna(row['고시가']):
                        try:
                            # 고시가 원본을 저장
                            original_고시가 = str(row['고시가']).replace(',', '').strip()
                            formatted_고시가 = f"{int(float(original_고시가)):,}" if original_고시가 else 'N/A'
                            
                        except (ValueError, TypeError):
                            formatted_고시가 = str(row['고시가'])
                    
                    thickness_value = row.get('두께', 'N/A')

                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': formatted_고시가,
                        '고시가_원본': original_고시가,
                        '두께': thickness_value,
                        '시트명': row.get('시트명', 'N/A')
                    })
            
        logo_path = image_file_name
        return render_template('index.html',
                               authenticated=authenticated,
                               results=search_results,
                               keyword=search_keyword,
                               message=message,
                               logo_path=logo_path,
                               current_sort_by=sort_by,
                               current_sort_order=sort_order,
                               seneca_result=seneca_result,
                               seneca_page_count=seneca_page_count,
                               seneca_selected_thickness=seneca_selected_thickness,
                               seneca_selected_product_info=seneca_selected_product_info)
    else:
        return render_template('index.html', authenticated=False)

# 새로운 API 엔드포인트: 세네카 계산을 위한 비동기 요청 처리
@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    if not session.get('authenticated'):
        return jsonify({'error': 'Unauthorized'}), 401
    
    data = request.get_json()
    page_count = data.get('page_count', '').strip()
    thickness = data.get('thickness', '').strip()

    if not page_count or not thickness or thickness == 'N/A' or thickness == 'None':
        return jsonify({'error': '세네카 계산을 위한 페이지 수와 유효한 품목 두께를 입력해주세요.'}), 400
    
    seneca_result = calculate_seneca(page_count, thickness)
    return jsonify({'result': seneca_result})


if __name__ == '__main__':
    app.run(debug=True)
