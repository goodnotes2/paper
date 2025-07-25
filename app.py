# app.py
import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
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
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔'] 

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
                    df['품목'] = (df['품'].fillna('') + df['품'].fillna('')).replace('', '알 수 없음') # '품'과 '목' 컬럼을 합쳐 '품목' 생성
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
            
            # '두께' 컬럼은 ffill 없이 로드됩니다 (엑셀에 있는 그대로).
            # '두께' 컬럼이 존재한다면, 나중에 검색 결과에 포함됩니다.

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
        pc = float(page_count) # 페이지 수는 소수점도 가능하게 float으로 받습니다.
        t = float(thickness)   # 두께도 소수점 가능하게 float으로 받습니다.
        
        if t == 0: # 두께가 0이면 0으로 나누기 오류 방지
            return "두께는 0이 될 수 없습니다."

        seneca_result = (pc / 2) * t / 1000 # 새로운 계산식 적용
        return f"{seneca_result:,.1f}" # 소수점 첫째 자리까지 콤마 포맷팅 및 반올림
    except ValueError:
        return "유효한 숫자 값을 입력해주세요."
    except Exception as e:
        return f"계산 오류: {e}"

# --- Page Cost Calculation Function (이 함수는 이제 사용되지 않으므로 제거됩니다) ---
# def calculate_page_cost(page_count):
#     """
#     페이지 수 기반 계산 로직을 여기에 구현합니다.
#     아직 검색된 종이 정보와 연동되지 않은 단순 플레이스홀더입니다.
#     """
#     try:
#         pc = int(page_count)
#         # TODO: 실제 페이지 수 기반 계산 공식을 여기에 적용하세요.
#         # 예시: 페이지 수 * 100 (단순 예시)
#         calculated_value = pc * 100
#         return f"{calculated_value:,.0f} 원" # 정수로 콤마 포맷팅
#     except ValueError:
#         return "유효한 페이지 수를 입력해주세요."
#     except Exception as e:
#         return f"계산 오류: {e}"


# --- Web Routes ---
@app.route('/', methods=['GET', 'POST'])
def index():
    authenticated = session.get('authenticated', False)
    
    # 세네카 계산 결과 및 입력값 초기화
    seneca_result = None
    seneca_page_count = "" 
    seneca_selected_thickness = "" # 선택된 두께 값
    seneca_selected_product_info = "" # 선택된 품목 정보 (표시용)

    # 페이지 수 계산 관련 변수 제거
    # page_count_input = ""
    # calculated_page_value = None

    if request.method == 'POST':
        # 비밀번호 제출 처리
        if 'password' in request.form:
            entered_password = request.form.get('password')
            if entered_password == ACCESS_PASSWORD:
                session['authenticated'] = True
                return redirect(url_for('index'))
            else:
                flash('비밀번호가 틀렸습니다.', 'danger')
                return render_template('index.html', authenticated=False)
        
        # 세네카 계산 폼 제출 처리
        elif 'calculate_seneca_btn' in request.form and authenticated:
            seneca_page_count = request.form.get('seneca_page_count', '').strip()
            seneca_selected_thickness = request.form.get('seneca_selected_thickness_hidden', '').strip() # 숨겨진 필드에서 두께 값 가져옴
            seneca_selected_product_info = request.form.get('seneca_selected_product_info_hidden', '').strip() # 숨겨진 필드에서 품목 정보 가져옴
            
            if seneca_page_count and seneca_selected_thickness and seneca_selected_thickness != 'N/A':
                seneca_result = calculate_seneca(seneca_page_count, seneca_selected_thickness)
            else:
                flash("세네카 계산을 위한 페이지 수와 유효한 품목 두께를 선택해주세요.", 'info')
                seneca_result = "값 부족 또는 품목 미선택"
        
        # 페이지 수 계산 폼 제출 처리 로직 제거
        # elif 'calculate_page_btn' in request.form and authenticated:
        #     page_count_input = request.form.get('page_count_input', '').strip()
        #     
        #     if page_count_input:
        #         calculated_page_value = calculate_page_cost(page_count_input)
        #     else:
        #         flash("페이지 수를 입력해주세요.", 'info')
        #         calculated_page_value = "값 부족"

        # 이미 인증된 상태에서 검색 폼 제출 처리 (세네카 계산 폼이 아니면서 POST 요청인 경우)
        elif authenticated: 
            search_results = []
            search_keyword = request.form.get('keyword', '').strip()
            message = ""
            
            sort_by = request.args.get('sort_by')
            sort_order = request.args.get('sort_order', 'asc')

            result_df = df_all.copy()

            if not search_keyword:
                result_df = df_all.copy()
            elif df_all.empty:
                message = "로드된 데이터가 없습니다. 검색을 수행할 수 없습니다."
            else:
                if search_keyword in sheets:
                    result_df = df_all[df_all['시트명'].astype(str).str.lower() == search_keyword.lower()].copy()
                elif '품목' not in df_all.columns:
                    message = "'품목' 컬럼을 찾을 수 없습니다. Excel 파일 구조를 확인해주세요."
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
                    if '고시가' in row and pd.notna(row['고시가']):
                        try:
                            formatted_고시가 = f"{int(row['고시가']):,}"
                        except ValueError:
                            formatted_고시가 = str(row['고시가'])
                    
                    thickness_value = row.get('두께', 'N/A') 

                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': formatted_고시가,
                        '두께': thickness_value, # '두께' 데이터 추가
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
                                   seneca_selected_product_info=seneca_selected_product_info) # page_count_input, calculated_page_value 제거
        else:
            return render_template('index.html', authenticated=False)

    # GET 요청 처리 (초기 로드, 정렬 또는 계산 후 리다이렉트)
    else:
        if not authenticated:
            return render_template('index.html', authenticated=False)
        else:
            search_results = []
            search_keyword = request.args.get('keyword', '').strip()
            message = ""
            
            sort_by = request.args.get('sort_by')
            sort_order = request.args.get('sort_order', 'asc')

            result_df = df_all.copy()

            if not search_keyword:
                result_df = df_all.copy()
            elif df_all.empty:
                message = "로드된 데이터가 없습니다. 검색을 수행할 수 없습니다."
            else:
                if search_keyword in sheets:
                    result_df = df_all[df_all['시트명'].astype(str).str.lower() == search_keyword.lower()].copy()
                elif '품목' not in df_all.columns:
                    message = "'품목' 컬럼을 찾을 수 없습니다. Excel 파일 구조를 확인해주세요."
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
                    if '고시가' in row and pd.notna(row['고시가']):
                        try:
                            formatted_고시가 = f"{int(row['고시가']):,}"
                        except ValueError:
                            formatted_고시가 = str(row['고시가'])
                    
                    thickness_value = row.get('두께', 'N/A')

                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': formatted_고시가,
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
                                   seneca_selected_product_info=seneca_selected_product_info) # page_count_input, calculated_page_value 제거

if __name__ == '__main__':
    app.run(debug=True)
