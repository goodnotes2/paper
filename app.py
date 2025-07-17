# app.py
import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
import sys
import traceback

app = Flask(__name__)
# 세션을 사용하기 위한 SECRET_KEY 설정. 환경 변수에서 가져오거나 기본값 사용.
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_here_for_session_default')

# --- Configuration ---
excel_file_name = 'search.xlsx'
image_file_name = 'search.png' # search.png 파일을 'static' 폴더 안에 넣어주세요.
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔']

# --- Access Password ---
# 웹 앱에 접근하기 위한 단일 비밀번호를 환경 변수에서 가져오거나 기본값 사용.
ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', 'your_secret_password_default') # 여기에 기본 비밀번호를 설정하세요.

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
                    df['품목'] = (df['품'].fillna('') + df['목'].fillna('')).replace('', '알 수 없음')
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

# --- Web Routes ---
@app.route('/', methods=['GET', 'POST'])
def index():
    # 세션에서 'authenticated' 상태 확인
    authenticated = session.get('authenticated', False)
    
    if request.method == 'POST':
        # 비밀번호 제출 처리
        if 'password' in request.form:
            entered_password = request.form.get('password')
            if entered_password == ACCESS_PASSWORD:
                session['authenticated'] = True
                # 비밀번호 입력 후 메인 페이지로 리다이렉트하여 GET 요청으로 전환
                return redirect(url_for('index'))
            else:
                flash('비밀번호가 틀렸습니다.', 'danger')
                return render_template('index.html', authenticated=False) # 틀리면 다시 비밀번호 입력 화면
        
        # 이미 인증된 상태에서 검색 폼 제출 처리
        elif authenticated:
            search_results = []
            search_keyword = request.form.get('keyword', '').strip()
            message = ""
            
            sort_by = request.args.get('sort_by')
            sort_order = request.args.get('sort_order', 'asc')

            result_df = pd.DataFrame()

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
                if sort_by in ['평량', '고시가']:
                    result_df[f'{sort_by}_sortable'] = pd.to_numeric(result_by[sort_by], errors='coerce')
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
                    
                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': formatted_고시가,
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
                                   current_sort_order=sort_order)
        else:
            # 인증되지 않은 상태에서 POST 요청이 오면 비밀번호 입력 화면을 다시 보여줌
            return render_template('index.html', authenticated=False)

    # GET 요청 처리 (초기 로드 또는 정렬)
    else:
        if not authenticated:
            return render_template('index.html', authenticated=False)
        else:
            search_results = []
            search_keyword = request.args.get('keyword', '').strip()
            message = ""
            
            sort_by = request.args.get('sort_by')
            sort_order = request.args.get('sort_order', 'asc')

            result_df = pd.DataFrame()

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
                if sort_by in ['평량', '고시가']:
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
                    
                    search_results.append({
                        '품목': row.get('품목', 'N/A'),
                        '사이즈': row.get('사이즈', 'N/A'),
                        '평량': row.get('평량', 'N/A'),
                        '색상_및_패턴': row.get('색상 및 패턴', 'N/A'),
                        '고시가': formatted_고시가,
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
                                   current_sort_order=sort_order)

if __name__ == '__main__':
    app.run(debug=True)
