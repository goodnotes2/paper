from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash
import pandas as pd
import numpy as np
import os
import math
import re

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 실제 배포 시에는 더 복잡하고 안전한 키로 변경하세요.
# 업로드된 엑셀 파일의 경로
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# 하드코딩된 비밀번호 (실제 배포 시 환경 변수 사용 권장)
ADMIN_PASSWORD = '111'

# 전역 변수로 데이터프레임과 마지막 검색 결과를 저장
df = None
last_results = None
last_keyword = ""
last_sort_by = ""
last_sort_order = ""

# 세네카 계산을 위한 전역 변수
seneca_selected_thickness = None
seneca_page_count = None
seneca_result = None
seneca_selected_product_info = None

def load_data():
    """엑셀 파일을 읽어 데이터프레임으로 로드하고 전역 변수에 저장합니다."""
    global df
    file_path = os.path.join(app.root_path, 'search.xlsx')
    if not os.path.exists(file_path):
        return False
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, engine='openpyxl')

        # '두께' 열의 'N/A' 또는 'None' 값을 0으로 변환
        df['두께'] = pd.to_numeric(df['두께'], errors='coerce').fillna(0)
        
        # '품목', '사이즈', '평량', '색상 및 패턴' 열의 공백, 하이픈(-)을 NaN으로 변환 후 forward fill
        for col in ['품목', '사이즈', '평량', '색상 및 패턴']:
            df[col] = df[col].replace(['', '-', ' '], np.nan)
        df[['품목', '사이즈', '평량', '색상 및 패턴']] = df[['품목', '사이즈', '평량', '색상 및 패턴']].fillna(method='ffill')

        # '고시가' 열의 공백, 하이픈, 쉼표를 제거하고 숫자형으로 변환
        if '고시가' in df.columns:
            df['고시가'] = df['고시가'].astype(str).str.replace(r'[,\s-]+', '', regex=True)
            df['고시가'] = pd.to_numeric(df['고시가'], errors='coerce').fillna(0)
        
        return True
    except Exception as e:
        print(f"Error loading data: {e}")
        return False

# 애플리케이션 시작 시 데이터 로드
if not load_data():
    print("Warning: search.xlsx not found or could not be loaded.")

@app.route('/', methods=['GET', 'POST'])
def index():
    global last_results, last_keyword, last_sort_by, last_sort_order
    global seneca_selected_thickness, seneca_page_count, seneca_result, seneca_selected_product_info
    
    # 세네카 계산기용 로고 파일 경로
    logo_path = 'logo.png'

    # 비밀번호 인증 로직
    if 'authenticated' not in session:
        session['authenticated'] = False
    
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADMIN_PASSWORD:
            session['authenticated'] = True
            flash('로그인 성공', 'success')
            return redirect(url_for('index'))
        elif password:
            flash('비밀번호가 올바르지 않습니다.', 'danger')
            return redirect(url_for('index'))

    if not session.get('authenticated'):
        return render_template('index.html', authenticated=False)
        
    keyword = request.form.get('keyword', last_keyword)
    last_keyword = keyword
    
    sort_by = request.args.get('sort_by', last_sort_by)
    sort_order = request.args.get('sort_order', last_sort_order)
    last_sort_by = sort_by
    last_sort_order = sort_order

    seneca_page_count = request.form.get('seneca_page_count', seneca_page_count)
    seneca_selected_thickness = request.form.get('seneca_selected_thickness_hidden', seneca_selected_thickness)
    seneca_selected_product_info = request.form.get('seneca_selected_product_info_hidden', seneca_selected_product_info)
    
    if not isinstance(df, pd.DataFrame):
        message = "데이터 파일을 찾을 수 없거나 로드할 수 없습니다. 파일을 확인해 주세요."
        return render_template('index.html', message=message, authenticated=True, logo_path=logo_path)

    results = []
    message = ""

    if keyword:
        keyword = keyword.lower()
        
        # '두께' 열을 문자열로 변환하여 검색에 포함
        search_df = df.astype(str)
        
        # '두께' 열이 있을 경우, 소수점 두 자리까지만 표시되도록 포맷팅
        if '두께' in df.columns:
            df_display = df.copy()
            df_display['두께'] = df_display['두께'].apply(lambda x: f'{x:.2f}' if isinstance(x, (float, np.floating)) else x)
        else:
            df_display = df.copy()
        
        # '고시가' 열이 있을 경우, 쉼표 포맷팅
        if '고시가' in df_display.columns:
            df_display['고시가'] = df_display['고시가'].apply(lambda x: f'{int(x):,}' if pd.api.types.is_numeric_dtype(type(x)) else x)

        # 여러 열에 걸쳐 키워드 검색
        results_df = df_display[
            search_df.apply(lambda row: row.astype(str).str.lower().str.contains(keyword, na=False).any(), axis=1)
        ].copy()

        if not results_df.empty:
            # 정렬
            if sort_by and sort_by in results_df.columns:
                try:
                    # 두께와 고시가는 숫자형으로 정렬
                    if sort_by in ['두께', '고시가']:
                        sorted_df = results_df.sort_values(
                            by=sort_by,
                            ascending=(sort_order == 'asc'),
                            kind='stable',
                            na_position='last'
                        )
                    else:
                        # 나머지 열은 문자열로 정렬
                        sorted_df = results_df.sort_values(
                            by=sort_by,
                            ascending=(sort_order == 'asc'),
                            key=lambda x: x.astype(str).str.lower(),
                            kind='stable'
                        )
                    results = sorted_df.to_dict('records')
                except KeyError:
                    results = results_df.to_dict('records')
            else:
                results = results_df.to_dict('records')
        else:
            message = "검색 결과가 없습니다."
    
    else: # 키워드가 없을 때 전체 데이터 표시 (초기 화면)
        if sort_by and sort_by in df.columns:
            try:
                # '두께'와 '고시가'는 숫자형으로 정렬
                if sort_by in ['두께', '고시가']:
                    sorted_df = df.sort_values(
                        by=sort_by,
                        ascending=(sort_order == 'asc'),
                        kind='stable',
                        na_position='last'
                    )
                else:
                    # 나머지 열은 문자열로 정렬
                    sorted_df = df.sort_values(
                        by=sort_by,
                        ascending=(sort_order == 'asc'),
                        key=lambda x: x.astype(str).str.lower(),
                        kind='stable'
                    )
                
                # 정렬된 데이터프레임의 '두께' 열 포맷팅
                if '두께' in sorted_df.columns:
                    sorted_df['두께'] = sorted_df['두께'].apply(lambda x: f'{x:.2f}' if isinstance(x, (float, np.floating)) else x)
                
                # '고시가' 열 포맷팅
                if '고시가' in sorted_df.columns:
                    sorted_df['고시가'] = sorted_df['고시가'].apply(lambda x: f'{int(x):,}' if pd.api.types.is_numeric_dtype(type(x)) and not math.isnan(x) else x)
                    
                results = sorted_df.to_dict('records')
            except KeyError:
                results = df.to_dict('records')
        else:
            df_display = df.copy()
            if '두께' in df_display.columns:
                df_display['두께'] = df_display['두께'].apply(lambda x: f'{x:.2f}' if isinstance(x, (float, np.floating)) else x)
            if '고시가' in df_display.columns:
                 df_display['고시가'] = df_display['고시가'].apply(lambda x: f'{int(x):,}' if pd.api.types.is_numeric_dtype(type(x)) and not math.isnan(x) else x)
            results = df_display.to_dict('records')

    return render_template(
        'index.html', 
        authenticated=True, 
        results=results, 
        message=message, 
        keyword=keyword, 
        current_sort_by=sort_by, 
        current_sort_order=sort_order,
        logo_path=logo_path,
        seneca_page_count=seneca_page_count,
        seneca_result=seneca_result,
        seneca_selected_product_info=seneca_selected_product_info,
        seneca_selected_thickness=seneca_selected_thickness
    )

@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    """세네카 계산을 수행하는 API 엔드포인트"""
    data = request.json
    page_count_str = data.get('page_count')
    thickness_str = data.get('thickness')
    
    if not page_count_str or not thickness_str:
        return jsonify({'error': '페이지 수와 두께 정보가 필요합니다.'}), 400
    
    try:
        page_count = float(page_count_str)
        thickness = float(thickness_str)
        
        # 세네카 계산: (페이지 수 / 2) * 두께
        seneca = (page_count / 2) * thickness
        
        return jsonify({'result': round(seneca, 2)})
    except (ValueError, TypeError):
        return jsonify({'error': '유효한 숫자 값을 입력해 주세요.'}), 400
        
if __name__ == '__main__':
    app.run(debug=True)