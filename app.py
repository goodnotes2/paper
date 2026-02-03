import pd as pd
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import sys
import traceback
import base64

app = Flask(__name__)

# 1. SECRET_KEY 설정
raw_secret_key_base64 = os.environ.get('FLASK_SECRET_KEY', 'your_single_access_secret_key_base64_default')
try:
    app.config['SECRET_KEY'] = base64.urlsafe_b64decode(raw_secret_key_base64)
except:
    app.config['SECRET_KEY'] = b'fallback_secret_key_if_decoding_fails'

# 2. 설정 변수
excel_file_name = 'search.xlsx'
image_file_name = 'search.png'
sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']
ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', '1234')  # 기본비번 설정

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

# 3. 데이터 로드 함수 (Render 경로 최적화)
def load_data():
    # 현재 파일의 절대 경로를 기준으로 설정
    base_path = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_path, excel_file_name)
    
    if not os.path.exists(excel_path):
        print(f"[ERROR] 파일을 찾을 수 없습니다: {excel_path}", file=sys.stderr)
        return pd.DataFrame()

    print(f"[DEBUG] 파일 로드 시작: {excel_path}", file=sys.stderr)
    data_list = []

    for sheet in sheets:
        try:
            # openpyxl 엔진 사용
            df = pd.read_excel(excel_path, sheet_name=sheet, engine='openpyxl')
            df.columns = df.columns.str.strip() # 컬럼명 공백 제거
            
            # 검색에 필요한 컬럼들이 있는지 확인 및 전처리
            target_cols = ['품목', '사이즈', '평량', '색상 및 패턴', '두께', '고시가']
            for col in target_cols:
                if col in df.columns:
                    # 병합된 셀(NaN) 채우기
                    df[col] = df[col].ffill()
            
            # HTML에서 접근하기 쉽게 공백이 포함된 컬럼명은 언더바로 변경된 데이터 생성
            if '색상 및 패턴' in df.columns:
                df['색상_및_패턴'] = df['색상 및 패턴']
            
            # 검색용 원본 고시가 보관
            if '고시가' in df.columns:
                df['고시가_원본'] = df['고시가']

            df['시트명'] = sheet
            # 해당 제조사 URL 연결
            df['url'] = company_urls.get(sheet, '#') + df['품목'].astype(str)
            
            data_list.append(df)
            print(f"[SUCCESS] 시트 로드 성공: {sheet}", file=sys.stderr)
            
        except Exception as e:
            print(f"[ERROR] 시트 '{sheet}' 로드 중 오류 발생: {e}", file=sys.stderr)
    
    if data_list:
        combined_df = pd.concat(data_list, ignore_index=True)
        # 모든 NaN 값을 빈 문자열로 변환 (검색 시 에러 방지)
        return combined_df.fillna('')
    return pd.DataFrame()

# 앱 실행 시 데이터 로드
df_all = load_data()

# 4. 세네카 계산 로직
def calculate_seneca_value(page_count, thickness):
    try:
        p = float(page_count)
        t = float(thickness)
        # 일반적인 세네카 공식: (페이지 / 2) * 두께(mm)
        result = (p / 2) * t
        return round(result, 2)
    except:
        return None

# 5. 라우트 설정
@app.route('/', methods=['GET', 'POST'])
def index():
    # 비밀번호 인증 로직
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == ACCESS_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            flash('비밀번호가 올바르지 않습니다.', 'danger')

    authenticated = session.get('authenticated', False)
    
    # 검색 및 결과 처리
    keyword = request.form.get('keyword', request.args.get('keyword', ''))
    results = []
    
    if authenticated and not df_all.empty and keyword:
        # 품목명이나 시트명에 키워드가 포함된 경우 검색 (대소문자 무시)
        mask = (
            df_all['품목'].str.contains(keyword, case=False, na=False) |
            df_all['시트명'].str.contains(keyword, case=False, na=False)
        )
        # 결과를 딕셔너리 형태로 변환하여 전송
        results = df_all[mask].to_dict('records')

    return render_template('index.html', 
                           authenticated=authenticated, 
                           results=results, 
                           keyword=keyword,
                           logo_path=image_file_name)

@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    data = request.get_json()
    page_count = data.get('page_count')
    thickness = data.get('thickness')
    
    result = calculate_seneca_value(page_count, thickness)
    if result is not None:
        return jsonify({'result': result})
    return jsonify({'error': '계산 불가'}), 400

if __name__ == '__main__':
    # Render 환경의 포트 대응
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)