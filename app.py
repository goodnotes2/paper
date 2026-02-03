# app.py
import pandas as pd
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
# 요청하신 비밀번호 반영
ACCESS_PASSWORD = os.environ.get('APP_ACCESS_PASSWORD', '03877')

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

# 3. 데이터 로드 함수
def load_data():
    base_path = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_path, excel_file_name)
    
    if not os.path.exists(excel_path):
        print(f"[ERROR] 파일을 찾을 수 없음: {excel_path}", file=sys.stderr)
        return pd.DataFrame()

    all_data = []
    for sheet in sheets:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet, engine='openpyxl')
            df.columns = df.columns.str.strip()
            
            # 병합된 셀 처리
            cols_to_fill = ['품목', '사이즈', '평량', '색상 및 패턴']
            for col in cols_to_fill:
                if col in df.columns:
                    df[col] = df[col].ffill()

            # HTML 템플릿 호환용 컬럼명 생성
            if '색상 및 패턴' in df.columns:
                df['색상_및_패턴'] = df['색상 및 패턴']
            if '고시가' in df.columns:
                df['고시가_원본'] = df['고시가']

            df['시트명'] = sheet
            df['url'] = company_urls.get(sheet, '#') + df['품목'].astype(str)
            all_data.append(df)
            print(f"[SUCCESS] '{sheet}' 시트 로드 완료", file=sys.stderr)
        except Exception as e:
            print(f"[ERROR] '{sheet}' 시트 로드 실패: {e}", file=sys.stderr)

    if all_data:
        return pd.concat(all_data, ignore_index=True).fillna('')
    return pd.DataFrame()

df_all = load_data()

# 4. 세네카 계산 API (마이크로미터 변환 로직 포함)
@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    data = request.get_json()
    try:
        page = float(data.get('page_count', 0))
        # 엑셀의 마이크로미터(um) 단위를 밀리미터(mm)로 변환하기 위해 1000으로 나눔
        thickness_um = float(data.get('thickness', 0))
        thickness_mm = thickness_um / 1000 
        
        # 세네카 공식: (페이지 / 2) * 두께(mm)
        result_mm = (page / 2) * thickness_mm
        
        return jsonify({'result': round(result_mm, 2)})
    except Exception as e:
        print(f"[ERROR] 계산 중 오류 발생: {e}", file=sys.stderr)
        return jsonify({'error': '계산 불가'}), 400

# 5. 메인 페이지 라우트
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == ACCESS_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            flash('비밀번호가 올바르지 않습니다.', 'danger')

    authenticated = session.get('authenticated', False)
    keyword = request.form.get('keyword', request.args.get('keyword', ''))
    results = []

    if authenticated and not df_all.empty and keyword:
        # 검색 필터링 (품목명 또는 시트명 포함 여부)
        mask = (
            df_all['품목'].astype(str).str.contains(keyword, case=False, na=False) |
            df_all['시트명'].astype(str).str.contains(keyword, case=False, na=False)
        )
        results = df_all[mask].to_dict('records')

    return render_template('index.html', 
                           authenticated=authenticated, 
                           results=results, 
                           keyword=keyword,
                           logo_path=image_file_name)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)