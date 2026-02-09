import pandas as pd
from flask import Flask, render_template, request, jsonify
import os

app = Flask(__name__)
app.secret_key = 'paper_search_key_1234'

# 전역 변수로 데이터 캐싱
cached_df = None

def load_data_once():
    global cached_df
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        print("❌ search.xlsx 파일을 찾을 수 없습니다.")
        cached_df = pd.DataFrame()
        return

    try:
        # 모든 시트 읽기
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        combined_list = []

        for sheet_name, df in all_sheets.items():
            # 컬럼명 정리
            df.columns = [str(col).strip() for col in df.columns]
            
            # 📌 1. 병합 셀/빈칸 처리 (가장 중요한 부분)
            # '품목'이 비어있으면 위 행의 값을 가져옵니다.
            target_cols = ['품목', '사이즈', '평량', '두께']
            for col in target_cols:
                if col in df.columns:
                    df[col] = df[col].ffill()

            # 시트명 저장 및 고시가 원본 보존
            df['시트명'] = sheet_name
            if '고시가' in df.columns:
                df['고시가_원본'] = df['고시가']

            combined_list.append(df)

        cached_df = pd.concat(combined_list, ignore_index=True).fillna('')
        print(f"✅ 총 {len(cached_df)}행 데이터 로드 완료")
    except Exception as e:
        print(f"❌ 데이터 로드 에러: {e}")
        cached_df = pd.DataFrame()

# 서버 시작 시 데이터 로드
load_data_once()

@app.route('/', methods=['GET', 'POST'])
def index():
    keyword = ""
    results = []
    authenticated = True  # 비밀번호 로직은 기존 유지하되 여기서는 True로 설정

    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        
        if keyword and cached_df is not None:
            # 📌 2. 검색 로직 강화: 품목명 또는 시트명에서 검색
            mask = cached_df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)
            search_df = cached_df[mask].copy()

            # 📌 3. 각 제지사별 정확한 검색 URL 생성
            def make_url(row):
                p = str(row['품목'])
                s = str(row['시트명'])
                if '두성' in s:
                    return f"https://www.doosungpaper.co.kr/product/search?q={p}"
                elif '삼원' in s:
                    return f"https://www.samwonpaper.com/product/search?kwd={p}"
                elif '무림' in s:
                    return f"https://www.moorim.co.kr/ko/product/search.do?searchKeyword={p}"
                elif '한솔' in s:
                    return f"https://www.hansolpaper.co.kr/product/search?q={p}"
                else:
                    return f"https://www.google.com/search?q={s}+{p}"

            search_df['url'] = search_df.apply(make_url, axis=1)
            results = search_df.to_dict('records')

    return render_template('index.html', 
                           results=results, 
                           keyword=keyword, 
                           authenticated=authenticated,
                           logo_path='search.png')

# 세네카 계산 API
@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    data = request.json
    try:
        pages = float(data.get('page_count', 0))
        thickness = float(data.get('thickness', 0))
        # 계산식: (페이지/2) * 두께 / 1000
        result = (pages / 2) * thickness / 1000
        return jsonify({'result': round(result, 2)})
    except:
        return jsonify({'error': '계산 불가'}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)