import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for, flash
import os

app = Flask(__name__)
# 세션 보안을 위한 키
app.secret_key = 'paper_system_secure_99'

# 📌 설정하신 비밀번호
SITE_PASSWORD = "03877"

# 전역 변수로 데이터를 저장 (메모리 캐싱으로 속도 향상)
cached_data = []

def load_data():
    global cached_data
    file_path = 'search.xlsx'
    
    if not os.path.exists(file_path):
        print("❌ search.xlsx 파일을 찾을 수 없습니다.")
        return

    try:
        # 엑셀의 모든 시트를 읽어옵니다.
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        combined_list = []

        for sheet_name, df in all_sheets.items():
            # 1. 컬럼명 양쪽 공백 제거 및 문자열 변환
            df.columns = [str(col).strip() for col in df.columns]
            
            # 2. 병합된 셀(빈칸)을 위 데이터로 채움 (ffill)
            df = df.ffill()
            
            # 3. 유연한 컬럼 매핑 (엑셀마다 조금씩 다른 컬럼명 대응)
            col_map = {
                '품목': next((c for c in df.columns if '품목' in c), '품목'),
                '사이즈': next((c for c in df.columns if '사이즈' in c or '규격' in c), '사이즈'),
                '평량': next((c for c in df.columns if '평량' in c), '평량'),
                '색상': next((c for c in df.columns if '색상' in c or '패턴' in c), '색상 및 패턴'),
                '고시가': next((c for c in df.columns if '고시가' in c), '고시가'),
                '두께': next((c for c in df.columns if '두께' in c), '두께')
            }

            # 4. 데이터 표준화 및 문자열 변환 (평량, 색상 누락 방지)
            temp_df = pd.DataFrame()
            temp_df['품목'] = df[col_map['품목']].astype(str)
            temp_df['사이즈'] = df[col_map['사이즈']].astype(str)
            temp_df['평량'] = df[col_map['평량']].astype(str)
            temp_df['색상_및_패턴'] = df[col_map['색상']].astype(str)
            temp_df['두께'] = df[col_map['두께']].astype(str)
            
            # 고시가는 숫자로 유지 (나중에 500으로 나누기 계산 위해)
            temp_df['고시가_원본'] = pd.to_numeric(df[col_map['고시가']], errors='coerce').fillna(0)
            # 초기 화면 표시용 콤마 적용
            temp_df['고시가'] = temp_df['고시가_원본'].apply(lambda x: f"{int(x):,}" if x > 0 else "0")
            
            temp_df['시트명'] = sheet_name
            combined_list.append(temp_df)

        # 모든 시트 합치기
        if combined_list:
            full_df = pd.concat(combined_list, ignore_index=True).fillna('')
            cached_data = full_df.to_dict('records')
            print(f"✅ 데이터 로드 완료: {len(cached_data)}행")
            
    except Exception as e:
        print(f"❌ 데이터 로드 중 오류 발생: {e}")

# 서버 시작 시 데이터 로드
load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    # 비밀번호 인증 체크
    if 'authenticated' not in session:
        if request.method == 'POST' and 'password' in request.form:
            if request.form.get('password') == SITE_PASSWORD:
                session['authenticated'] = True
                return redirect(url_for('index'))
            else:
                flash('비밀번호가 틀렸습니다.', 'danger')
        return render_template('index.html', authenticated=False)

    keyword = ""
    results = []

    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        if keyword:
            # 품목명 또는 시트명(제지사)으로 검색
            results = [
                item for item in cached_data 
                if keyword.lower() in item['품목'].lower() or 
                   keyword.lower() in item['시트명'].lower()
            ]
            
            # 🔗 요청하신 제지사별 URL 딕셔너리 반영 (검색어 파라미터 포함)
            for item in results:
                p, s = item['품목'], item['시트명']
                if '두성' in s:
                    item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s:
                    item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                elif '한국' in s:
                    item['url'] = "https://www.hankukpaper.com/ko/product/listProductinfo.do"
                elif '무림' in s:
                    item['url'] = "https://www.moorim.co.kr:13002/product/paper.php"
                elif '삼화' in s:
                    item['url'] = f"https://www.samwhapaper.com/product/samwhapaper/all?keyword={p}"
                elif '서경' in s:
                    item['url'] = f"https://wedesignpaper.com/search?type=shopping&sort=consensus_desc&keyword={p}"
                elif '한솔' in s:
                    item['url'] = "https://www.hansolpaper.co.kr/product/insper"
                elif '전주' in s:
                    item['url'] = "https://jeonjupaper.com/publicationpaper"
                else:
                    item['url'] = f"https://www.google.com/search?q={s}+{p}"

    return render_template('index.html', results=results, keyword=keyword, authenticated=True, logo_path='search.png')

@app.route('/logout')
def logout():
    session.pop('authenticated', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)