import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'paper_system_99'
SITE_PASSWORD = "03877"

cached_data = []

def load_data():
    global cached_data
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return

    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        combined_list = []

        for sheet_name, df in all_sheets.items():
            df.columns = [str(col).strip() for col in df.columns]
            df = df.ffill()
            
            # 컬럼 매핑 보정 (평량, 색상 등 누락 방지)
            col_map = {
                '품목': next((c for c in df.columns if '품목' in c), '품목'),
                '사이즈': next((c for c in df.columns if '사이즈' in c or '규격' in c), '사이즈'),
                '평량': next((c for c in df.columns if '평량' in c), '평량'),
                '색상': next((c for c in df.columns if '색상' in c or '패턴' in c), '색상 및 패턴'),
                '고시가': next((c for c in df.columns if '고시가' in c), '고시가'),
                '두께': next((c for c in df.columns if '두께' in c), '두께')
            }

            new_df = pd.DataFrame()
            new_df['품목'] = df[col_map['품목']].astype(str)
            new_df['사이즈'] = df[col_map['사이즈']].astype(str)
            new_df['평량'] = df[col_map['평량']].astype(str)
            new_df['색상_및_패턴'] = df[col_map['색상']].astype(str)
            new_df['두께'] = df[col_map['두께']].astype(str)
            new_df['고시가_원본'] = df[col_map['고시가']]
            new_df['시트명'] = sheet_name
            
            combined_list.append(new_df)

        full_df = pd.concat(combined_list, ignore_index=True).fillna('')
        cached_data = full_df.to_dict('records')
    except Exception as e:
        print(f"Error: {e}")

load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    if 'authenticated' not in session:
        if request.method == 'POST' and 'password' in request.form:
            if request.form.get('password') == SITE_PASSWORD:
                session['authenticated'] = True
                return redirect(url_for('index'))
            flash('비밀번호가 틀렸습니다.', 'danger')
        return render_template('index.html', authenticated=False)

    keyword = ""
    results = []
    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        if keyword:
            results = [
                item for item in cached_data 
                if keyword.lower() in item['품목'].lower() or 
                   keyword.lower() in item['시트명'].lower()
            ]
            
            # 🔗 제지사별 URL 딕셔너리 반영 로직
            for item in results:
                p, s = item['품목'], item['시트명']
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/product/search?q={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list" # 삼원은 목록 페이지로
                elif '한국' in s: item['url'] = f"https://www.hankukpaper.com/ko/product/listProductinfo.do"
                elif '무림' in s: item['url'] = f"https://www.moorim.co.kr/ko/product/search.do?searchKeyword={p}"
                elif '삼화' in s: item['url'] = f"https://www.samwhapaper.com/"
                elif '서경' in s: item['url'] = f"https://wedesignpaper.com/#"
                elif '한솔' in s: item['url'] = f"https://www.hansolpaper.co.kr/product/insper"
                elif '전주' in s: item['url'] = f"https://jeonjupaper.com/publicationpaper"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"

    return render_template('index.html', results=results, keyword=keyword, authenticated=True, logo_path='search.png')

# ... 나머지 logout 및 실행 코드 동일 ...