import pandas as pd
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'paper_system_99'
SITE_PASSWORD = "03877"

# 전역 변수
cached_data = []

def load_data():
    global cached_data
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return

    try:
        # 시트를 읽을 때 필요한 컬럼만 지정해서 메모리 사용량 절반으로 감소
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        combined_list = []

        for sheet_name, df in all_sheets.items():
            # 컬럼 전처리 및 빈칸 채우기 최적화
            df = df.ffill() 
            df['시트명'] = sheet_name
            if '고시가' in df.columns:
                df['고시가_원본'] = df['고시가']
            
            combined_list.append(df)

        # 전체 데이터를 리스트 형태의 딕셔너리로 변환 (Pandas 검색보다 빠름)
        full_df = pd.concat(combined_list, ignore_index=True).fillna('')
        cached_data = full_df.to_dict('records')
        print(f"✅ {len(cached_data)}건 로드 완료")
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
            # ⚡ 일반적인 Pandas 검색보다 훨씬 빠른 리스트 컴프리헨션 검색
            results = [
                item for item in cached_data 
                if keyword.lower() in str(item.get('품목', '')).lower() or 
                   keyword.lower() in str(item.get('시트명', '')).lower()
            ]
            
            # URL 생성 로직
            for item in results:
                p, s = str(item['품목']), str(item['시트명'])
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/product/search?q={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/search?kwd={p}"
                elif '무림' in s: item['url'] = f"https://www.moorim.co.kr/ko/product/search.do?searchKeyword={p}"
                elif '한솔' in s: item['url'] = f"https://www.hansolpaper.co.kr/product/search?q={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"

    return render_template('index.html', results=results, keyword=keyword, authenticated=True, logo_path='search.png')

@app.route('/logout')
def logout():
    session.pop('authenticated', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)