import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'paper_system_final_v9_perfect_fav'

SITE_PASSWORD = "03877"
cached_data = []
last_updated = ""

def load_data():
    global cached_data, last_updated
    file_path = 'search.xlsx'
    if not os.path.exists(file_path): return
    mtime = os.path.getmtime(file_path)
    last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        combined_list = []
        for sheet_name, df in all_sheets.items():
            df.columns = [str(col).strip() for col in df.columns]
            df = df.ffill()
            col_map = {
                '품목': next((c for c in df.columns if '품목' in c), '품목'),
                '색상': next((c for c in df.columns if '색상' in c or '패턴' in c), '색상'),
                '사이즈': next((c for c in df.columns if '사이즈' in c or '규격' in c), '사이즈'),
                '평량': next((c for c in df.columns if '평량' in c), '평량'),
                '고시가': next((c for c in df.columns if '고시가' in c), '고시가'),
                '두께': next((c for c in df.columns if '두께' in c), '두께')
            }
            temp_df = pd.DataFrame()
            temp_df['품목'] = df[col_map['품목']].astype(str).str.strip()
            # 색상 정보 정제 (nan이나 공백을 '-'로 통일)
            temp_df['색상'] = df[col_map['색상']].astype(str).str.strip() if col_map['색상'] in df.columns else "-"
            temp_df['색상'] = temp_df['색상'].replace(['nan', 'None', ''], '-')
            
            temp_df['사이즈'] = df[col_map['사이즈']].astype(str).str.strip()
            temp_df['평량'] = df[col_map['평량']].astype(str).str.strip()
            temp_df['두께'] = df[col_map['두께']].astype(str).str.strip()
            temp_df['고시가_원본'] = pd.to_numeric(df[col_map['고시가']], errors='coerce').fillna(0)
            temp_df['고시가'] = temp_df['고시가_원본'].apply(lambda x: f"{int(x):,}" if x > 0 else "0")
            temp_df['시트명'] = sheet_name.strip()
            combined_list.append(temp_df)
        if combined_list:
            cached_data = pd.concat(combined_list, ignore_index=True).fillna('-').to_dict('records')
    except Exception as e: print(f"Error: {e}")

load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == SITE_PASSWORD:
            session['authenticated'] = True
            session.permanent = True
            return redirect(url_for('index'))
        else:
            return render_template('index.html', authenticated=False, error="비밀번호가 일치하지 않습니다.")

    if not session.get('authenticated'):
        return render_template('index.html', authenticated=False)

    keyword = ""
    results = []
    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        if keyword:
            results = [item for item in cached_data if keyword.lower() in item['품목'].lower() or keyword.lower() in item['시트명'].lower() or keyword.lower() in item['색상'].lower()]
            for item in results:
                p, s = item['품목'], item['시트명']
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
    return render_template('index.html', results=results, keyword=keyword, authenticated=True, last_updated=last_updated)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)