import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'paper_system_v60_final_full'

SITE_PASSWORD = "03877"
cached_data = []
board_data = [] 
last_updated = ""

def load_data():
    global cached_data, last_updated, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    if os.path.exists(file_path):
        try:
            mtime = os.path.getmtime(file_path)
            last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df.columns = [str(col).strip() for col in df.columns]
                df = df.astype(str).replace(['nan', 'None', 'nan '], '')
                
                # 열 이름을 유연하게 찾는 함수
                def find_col(keywords, default):
                    for c in df.columns:
                        if any(k in c for k in keywords): return c
                    return default

                col_map = {
                    '품목': find_col(['품목', '종이', '품명'], '품목'),
                    '색상': find_col(['색상', '패턴', '컬러'], '색상'),
                    '사이즈': find_col(['사이즈', '규격', '크기'], '사이즈'),
                    '평량': find_col(['평량', 'g', '무게'], '평량'),
                    '고시가': find_col(['고시가', '단가', '가격'], '고시가'),
                    '두께': find_col(['두께', 'μm', 'um'], '두께')
                }
                
                if col_map['품목'] not in df.columns: continue

                temp_df = pd.DataFrame()
                temp_df['품목'] = df[col_map['품목']].str.strip()
                temp_df['색상'] = df.get(col_map['색상'], pd.Series(['']*len(df))).str.strip()
                temp_df['사이즈'] = df.get(col_map['사이즈'], pd.Series(['']*len(df))).str.strip()
                temp_df['평량'] = df.get(col_map['평량'], pd.Series(['']*len(df))).str.replace(r'\.0$', '', regex=True)
                temp_df['두께'] = pd.to_numeric(df.get(col_map['두께'], 0), errors='coerce').fillna(0).astype(str)
                
                price_col = df.get(col_map['고시가'], 0)
                nums = pd.to_numeric(price_col, errors='coerce').fillna(0)
                temp_df['고시가'] = nums.apply(lambda x: f"{int(x):,}" if x > 0 else "0")
                
                temp_df['시트명'] = str(sheet_name).strip()
                # ID 생성 시 모든 요소를 문자열로 강제 결합
                def make_id(row):
                    raw = str(row['품목']) + str(row['평량']) + str(row['시트명'])
                    return f"id_{re.sub(r'[^a-zA-Z0-9가-힣]', '', raw)}"
                
                temp_df['row_id'] = temp_df.apply(make_id, axis=1)
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Search Excel Error: {e}")

    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path, engine='openpyxl')
            df_qq.columns = [str(col).strip() for col in df_qq.columns]
            board_data = df_qq.to_dict(orient='records')
        except:
            board_data = [{'합지명': '기본 1000g', '두께': 1.6}]
    else:
        board_data = [{'합지명': '1000g(기본)', '두께': 1.6}, {'합지명': '1200g(기본)', '두께': 1.9}]

load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST' and 'password' in request.form:
        if request.form.get('password') == SITE_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
    if not session.get('authenticated'):
        return render_template('index.html', authenticated=False)

    keyword = request.form.get('keyword', '').strip() if request.method == 'POST' else ""
    results = []
    if keyword:
        k = keyword.lower()
        results = [item for item in cached_data if k in item['품목'].lower() or k in item['색상'].lower()]
        for item in results:
            p, s = item['품목'], item['시트명']
            if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
            elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
            else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
    
    return render_template('index.html', results=results, keyword=keyword, 
                           authenticated=True, last_updated=last_updated, boards=board_data)

if __name__ == '__main__':
    # Render 필수 포트 설정
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)