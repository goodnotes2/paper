import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'paper_system_v32_final'

SITE_PASSWORD = "03877"
cached_data = []
board_data = []
last_updated = ""

def load_data():
    global cached_data, last_updated, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    # 1. search.xlsx 로드 (오류 방지 강화)
    if os.path.exists(file_path):
        try:
            mtime = os.path.getmtime(file_path)
            last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df.columns = [str(col).strip() for col in df.columns]
                # fillna 이전에 모든 데이터를 문자열로 변환하여 'str object has no attribute fillna' 오류 방지
                df = df.astype(str).replace(['nan', 'None', 'nan '], '')
                
                col_map = {
                    '품목': next((c for c in df.columns if '품목' in c), '품목'),
                    '두께': next((c for c in df.columns if '두께' in c), '두께'),
                    '평량': next((c for c in df.columns if '평량' in c), '평량')
                }
                
                temp_df = pd.DataFrame()
                temp_df['품목'] = df[col_map['품목']].str.strip()
                temp_df['두께'] = pd.to_numeric(df[col_map['두께']], errors='coerce').fillna(0).astype(str)
                temp_df['평량'] = df[col_map['평량']].str.replace(r'\.0$', '', regex=True)
                temp_df['시트명'] = str(sheet_name).strip()
                temp_df['row_id'] = temp_df.apply(lambda r: f"id_{re.sub(r'[^a-zA-Z0-9가-힣]', '', r['품목']+r['평량']+r['시트명'])}", axis=1)
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Excel Load Error: {e}")

    # 2. qq.xlsx 로드 (도면용 합지 데이터)
    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path, engine='openpyxl')
            board_data = df_qq.to_dict(orient='records')
        except:
            board_data = [{'합지명': '기본 1000g', '두께': 1.6}]
    else:
        board_data = [{'합지명': '기본 1000g', '두께': 1.6}]

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
        k_lower = keyword.lower()
        for item in cached_data:
            if k_lower in item['품목'].lower():
                results.append(item)
    
    return render_template('index.html', results=results, keyword=keyword, 
                           authenticated=True, last_updated=last_updated, boards=board_data)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)