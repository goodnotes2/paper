import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'paper_system_v35_render_fix'

SITE_PASSWORD = "03877"
cached_data = []
board_data = [] 
last_updated = ""

def load_data():
    global cached_data, last_updated, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    # 1. search.xlsx 로드 (오류 철저 방어)
    if os.path.exists(file_path):
        try:
            mtime = os.path.getmtime(file_path)
            last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df.columns = [str(col).strip() for col in df.columns]
                # 모든 데이터를 문자열로 변환 후 결측치 처리 (fillna 에러 방지)
                df = df.astype(str).replace(['nan', 'None', 'nan '], '')
                
                col_map = {
                    '품목': next((c for c in df.columns if '품목' in c), '품목'),
                    '색상': next((c for c in df.columns if '색상' in c or '패턴' in c), '색상'),
                    '사이즈': next((c for c in df.columns if '사이즈' in c or '규격' in c), '사이즈'),
                    '평량': next((c for c in df.columns if '평량' in c), '평량'),
                    '고시가': next((c for c in df.columns if '고시가' in c), '고시가'),
                    '두께': next((c for c in df.columns if '두께' in c), '두께')
                }
                
                temp_df = pd.DataFrame()
                temp_df['품목'] = df[col_map['품목']].str.strip()
                temp_df['색상'] = df[col_map['색상']].str.strip()
                temp_df['사이즈'] = df[col_map['사이즈']].str.strip()
                temp_df['평량'] = df[col_map['평량']].str.strip().str.replace(r'\.0$', '', regex=True)
                temp_df['두께'] = pd.to_numeric(df[col_map['두께']], errors='coerce').fillna(0).astype(str)
                
                if col_map['고시가'] in df.columns:
                    nums = pd.to_numeric(df[col_map['고시가']], errors='coerce').fillna(0)
                    temp_df['고시가'] = nums.apply(lambda x: f"{int(x):,}" if x > 0 else "0")
                else:
                    temp_df['고시가'] = "0"
                
                temp_df['시트명'] = str(sheet_name).strip()
                temp_df['row_id'] = temp_df.apply(lambda r: f"id_{re.sub(r'[^a-zA-Z0-9가-힣]', '', r['품목']+r['평량']+r['시트명'])}", axis=1)
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Search Excel Error: {e}")

    # 2. qq.xlsx 로드 (합지 데이터)
    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path, engine='openpyxl')
            df_qq.columns = [str(col).strip() for col in df_qq.columns]
            board_data = df_qq.to_dict(orient='records')
        except:
            board_data = [{'합지명': '1000g(기본)', '두께': 1.6}]
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
        k_lower = keyword.lower()
        for item in cached_data:
            if k_lower in item['품목'].lower() or k_lower in item['색상'].lower():
                # 두성/삼원 링크 생성 로직 유지
                p, s = item['품목'], item['시트명']
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
                results.append(item)
    
    return render_template('index.html', results=results, keyword=keyword, 
                           authenticated=True, last_updated=last_updated, boards=board_data)

if __name__ == '__main__':
    # Render 포트 오류 해결을 위한 핵심 코드
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)