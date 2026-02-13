import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'paper_system_final_v26'

SITE_PASSWORD = "03877"
cached_data = []
last_updated = ""

def load_data():
    global cached_data, last_updated
    file_path = 'search.xlsx'
    if not os.path.exists(file_path): return
        
    try:
        mtime = os.path.getmtime(file_path)
        last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
        
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
            
            # 모든 데이터를 문자열로 변환 (이미지의 float 에러 방지)
            temp_df['품목'] = df[col_map['품목']].fillna('').astype(str).str.strip() if col_map['품목'] in df.columns else "Unknown"
            temp_df['색상'] = df[col_map['색상']].fillna('-').astype(str).str.strip().replace(['nan', 'None', ''], '-') if col_map['색상'] in df.columns else "-"
            temp_df['사이즈'] = df[col_map['사이즈']].fillna('-').astype(str).str.strip() if col_map['사이즈'] in df.columns else "-"
            temp_df['평량'] = df[col_map['평량']].fillna('0').astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if col_map['평량'] in df.columns else "0"
            
            # 두께 및 가격 처리
            temp_df['두께'] = pd.to_numeric(df[col_map['두께']], errors='coerce').fillna(0).astype(str) if col_map['두께'] in df.columns else "0"
            if col_map['고시가'] in df.columns:
                nums = pd.to_numeric(df[col_map['고시가']], errors='coerce').fillna(0)
                temp_df['고시가'] = nums.apply(lambda x: f"{int(x):,}" if x > 0 else "0")
            else:
                temp_df['고시가'] = "0"
                
            temp_df['시트명'] = str(sheet_name).strip()
            
            # [긴급 수정] 즐겨찾기 ID에서 괄호, 공백 등 모든 특수문자 제거
            def create_safe_id(row):
                # 품목+평량+시트명+색상을 합친 뒤, 한글/영어/숫자만 남기고 싹 지웁니다.
                base = f"{row['품목']}{row['평량']}{row['시트명']}{row['색상']}"
                safe_id = re.sub(r'[^a-zA-Z0-9가-힣]', '', base)
                return "id_" + safe_id # 앞에 id_를 붙여서 숫자로 시작하지 않게 보강

            temp_df['row_id'] = temp_df.apply(create_safe_id, axis=1)
            combined_list.append(temp_df)
            
        if combined_list:
            cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
            
    except Exception as e:
        print(f"Excel Loading Error: {e}")

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
            if k_lower in str(item.get('품목', '')).lower() or k_lower in str(item.get('색상', '')).lower():
                p, s = item.get('품목', ''), item.get('시트명', '')
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
                results.append(item)
        
    return render_template('index.html', results=results, keyword=keyword, authenticated=True, last_updated=last_updated)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)