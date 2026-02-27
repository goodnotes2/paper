import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
from datetime import datetime
import re

app = Flask(__name__)
app.secret_key = 'paper_system_v110_final_confirmed'

SITE_PASSWORD = "03877"
cached_data = []
board_data = [] 
last_updated = ""

def load_data():
    global cached_data, last_updated, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    # 1. 메인 검색 데이터 로드
    if os.path.exists(file_path):
        try:
            mtime = os.path.getmtime(file_path)
            last_updated = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df.columns = [str(col).strip() for col in df.columns]
                # 모든 데이터를 문자열로 변환 (검색 에러 방지 핵심)
                df = df.fillna('').astype(str)
                
                def find_col(keywords, default):
                    for c in df.columns:
                        if any(k in c for k in keywords): return c
                    return default

                col_map = {
                    '품목': find_col(['품목', '종이', '품명'], '품목'),
                    '색상': find_col(['색상', '패턴', '컬러'], '색상'),
                    '평량': find_col(['평량', 'g', '무게'], '평량'),
                    '고시가': find_col(['고시가', '단가', '가격'], '고시가'),
                    '두께': find_col(['두께', 'μm', 'um'], '두께')
                }
                
                if col_map['품목'] not in df.columns: continue

                temp_df = pd.DataFrame()
                temp_df['품목'] = df[col_map['품목']].str.strip()
                temp_df['색상'] = df[col_map['색상']].str.strip() if col_map['색상'] in df.columns else ''
                temp_df['평량'] = df[col_map['평량']].str.replace(r'\.0$', '', regex=True) if col_map['평량'] in df.columns else '0'
                
                # 두께 숫자 추출
                thick_val = df[col_map['두께']] if col_map['두께'] in df.columns else '0'
                temp_df['두께'] = pd.to_numeric(thick_val.str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0).astype(str)
                
                # 고시가 콤마 처리
                price_val = df[col_map['고시가']] if col_map['고시가'] in df.columns else '0'
                nums = pd.to_numeric(price_val.str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
                temp_df['고시가'] = nums.apply(lambda x: f"{int(x):,}" if x > 0 else "0")
                
                temp_df['시트명'] = str(sheet_name).strip()
                temp_df['row_id'] = temp_df.apply(lambda r: f"id_{re.sub(r'[^a-zA-Z0-9]', '', str(r['품목'])+str(r['평량'])+str(r['시트명']))}", axis=1)
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Search Excel Error: {e}")

    # 2. 합지 데이터 로드 (qq.xlsx 수정)
    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path, engine='openpyxl')
            df_qq.columns = [str(col).strip() for col in df_qq.columns]
            # 합지명과 두께 컬럼을 명확히 지정
            name_col = next((c for c in df_qq.columns if '합지' in c or '품명' in c), df_qq.columns[0])
            thick_col = next((c for c in df_qq.columns if '두께' in c or 'mm' in c), df_qq.columns[1])
            
            board_data = []
            for _, row in df_qq.iterrows():
                board_data.append({
                    '합지명': str(row[name_col]).strip(),
                    '두께': pd.to_numeric(str(row[thick_col]).replace('mm',''), errors='coerce') or 0
                })
        except Exception as e:
            print(f"Board Excel Error: {e}")
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
        for item in cached_data:
            # 품목과 색상을 확실히 문자열로 처리하여 검색 (에러 발생 지점 해결)
            name_str = str(item.get('품목', '')).lower()
            color_str = str(item.get('색상', '')).lower()
            if k in name_str or k in color_str:
                p, s = item['품목'], item['시트명']
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
                results.append(item)
    
    return render_template('index.html', results=results, keyword=keyword, 
                           authenticated=True, last_updated=last_updated, boards=board_data)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)