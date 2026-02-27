import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
import re

app = Flask(__name__)
app.secret_key = 'paper_pro_v130_final'

SITE_PASSWORD = "03877"
cached_data = []
board_data = [] 

def load_data():
    global cached_data, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    if os.path.exists(file_path):
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df = df.fillna('').astype(str)
                df.columns = [str(col).strip() for col in df.columns]
                
                # 컬럼 매핑
                c_name = next((c for c in df.columns if any(k in c for k in ['품목', '종이', '품명'])), None)
                c_thick = next((c for c in df.columns if any(k in c for k in ['두께', 'μm', 'um'])), None)
                c_gram = next((c for c in df.columns if any(k in c for k in ['평량', 'g'])), None)
                c_color = next((c for c in df.columns if any(k in c for k in ['색상', '컬러'])), None)
                c_price = next((c for c in df.columns if any(k in c for k in ['고시가', '단가'])), None)

                if not c_name: continue

                temp_df = pd.DataFrame()
                temp_df['품목'] = df[c_name].str.strip()
                temp_df['색상'] = df[c_color].str.strip() if c_color else ''
                temp_df['평량'] = df[c_gram].str.replace(r'\.0$', '', regex=True) if c_gram else '0'
                
                # 두께 데이터 정제: 숫자와 소수점만 남김 (계산 오류 방지 핵심)
                def clean_thick(val):
                    res = re.sub(r'[^0-9.]', '', str(val))
                    return res if res else '0'
                
                temp_df['두께'] = df[c_thick].apply(clean_thick) if c_thick else '0'
                
                # 가격 콤마
                def clean_price(val):
                    try:
                        num = int(float(re.sub(r'[^0-9.]', '', str(val))))
                        return f"{num:,}"
                    except: return "0"
                temp_df['고시가'] = df[c_price].apply(clean_price) if c_price else "0"
                
                temp_df['시트명'] = str(sheet_name).strip()
                temp_df['row_id'] = temp_df.apply(lambda r: f"id_{re.sub(r'[^a-zA-Z0-9]', '', r['품목']+r['평량']+r['시트명'])}", axis=1)
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e: print(f"Search Error: {e}")

    # 합지 데이터 (qq.xlsx)
    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path).fillna('')
            board_data = []
            for _, row in df_qq.iterrows():
                name = str(row.iloc[0]).strip()
                try: thick = float(re.sub(r'[^0-9.]', '', str(row.iloc[1])))
                except: thick = 0
                if name: board_data.append({'합지명': name, '두께': thick})
        except: board_data = [{'합지명': '1000g(기본)', '두께': 1.6}]
    
    if not board_data: board_data = [{'합지명': '1000g(기본)', '두께': 1.6}]

load_data()

@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('authenticated'):
        if request.method == 'POST' and request.form.get('password') == SITE_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        return render_template('index.html', authenticated=False)

    keyword = request.form.get('keyword', '').strip()
    results = []
    if keyword:
        k = keyword.lower()
        for item in cached_data:
            if k in item['품목'].lower() or k in item['색상'].lower():
                p, s = item['품목'], item['시트명']
                if '두성' in s: item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s: item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else: item['url'] = f"https://www.google.com/search?q={s}+{p}"
                results.append(item)
    
    return render_template('index.html', results=results, keyword=keyword, authenticated=True, boards=board_data)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))