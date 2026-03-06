import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
import re
from urllib.parse import quote

app = Flask(__name__)
app.secret_key = 'expert_version_v20'

SITE_PASSWORD = "03877"
cached_data = []

BOARD_DATA = [
    {'합지명': '1000g', '두께': 1.6, '각양장_앞뒤': 3.0, '미소_앞뒤': 2.0},
    {'합지명': '1100g', '두께': 1.8, '각양장_앞뒤': 3.0, '미소_앞뒤': 2.0},
    {'합지명': '1200g', '두께': 2.0, '각양장_앞뒤': 3.5, '미소_앞뒤': 2.5},
    {'합지명': '1300g', '두께': 2.2, '각양장_앞뒤': 4.0, '미소_앞뒤': 3.0},
    {'합지명': '1400g', '두께': 2.4, '각양장_앞뒤': 4.5, '미소_앞뒤': 3.5},
]

def load_data():
    global cached_data
    file_path = 'search.xlsx'
    if os.path.exists(file_path):
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df = df.fillna('').astype(str)
                df.columns = [str(col).strip() for col in df.columns]

                c_name  = next((c for c in df.columns if any(k in c for k in ['품목', '종이', '품명'])), None)
                c_thick = next((c for c in df.columns if any(k in c for k in ['두께', 'μm', 'um'])), None)
                c_gram  = next((c for c in df.columns if any(k in c for k in ['평량', 'g'])), None)
                c_color = next((c for c in df.columns if any(k in c for k in ['색상', '컬러'])), None)
                c_price = next((c for c in df.columns if any(k in c for k in ['고시가', '단가'])), None)
                c_size  = next((c for c in df.columns if any(k in c for k in ['사이즈', '규격', 'size'])), None)

                if not c_name:
                    continue

                def extract_num(val):
                    res = re.sub(r'[^0-9.]', '', str(val))
                    return res if res and res != '.' else '0'

                temp_df = pd.DataFrame()
                temp_df['품목']  = df[c_name].str.strip()
                temp_df['색상']  = df[c_color].str.strip() if c_color else ''
                temp_df['평량']  = df[c_gram].str.replace(r'\.0$', '', regex=True) if c_gram else '0'
                temp_df['두께']  = df[c_thick].apply(extract_num) if c_thick else '0'
                temp_df['사이즈'] = df[c_size].str.strip() if c_size else ''
                temp_df['고시가'] = df[c_price].apply(
                    lambda x: f"{int(float(extract_num(x))):,}" if float(extract_num(x)) > 0 else "0"
                )
                temp_df['시트명'] = str(sheet_name).strip()
                temp_df['row_id'] = temp_df.apply(
                    lambda r: f"id_{re.sub(r'[^a-zA-Z0-9]', '', r['품목']+r['평량']+r['시트명'])}", axis=1
                )
                combined_list.append(temp_df)

            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Search Error: {e}")

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
                item_copy = dict(item)
                p, s = item_copy['품목'], item_copy['시트명']
                if '두성' in s:
                    item_copy['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={quote(p)}"
                elif '삼원' in s:
                    item_copy['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={quote(p)}"
                else:
                    item_copy['url'] = f"https://www.google.com/search?q={quote(s+' '+p)}"
                results.append(item_copy)

    return render_template('index.html',
                           results=results,
                           keyword=keyword,
                           authenticated=True,
                           boards=BOARD_DATA)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
