import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for
import os
import re

app = Flask(__name__)
app.secret_key = 'expert_version_v20'

SITE_PASSWORD = "03877"
cached_data = []
board_data = []

def load_data():
    global cached_data, board_data
    file_path = 'search.xlsx'
    qq_path = 'qq.xlsx'
    
    # 1. search.xlsx 로드 (전체 열 검색 강화)
    if os.path.exists(file_path):
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            combined_list = []
            for sheet_name, df in all_sheets.items():
                df = df.fillna('').astype(str)
                df.columns = [str(col).strip() for col in df.columns]
                
                # 품목/비고/색상 열 모두 찾기
                c_name = next((c for c in df.columns if any(k in str(c).upper() for k in ['품목', '품명', 'NAME'])), None)
                c_color = next((c for c in df.columns if any(k in str(c).upper() for k in ['색상', 'COLOR', '패턴'])), None)
                c_note = next((c for c in df.columns if any(k in str(c).upper() for k in ['비고', '메모', 'NOTE'])), None)
                c_thick = next((c for c in df.columns if any(k in c for k in ['두께', 'μm', 'um'])), None)
                c_gram = next((c for c in df.columns if any(k in c for k in ['평량', 'g'])), None)
                c_price = next((c for c in df.columns if any(k in c for k in ['고시가', '단가'])), None)
                
                if not c_name: 
                    continue
                
                # 검색용 전체 텍스트 (품목+색상+비고+시트명)
                search_parts = [df[c_name]]
                if c_color: search_parts.append(' ' + df[c_color])
                if c_note: search_parts.append(' ' + df[c_note])
                search_text = ''.join(search_parts).str.strip()
                search_nospace = search_text.str.replace(' ', '').str.lower()
                
                temp_df = pd.DataFrame()
                temp_df['search_text'] = search_text
                temp_df['search_nospace'] = search_nospace
                temp_df['sheet_name'] = str(sheet_name).strip()
                temp_df['품목'] = df[c_name].str.strip()
                temp_df['색상'] = df[c_color].str.strip() if c_color else ''
                temp_df['평량'] = df[c_gram].str.replace(r'\\.0$', '', regex=True) if c_gram else '0'
                
                def extract_num(val):
                    res = re.sub(r'[^0-9.]', '', str(val))
                    return res if res and res != '.' else '0'
                
                temp_df['두께'] = df[c_thick].apply(extract_num) if c_thick else '0'
                temp_df['고시가'] = df[c_price].apply(lambda x: f"{int(float(extract_num(x))):,}" if float(extract_num(x)) > 0 else "0")
                temp_df['row_id'] = temp_df.apply(lambda r: f"id_{re.sub(r'[^a-zA-Z0-9]', '', r['품목']+r['평량']+r['sheet_name'])}", axis=1)
                
                combined_list.append(temp_df)
            
            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
        except Exception as e:
            print(f"Search Error: {e}")
    
    # 2. qq.xlsx 로드 (기존과 동일)
    final_boards = []
    if os.path.exists(qq_path):
        try:
            df_qq = pd.read_excel(qq_path).fillna('')
            for _, row in df_qq.iterrows():
                name_val = str(row.iloc[0]).strip()
                thick_raw = str(row.iloc[1])
                num_str = re.sub(r'[^0-9.]', '', thick_raw)
                if name_val and num_str and len(name_val) < 15:
                    try:
                        val = float(num_str)
                        if val > 0:
                            item = {'합지명': name_val, '두께': val}
                            if '1000' in name_val: 
                                final_boards.insert(0, item)
                            else: 
                                final_boards.append(item)
                    except: 
                        continue
        except: 
            pass
    
    if not final_boards or not any('1000' in b['합지명'] for b in final_boards):
        default_item = {'합지명': '1000g(기본)', '두께': 1.6}
        final_boards.insert(0, default_item)
    board_data = final_boards

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
        k_lower = keyword.lower()
        k_nospace = k_lower.replace(' ', '')
        for item in cached_data:
            # 전체 텍스트에서 검색 (띄어쓰기 무시 포함)
            full_text = f"{item['품목']} {item['색상']} {item['sheet_name']}".lower()
            full_nospace = full_text.replace(' ', '')
            
            if (k_lower in full_text or 
                k_nospace in full_nospace or
                k_nospace in item['search_nospace']):
                p, s = item['품목'], item['sheet_name']
                if '두성' in s:
                    item['url'] = f"https://www.doosungpaper.co.kr/goods/goods_search.php?keyword={p}"
                elif '삼원' in s:
                    item['url'] = f"https://www.samwonpaper.com/product/paper/list?search.searchString={p}"
                else:
                    item['url'] = f"https://www.google.com/search?q={s}+{p}"
                results.append(item)
    
    return render_template('index.html', results=results, keyword=keyword, 
                         authenticated=True, boards=board_data)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
