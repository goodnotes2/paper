import pandas as pd
import os
import re
from flask import Flask, render_template, request, session, redirect, url_for, jsonify

app = Flask(__name__)
app.secret_key = 'expert_version_v20'

SITE_PASSWORD = "03877"

sheets = ['두성', '삼원', '한국', '무림', '삼화', '서경', '한솔', '전주']

company_urls = {
    '두성': 'https://www.doosungpaper.co.kr/goods/goods_search.php?keyword=',
    '삼원': 'https://www.samwonpaper.com/product/paper/list?search.searchString=',
    '한국': 'https://www.hankukpaper.com/ko/product/listProductinfo.do',
    '무림': 'https://www.moorim.co.kr:13002/product/paper.php',
    '삼화': 'https://www.samwhapaper.com/product/samwhapaper/all?keyword=',
    '서경': 'https://wedesignpaper.com/search?type=shopping&sort=consensus_desc&keyword=',
    '한솔': 'https://www.hansolpaper.co.kr/product/insper',
    '전주': 'https://jeonjupaper.com/publicationpaper'
}

cached_data = []

board_data = [
    {'합지명': '1000g(기본)', '두께': 1.6, '각양장_앞뒤': 3.0, '미소_앞뒤': 2.0, 'bleed': 17},
    {'합지명': '1100g',       '두께': 1.6, '각양장_앞뒤': 3.0, '미소_앞뒤': 2.0, 'bleed': 17},
    {'합지명': '1200g',       '두께': 1.6, '각양장_앞뒤': 3.5, '미소_앞뒤': 2.5, 'bleed': 17},
    {'합지명': '1300g',       '두께': 1.6, '각양장_앞뒤': 4.0, '미소_앞뒤': 3.0, 'bleed': 18},
    {'합지명': '1400g',       '두께': 1.6, '각양장_앞뒤': 4.5, '미소_앞뒤': 3.5, 'bleed': 18},
    {'합지명': '1500g',       '두께': 1.6, '각양장_앞뒤': 5.0, '미소_앞뒤': 4.0, 'bleed': 18},
]

def load_data():
    global cached_data
    file_path = 'search.xlsx'
    if os.path.exists(file_path):
        try:
            combined_list = []
            for sheet in sheets:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
                except Exception:
                    continue

                df.columns = [str(col).strip() for col in df.columns]

                c_name  = next((c for c in df.columns if any(k in c for k in ['품목', '품명'])), None)
                c_thick = next((c for c in df.columns if any(k in c for k in ['두께', 'μm', 'um'])), None)
                c_gram  = next((c for c in df.columns if any(k in c for k in ['평량', 'g'])), None)
                c_color = next((c for c in df.columns if any(k in c for k in ['색상', '컬러', '패턴'])), None)
                c_price = next((c for c in df.columns if any(k in c for k in ['고시가', '단가'])), None)
                c_note  = next((c for c in df.columns if any(k in c for k in ['비고', '메모'])), None)
                c_size  = next((c for c in df.columns if any(k in c for k in ['사이즈', '규격', '크기'])), None)

                if not c_name:
                    continue

                df[c_name] = df[c_name].astype(str).str.replace('\\n', ' ', regex=False).str.replace('\\r', ' ', regex=False)
                if c_color: df[c_color] = df[c_color].astype(str).str.replace('\\n', ' ', regex=False)
                if c_note:  df[c_note]  = df[c_note].astype(str).str.replace('\\n', ' ', regex=False)
                if c_size:  df[c_size]  = df[c_size].astype(str).str.replace('\\n', ' ', regex=False)

                df[c_name] = df[c_name].replace('nan', pd.NA).ffill().fillna('')
                if c_size:  df[c_size]  = df[c_size].replace('nan', pd.NA).ffill().fillna('')
                if c_gram:  df[c_gram]  = df[c_gram].replace('nan', pd.NA).ffill().fillna('')
                if c_color: df[c_color] = df[c_color].replace('nan', pd.NA).ffill().fillna('')

                def extract_num(val):
                    res = re.sub(r'[^0-9.]', '', str(val))
                    return res if res and res != '.' else '0'

                temp_df = pd.DataFrame(index=df.index)
                temp_df['품목']  = df[c_name].str.strip()
                temp_df['색상']  = df[c_color].str.strip() if c_color else ''
                temp_df['비고']  = df[c_note].str.strip()  if c_note  else ''
                temp_df['사이즈'] = df[c_size].str.strip()  if c_size  else ''
                temp_df['평량']  = df[c_gram].astype(str).str.replace(r'\.0$', '', regex=True) if c_gram else '0'
                temp_df['두께']  = df[c_thick].apply(extract_num) if c_thick else '0'
                temp_df['고시가'] = df[c_price].apply(
                    lambda x: f"{int(float(extract_num(x))):,}" if float(extract_num(x)) > 0 else "0"
                ) if c_price else '0'
                temp_df['시트명'] = sheet

                all_text = (temp_df['품목'] + ' ' + temp_df['색상'] + ' ' + temp_df['비고'] + ' ' + temp_df['사이즈']).str.lower()
                temp_df['search_full']    = all_text
                temp_df['search_nospace'] = all_text.str.replace(' ', '', regex=False)

                combined_list.append(temp_df)

            if combined_list:
                cached_data = pd.concat(combined_list, ignore_index=True).to_dict('records')
                print(f"[INFO] 총 {len(cached_data)}행 로드 완료")
        except Exception as e:
            print(f"[ERROR] load_data: {e}")

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
        k_nospace = k.replace(' ', '')
        row_counter = 0

        for item in cached_data:
            full    = item.get('search_full', '')
            nospace = item.get('search_nospace', '')

            if k in full or k_nospace in nospace:
                s = item['시트명']
                base_url = company_urls.get(s, '#')
                if s in ['두성', '삼원', '삼화', '서경']:
                    item['url'] = base_url + keyword
                else:
                    item['url'] = base_url
                item['row_id'] = row_counter
                row_counter += 1
                results.append(item)

    return render_template('index.html',
                           results=results,
                           keyword=keyword,
                           authenticated=True,
                           boards=board_data)

@app.route('/calculate_seneca_api', methods=['POST'])
def calculate_seneca_api():
    if not session.get('authenticated'):
        return jsonify({'error': 'Unauthorized'}), 401
    data = request.get_json()
    try:
        pc = float(data.get('page_count', 0))
        t  = float(data.get('thickness', 0))
        if t == 0:
            return jsonify({'error': '두께는 0이 될 수 없습니다.'}), 400
        result = (pc / 2) * t / 1000
        return jsonify({'result': f"{result:,.1f}"})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)