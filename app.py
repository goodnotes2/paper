import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for, jsonify
import os
import re
from urllib.parse import quote

app = Flask(__name__)
app.secret_key = 'expert_version_v20'

SITE_PASSWORD = "03877"
cached_data = []

# ─────────────────────────────────────────
# 합지 데이터 하드코딩 (qq.xlsx 대체)
# 표지두께: 합지두께(mm), 각양장 앞뒤(mm), 환양장미소 앞뒤(mm)
# ─────────────────────────────────────────
BOARD_DATA = [
    {'합지명': '1000g', '두께': 1.6, '각양장_앞뒤': 3.0,  '미소_앞뒤': 2.0},
    {'합지명': '1100g', '두께': 1.8, '각양장_앞뒤': 3.0,  '미소_앞뒤': 2.0},
    {'합지명': '1200g', '두께': 2.0, '각양장_앞뒤': 3.5,  '미소_앞뒤': 2.5},
    {'합지명': '1300g', '두께': 2.2, '각양장_앞뒤': 4.0,  '미소_앞뒤': 3.0},
    {'합지명': '1400g', '두께': 2.4, '각양장_앞뒤': 4.5,  '미소_앞뒤': 3.5},
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

# ─────────────────────────────────────────
# 환양장(인조있음) 세네카 계산 함수
# ─────────────────────────────────────────
def calc_injo_seneca(pages):
    if pages >= 40:
        return None  # 오류
    elif pages >= 35:
        return pages + 9
    elif pages >= 25:
        return pages + 8
    elif pages >= 15:
        return pages + 7
    elif pages >= 10:
        return pages + 6
    elif pages >= 1:
        return pages + 5
    else:
        return pages

# ─────────────────────────────────────────
# 세네카(척추) 두께 계산 API
# ─────────────────────────────────────────
@app.route('/calc_spine', methods=['POST'])
def calc_spine():
    data       = request.get_json()
    bind_type  = data.get('bind_type', '')   # 'miso' / 'gak' / 'injо'
    board_name = data.get('board', '')        # '1000g' 등
    pages      = int(data.get('pages', 0))
    thickness_um = float(data.get('thickness_um', 0))  # μm 단위

    # μm → mm
    thickness_mm = thickness_um / 1000

    # 내지 두께 (mm)
    naeji_mm = round(thickness_mm * pages, 2)

    board = next((b for b in BOARD_DATA if b['합지명'] == board_name), BOARD_DATA[0])

    if bind_type == 'miso':
        # 환양장미소: 세네카 = 내지두께, 앞뒤 = 미소_앞뒤
        seneca   = naeji_mm
        front_back = board['미소_앞뒤']
    elif bind_type == 'gak':
        # 각양장: 세네카 = 내지두께, 앞뒤 = 각양장_앞뒤
        seneca   = naeji_mm
        front_back = board['각양장_앞뒤']
    elif bind_type == 'injо':
        # 환양장(인조있음): 합지두께 고려 안함, 세네카는 별도 계산식
        result = calc_injo_seneca(pages)
        if result is None:
            return jsonify({'error': '페이지수 40 이상은 오류입니다'})
        seneca     = result
        front_back = 0  # 인조있음은 앞뒤 별도 없음
    else:
        return jsonify({'error': '제본 방식 오류'})

    return jsonify({
        'seneca':      round(seneca, 2),
        'front_back':  front_back,
        'board_thick': board['두께'],
        'naeji_mm':    naeji_mm,
    })

# ─────────────────────────────────────────
# 메인 라우트
# ─────────────────────────────────────────
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
