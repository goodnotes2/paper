import pandas as pd
from flask import Flask, render_template, request
import os

app = Flask(__name__)
app.secret_key = 'paper_search_key'

def load_data():
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return pd.DataFrame()
    
    try:
        # 엑셀의 모든 시트(두성, 삼원, 무림 등)를 한 번에 읽기
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        
        combined_df = []
        for sheet_name, df in all_sheets.items():
            # 컬럼명 공백 제거
            df.columns = [str(col).strip() for col in df.columns]
            
            # 병합된 셀(빈칸)을 위 데이터로 자동 채우기 (ffill)
            # 품목명이 한 줄에만 있어도 아래 줄까지 다 검색되게 합니다.
            fill_cols = ['품목', '사이즈', '평량']
            for col in fill_cols:
                if col in df.columns:
                    df[col] = df[col].ffill()
            
            # 시트명(제지사 이름) 저장
            df['시트명'] = sheet_name
            combined_df.append(df)
            
        return pd.concat(combined_df, ignore_index=True).fillna('')
    except Exception as e:
        print(f"Error: {e}")
        return pd.DataFrame()

@app.route('/', methods=['GET', 'POST'])
def index():
    keyword = ""
    results = []
    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        if keyword:
            df = load_data()
            if not df.empty:
                # 대소문자 무시하고 모든 칸에서 키워드 포함 여부 검색
                mask = df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)
                results = df[mask].to_dict('records')
    
    return render_template('index.html', results=results, keyword=keyword)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)