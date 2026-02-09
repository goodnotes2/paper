import pandas as pd
from flask import Flask, render_template, request, flash
import os

app = Flask(__name__)
app.secret_key = 'paper_search_key'

def load_data():
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return pd.DataFrame()
    
    try:
        # 📌 모든 시트를 읽어오기 (두성, 삼원, 한국 등 전체 검색 가능)
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        
        combined_df = []
        for sheet_name, df in all_sheets.items():
            # 컬럼명 공백 제거
            df.columns = [str(col).strip() for col in df.columns]
            
            # 📌 병합된 셀(빈칸)을 위쪽 데이터로 채우기 (이미지의 E-보드 문제 해결)
            cols_to_fill = ['품목', '사이즈', '평량']
            for col in cols_to_fill:
                if col in df.columns:
                    df[col] = df[col].ffill()
            
            # 어느 시트 데이터인지 표시 (선택 사항)
            df['시트명'] = sheet_name
            combined_df.append(df)
            
        final_df = pd.concat(combined_df, ignore_index=True)
        return final_df.fillna('')
    except Exception as e:
        print(f"Excel Load Error: {e}")
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
                # 📌 대소문자 구분 없이 모든 열에서 키워드 검색
                mask = df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)
                results = df[mask].to_dict('records')
            
            if not results:
                flash(f"'{keyword}'에 대한 결과가 없습니다.", 'info')

    return render_template('index.html', results=results, keyword=keyword)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)