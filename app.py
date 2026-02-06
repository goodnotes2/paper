import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

def load_data():
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return pd.DataFrame()
    
    try:
        # 엔진을 openpyxl로 명시하여 엑셀 로드 안정성 강화
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # 모든 컬럼명의 앞뒤 공백 제거 (매우 중요!)
        df.columns = [str(col).strip() for col in df.columns]
        
        if '평량' in df.columns:
            df['평량'] = df['평량'].ffill()
            
        df = df.fillna('')
        return df
    except Exception as e:
        print(f"Excel Load Error: {e}")
        return pd.DataFrame()

@app.route('/', methods=['GET', 'POST'])
def index():
    authenticated = True 
    keyword = ""
    results = []

    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        
        if keyword:
            df = load_data()
            if not df.empty:
                # 유연한 컬럼 매칭: '품목'이나 '시트'라는 글자가 들어간 모든 컬럼 찾기
                target_cols = [col for col in df.columns if '품목' in col or '시트' in col]
                
                # 만약 지정된 컬럼이 없으면 전체 컬럼에서 검색 시도
                if not target_cols:
                    target_cols = df.columns.tolist()

                def search_row(row):
                    for col in target_cols:
                        if keyword.lower() in str(row[col]).lower():
                            return True
                    return False

                mask = df.apply(search_row, axis=1)
                search_df = df[mask].copy()
                
                if not search_df.empty:
                    if '고시가' in search_df.columns:
                        search_df['고시가_원본'] = search_df['고시가']
                    results = search_df.to_dict('records')
            
            if not results:
                flash(f"'{keyword}'에 대한 검색 결과가 없습니다. (확인된 컬럼: {df.columns.tolist()})", 'info')

    return render_template('index.html', 
                           results=results, 
                           keyword=keyword, 
                           authenticated=authenticated)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)