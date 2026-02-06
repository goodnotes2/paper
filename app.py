import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # 보안을 위한 비밀키

# 엑셀 데이터 로드 함수
def load_data():
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        print("Error: search.xlsx 파일을 찾을 수 없습니다.") # 로그 확인용
        return pd.DataFrame()
    
    try:
        # 엑셀 파일을 읽어옵니다.
        df = pd.read_excel(file_path)
        
        # '평량' 컬럼 빈칸을 위쪽 숫자로 채우기
        if '평량' in df.columns:
            df['평량'] = df['평량'].ffill()
            
        # 데이터 전처리: NaN 값들을 빈 문자열로 변환
        df = df.fillna('')
        return df
    except Exception as e:
        print(f"Error loading excel: {e}")
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
                # 검색할 컬럼 리스트 (실제 엑셀에 있는 컬럼만 필터링)
                available_cols = [col for col in ['품목', '시트명'] if col in df.columns]
                
                if not available_cols:
                    flash('엑셀 파일에 "품목"이나 "시트명" 컬럼이 없습니다.', 'danger')
                else:
                    # 안전한 검색 로직: 존재하는 컬럼에서만 검색어 확인
                    def search_row(row):
                        for col in available_cols:
                            if keyword.lower() in str(row[col]).lower():
                                return True
                        return False

                    mask = df.apply(search_row, axis=1)
                    search_df = df[mask].copy() # 복사본 생성
                    
                    if not search_df.empty:
                        # '고시가' 컬럼이 있을 때만 원본 보존
                        if '고시가' in search_df.columns:
                            search_df['고시가_원본'] = search_df['고시가']
                        
                        results = search_df.to_dict('records')
            
            if not results:
                flash('검색 결과가 없습니다.', 'info')

    return render_template('index.html', 
                           results=results, 
                           keyword=keyword, 
                           authenticated=authenticated)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)