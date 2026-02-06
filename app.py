import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # 보안을 위한 비밀키

# 엑셀 데이터 로드 함수
def load_data():
    file_path = 'search.xlsx'
    if not os.path.exists(file_path):
        return pd.DataFrame()
    
    # 엑셀 파일을 읽어옵니다.
    df = pd.read_excel(file_path)
    
    # '평량' 컬럼 빈칸을 위쪽 숫자로 채우기 (아까 해결한 로직)
    if '평량' in df.columns:
        df['평량'] = df['평량'].ffill()
        
    # 데이터 전처리: NaN 값들을 빈 문자열이나 적절한 값으로 변환
    df = df.fillna('')
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    # 비밀번호 인증 상태 확인 (세션 대신 간단한 파라미터 예시, 필요시 보강 가능)
    authenticated = True 
    
    keyword = ""
    results = []

    if request.method == 'POST':
        # 비밀번호 처리 로직이 필요하다면 여기에 추가
        password = request.form.get('password')
        
        # 검색어 가져오기
        keyword = request.form.get('keyword', '').strip()
        
        if keyword:
            df = load_data()
            if not df.empty:
                # '품목' 또는 '시트명' 컬럼에서 검색어 포함 여부 확인
                mask = df.apply(lambda row: keyword.lower() in str(row['품목']).lower() or 
                                            keyword.lower() in str(row['시트명']).lower(), axis=1)
                search_df = df[mask]
                
                # 계산기에서 사용할 '고시가_원본' 보존 및 리스트 변환
                search_df['고시가_원본'] = search_df['고시가']
                results = search_df.to_dict('records')
            
            if not results:
                flash('검색 결과가 없습니다.', 'info')

    return render_template('index.html', 
                           results=results, 
                           keyword=keyword, 
                           authenticated=authenticated)

if __name__ == '__main__':
    # Render 등 외부 배포 환경을 위한 설정
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)