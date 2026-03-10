import pandas as pd

try:
    df = pd.read_excel('/home/sssssmmm/safian/베이비코 세피앙 주문서(26.03.02).xlsx')
    print("--- 엑셀 파일 로드 성공 ---")
    print("컬럼 목록:", df.columns.tolist())
    print("데이터 미리보기:\n", df.head(3).to_string())
except Exception as e:
    print("엑셀 읽기 실패:", e)
