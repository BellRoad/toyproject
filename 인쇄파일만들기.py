import sqlite3
import openpyxl

# SQLite 데이터베이스 연결
conn = sqlite3.connect('카페5월5일매출.db')
cursor = conn.cursor()

# 상품리스트 테이블에서 데이터 가져오기
cursor.execute("SELECT 수량_셀, 금액_셀 FROM 상품리스트")
rows = cursor.fetchall()

# 데이터를 리스트에 저장
data_list = rows

# 데이터베이스 연결 종료
conn.close()

# 엑셀 파일 열기
wb = openpyxl.load_workbook('보고서출력.xlsx')
sheet = wb.active

# 리스트의 각 항목에 대해
for 수량_셀, 금액_셀 in data_list:
    # 엑셀에서 해당 셀의 값을 정수 0으로 변경
    if 수량_셀 and isinstance(수량_셀, str):
        sheet[수량_셀] = 0  # 정수 0으로 설정
    if 금액_셀 and isinstance(금액_셀, str):
        sheet[금액_셀] = 0  # 정수 0으로 설정

# 변경사항 저장
wb.save('보고서출력.xlsx')

print("작업이 완료되었습니다.")
print(f"총 {len(data_list)}개의 항목이 처리되었습니다.")
