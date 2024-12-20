import sqlite3

# 데이터베이스 파일명
DB_FILENAME = '카페5월5일매출.db'

# 매출 입력 함수
def input_sales():
    date = input("매출 날짜 (YYYY-MM-DD 형식): ")
    
    conn = sqlite3.connect(DB_FILENAME)
    cursor = conn.cursor()

    # 상품 리스트 조회
    cursor.execute("SELECT * FROM 상품리스트 ORDER BY No ASC")
    products = cursor.fetchall()

    for product in products:
        product_id = product[1]  # id
        product_name = product[2]  # 상품명

        quantity = int(input(f"{product_name}의 수량을 입력하세요 (정수): "))
        amount = int(input(f"{product_name}의 금액을 입력하세요 (정수): "))

        # 매출 데이터 삽입
        cursor.execute('''
            INSERT INTO 매출 (날짜, 상품ID, 수량, 금액) VALUES (?, ?, ?, ?)
        ''', (date, product_id, quantity, amount))

    conn.commit()  # 데이터베이스에 반영
    conn.close()   # 연결 종료
    print("매출 데이터가 성공적으로 입력되었습니다.")

# 프로그램 시작 시 매출 입력 함수 호출
if __name__ == "__main__":
    input_sales()
