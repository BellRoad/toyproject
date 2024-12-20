import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime, timedelta
import sqlite3

# 데이터베이스 파일명
DB_FILENAME = '카페5월5일매출.db'

# 기본 설정
root = tk.Tk()
root.title('카페 매출 입력')
font = ('맑은고딕', 11)

# 날짜 관련 변수 및 함수
current_date = datetime.now() - timedelta(days=1)
date_var = tk.StringVar(value=current_date.strftime('%Y-%m-%d'))

def update_date(days):
    global current_date
    current_date += timedelta(days=days)
    date_var.set(current_date.strftime('%Y-%m-%d'))

def format_date(event=None):
    raw_date = date_var.get()
    if len(raw_date) == 8 and raw_date.isdigit():
        formatted = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:]}"
        date_var.set(formatted)

# 상품 리스트 가져오기
def fetch_product_list():
    conn = sqlite3.connect(DB_FILENAME)
    cursor = conn.cursor()
    cursor.execute("SELECT 상품명 FROM 상품리스트 ORDER BY No ASC")
    products = [row[0] for row in cursor.fetchall()]
    conn.close()
    return products

product_names = fetch_product_list()

# 수식 평가 함수
def evaluate_expression(expression):
    try:
        return eval(re.sub(r'[^0-9+\-*/().]', '', expression))
    except Exception:
        return expression

# 날짜 입력 프레임
frame_date = ttk.Frame(root)
frame_date.pack(pady=10)

btn_prev = ttk.Button(frame_date, text='<<', command=lambda: update_date(-1))
btn_prev.pack(side='left')

entry_date = ttk.Entry(frame_date, textvariable=date_var, font=font, width=12)
entry_date.pack(side='left', padx=5)
entry_date.bind('<FocusOut>', format_date)

btn_next = ttk.Button(frame_date, text='>>', command=lambda: update_date(1))
btn_next.pack(side='left')

# 상품 리스트 프레임
frame_products = ttk.Frame(root)
frame_products.pack(pady=10)

ttk.Label(frame_products, text='품명', font=font).grid(row=0, column=0, padx=5)
ttk.Label(frame_products, text='수량', font=font).grid(row=0, column=1, padx=5)
ttk.Label(frame_products, text='금액', font=font).grid(row=0, column=2, padx=5)

entries = []
for i, product_name in enumerate(product_names, start=1):
    ttk.Label(frame_products, text=product_name, font=font).grid(row=i, column=0, padx=5)
    
    entry_quantity = ttk.Entry(frame_products, font=font, width=10)
    entry_quantity.grid(row=i, column=1, padx=5)
    
    entry_amount = ttk.Entry(frame_products, font=font, width=10)
    entry_amount.grid(row=i, column=2, padx=5)
    
    # 수식 평가 이벤트 추가
    def on_focus_out(event=None, entry_field=None):
        value = entry_field.get()
        entry_field.delete(0, tk.END)
        entry_field.insert(0, str(evaluate_expression(value)))
    
    entry_quantity.bind('<FocusOut>', lambda e, field=entry_quantity: on_focus_out(e, field))
    entry_amount.bind('<FocusOut>', lambda e, field=entry_amount: on_focus_out(e, field))
    
    entries.append((product_name, entry_quantity, entry_amount))

# DB에 데이터 삽입 함수
def save_to_db():
    conn = sqlite3.connect(DB_FILENAME)
    cursor = conn.cursor()
    
    date = date_var.get()
    
    for product_name, quantity_entry, amount_entry in entries:
        quantity = quantity_entry.get()
        amount = amount_entry.get()
        
        if quantity and amount:
            cursor.execute('''
                INSERT INTO 매출 (날짜, 상품ID, 수량, 금액) 
                VALUES (?, (SELECT id FROM 상품리스트 WHERE 상품명=?), ?, ?)
            ''', (date, product_name, int(quantity), int(amount)))
    
    conn.commit()
    conn.close()
    messagebox.showinfo("성공", "데이터가 성공적으로 저장되었습니다.")

# 입력 버튼
btn_save = ttk.Button(root, text='입력', command=save_to_db)
btn_save.pack(pady=10)

root.mainloop()
