"""
db.py - ระบบจัดการฐานข้อมูลผ่าน Google Sheets
รองรับทั้งโหมด Local (ใช้ SQLite) และโหมด Cloud (ใช้ Google Sheets)
ระบบเลือกโหมดอัตโนมัติจาก st.secrets
"""
import pandas as pd
import os
import datetime

# ตรวจสอบว่ามี Google Sheets credentials หรือไม่
def _is_cloud_mode():
    try:
        import streamlit as st
        return "gcp_service_account" in st.secrets
    except:
        return False

# ======================================================
# โหมด Google Sheets (Cloud)
# ======================================================
def _get_gspread_client():
    import streamlit as st
    import gspread
    from google.oauth2.service_account import Credentials
    
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return gspread.authorize(creds)

def _get_sheet(sheet_name):
    import streamlit as st
    client = _get_gspread_client()
    # หา spreadsheet_url ได้ทั้ง top-level และใน gcp_service_account
    try:
        spreadsheet_url = st.secrets["spreadsheet_url"]
    except KeyError:
        try:
            spreadsheet_url = st.secrets["gcp_service_account"]["spreadsheet_url"]
        except KeyError:
            spreadsheet_url = st.secrets["sheets"]["spreadsheet_url"]
    sh = client.open_by_url(spreadsheet_url)
    return sh.worksheet(sheet_name)

@st.cache_data(show_spinner=False, ttl=60)
def _load_gsheet_as_df(sheet_name):
    ws = _get_sheet(sheet_name)
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame()

# ======================================================
# โหมด SQLite (Local)
# ======================================================
import sqlite3

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "school_orders.db")

def _get_sqlite_connection():
    return sqlite3.connect(DB_FILE)

DEFAULT_CLASSES = [
    'อนุบาลปีที่ 1', 'อนุบาลปีที่ 2', 'อนุบาลปีที่ 3',
    'ประถมศึกษาปีที่ 1', 'ประถมศึกษาปีที่ 2', 'ประถมศึกษาปีที่ 3',
    'ประถมศึกษาปีที่ 4', 'ประถมศึกษาปีที่ 5', 'ประถมศึกษาปีที่ 6',
    'มัธยมศึกษาปีที่ 1', 'มัธยมศึกษาปีที่ 2', 'มัธยมศึกษาปีที่ 3',
    'มัธยมศึกษาปีที่ 4', 'มัธยมศึกษาปีที่ 5', 'มัธยมศึกษาปีที่ 6'
]

# ======================================================
# ฟังก์ชัน Public (ทำงานได้ทั้ง 2 โหมดโดยอัตโนมัติ)
# ======================================================

def init_db():
    if _is_cloud_mode():
        # Cloud mode: Google Sheets ถูกสร้างไว้แล้ว ไม่ต้องทำอะไร
        return
    
    # Local mode: SQLite
    conn = _get_sqlite_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS budgets (
                    class_name TEXT PRIMARY KEY, 
                    budget_limit REAL
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS orders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    teacher_name TEXT,
                    class_name TEXT,
                    book_name TEXT,
                    publisher TEXT,
                    unit_price REAL,
                    quantity INTEGER,
                    total_price REAL,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )''')
    conn.commit()
    c.execute("SELECT COUNT(*) FROM budgets")
    if c.fetchone()[0] == 0:
        for cls in DEFAULT_CLASSES:
            c.execute("INSERT INTO budgets (class_name, budget_limit) VALUES (?, ?)", (cls, 50000.0))
        conn.commit()
    conn.close()


def load_budgets():
    if _is_cloud_mode():
        df = _load_gsheet_as_df("budgets")
        if df.empty:
            return {cls: 50000.0 for cls in DEFAULT_CLASSES}
        return dict(zip(df['class_name'], df['budget_limit'].astype(float)))
    
    conn = _get_sqlite_connection()
    df = pd.read_sql("SELECT * FROM budgets", conn)
    conn.close()
    return dict(zip(df['class_name'], df['budget_limit']))


def load_budgets_df():
    if _is_cloud_mode():
        df = _load_gsheet_as_df("budgets")
        if df.empty:
            return pd.DataFrame({'class_name': DEFAULT_CLASSES, 'budget_limit': [50000.0]*len(DEFAULT_CLASSES)})
        return df
    
    conn = _get_sqlite_connection()
    df = pd.read_sql("SELECT * FROM budgets", conn)
    conn.close()
    return df


def save_budgets(df):
    if _is_cloud_mode():
        import gspread
        ws = _get_sheet("budgets")
        ws.clear()
        ws.update([df.columns.tolist()] + df.fillna('').values.tolist())
        return
    
    conn = _get_sqlite_connection()
    c = conn.cursor()
    c.execute("DROP TABLE IF EXISTS budgets")
    c.execute("CREATE TABLE budgets (class_name TEXT PRIMARY KEY, budget_limit REAL)")
    df.to_sql('budgets', conn, if_exists='append', index=False)
    conn.commit()
    conn.close()


def save_order(teacher_name, class_name, cart_items):
    valid_items = []
    order_total = 0.0
    for item in cart_items:
        qty = int(item['qty'])
        if qty > 0:
            order_total += float(item['price']) * qty
            valid_items.append(item)
    
    if not valid_items:
        return False, "ไม่มีรายการหนังสือสั่งซื้อ"
    
    if _is_cloud_mode():
        ws = _get_sheet("orders")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # หา next id
        existing = ws.get_all_values()
        next_id = len(existing)  # header + rows
        
        rows_to_add = []
        for item in valid_items:
            qty = int(item['qty'])
            total = float(item['price']) * qty
            rows_to_add.append([
                next_id, teacher_name, class_name,
                item['book_name'], item['publisher'],
                float(item['price']), qty, total, timestamp
            ])
            next_id += 1
        ws.append_rows(rows_to_add)
        _load_gsheet_as_df.clear()  # ล้าง cache เพื่อให้ยอดอัพเดตทันที
        return True, "บันทึกสำเร็จ"
    
    conn = _get_sqlite_connection()
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in valid_items:
        qty = int(item['qty'])
        total = float(item['price']) * qty
        c.execute('''INSERT INTO orders (teacher_name, class_name, book_name, publisher, unit_price, quantity, total_price, timestamp)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                  (teacher_name, class_name, item['book_name'], item['publisher'], float(item['price']), qty, total, timestamp))
    conn.commit()
    conn.close()
    return True, "บันทึกสำเร็จ"


def load_orders():
    if _is_cloud_mode():
        df = _load_gsheet_as_df("orders")
        if df.empty:
            return pd.DataFrame(columns=['id','teacher_name','class_name','book_name','publisher','unit_price','quantity','total_price','timestamp'])
        for col in ['unit_price', 'total_price']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        if 'quantity' in df.columns:
            df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce').fillna(0).astype(int)
        return df
    
    conn = _get_sqlite_connection()
    df = pd.read_sql("SELECT * FROM orders ORDER BY timestamp DESC", conn)
    conn.close()
    return df


def get_submitted_total(class_name):
    df = load_orders()
    if df.empty or 'class_name' not in df.columns:
        return 0.0
    filtered = df[df['class_name'] == class_name]
    if filtered.empty:
        return 0.0
    return float(filtered['total_price'].sum())


def clear_orders():
    if _is_cloud_mode():
        ws = _get_sheet("orders")
        # เก็บแถว header ไว้
        headers = ws.row_values(1)
        ws.clear()
        ws.append_row(headers)
        _load_gsheet_as_df.clear()  # ล้าง cache
        return
    
    conn = _get_sqlite_connection()
    c = conn.cursor()
    c.execute("DELETE FROM orders")
    conn.commit()
    conn.close()


def sync_orders(edited_df, original_df):
    if _is_cloud_mode():
        # Cloud mode: เขียนทับทั้งหมดตาม edited_df
        ws = _get_sheet("orders")
        headers = ws.row_values(1)
        ws.clear()
        ws.append_row(headers)
        if not edited_df.empty:
            rows = edited_df.fillna('').values.tolist()
            ws.append_rows(rows)
        return True, "แก้ไขสำเร็จ"
    
    # Local SQLite mode
    conn = _get_sqlite_connection()
    c = conn.cursor()
    
    if 'id' not in original_df.columns:
        conn.close()
        return False, "Data lacking reference IDs"
    
    orig_ids = set(original_df['id'].dropna().astype(int).tolist())
    curr_ids = set(edited_df['id'].dropna().astype(int).tolist()) if 'id' in edited_df.columns else set()
    
    for d_id in orig_ids - curr_ids:
        c.execute("DELETE FROM orders WHERE id = ?", (int(d_id),))
    
    for _, row in edited_df.iterrows():
        row_id = row.get('id')
        try:
            qty = int(row.get('quantity', 0))
            unit = float(row.get('unit_price', 0.0))
        except:
            qty, unit = 0, 0.0
        total_p = qty * unit
        t_name = row.get('teacher_name', 'Admin แก้ไข')
        c_name = row.get('class_name', '-')
        b_name = row.get('book_name', '-')
        pub = row.get('publisher', '-')
        
        if pd.isna(row_id) or str(row_id).strip() == "":
            c.execute('''INSERT INTO orders (teacher_name, class_name, book_name, publisher, unit_price, quantity, total_price)
                         VALUES (?, ?, ?, ?, ?, ?, ?)''', (t_name, c_name, b_name, pub, unit, qty, total_p))
        else:
            c.execute('''UPDATE orders SET teacher_name=?, class_name=?, book_name=?,
                         publisher=?, unit_price=?, quantity=?, total_price=?
                         WHERE id=?''', (t_name, c_name, b_name, pub, unit, qty, total_p, int(row_id)))
    
    conn.commit()
    conn.close()
    return True, "แก้ไขสำเร็จ"
