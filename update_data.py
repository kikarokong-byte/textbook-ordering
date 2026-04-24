import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import sys
from tqdm import tqdm
import concurrent.futures
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
# ==========================================
# ⚙️ ตั้งค่าพื้นฐาน
# ==========================================
BASE_URL = 'http://202.29.173.190/textbook/web/index.php'
BASE_IMG = 'http://202.29.173.190/textbook/web/'
DELAY = 0.8  # หน่วงเวลา 0.8 วินาทีต่อหน้า

# รหัสบัญชีหนังสือ (bookmain) ที่ต้องดึงทั้งหมด
ACCOUNT_CODES = ['11', '12', '21', '22', '31', '32']

def code_to_label(code):
    return f'{code[0]}.{code[1]}'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'th,en;q=0.9',
    'Referer': 'http://202.29.173.190/textbook/web/index.php',
}

def get_page_html(page_num, session, bookmain='11,12,21,22,31,32'):
    params = {
        'ispage': page_num,
        'bookmain': bookmain,
        'bookgroup': '', 'class': '', 'bookprint': '',
        'bookcategory': '', 'name': '', 'bookeditor': '',
        'id_round': '', 'chksearch': 'true'
    }
    r = session.get(BASE_URL, params=params, headers=HEADERS, timeout=30)
    # โค้ดต้นฉบับใช้ utf-8 (หากภาษาไทยเพี้ยน สามารถเปลี่ยนเป็น 'windows-874' หรือ 'tis-620' ได้)
    r.encoding = 'utf-8'
    return r.text

def detect_total_pages(html):
    """ ตรวจจำนวนหน้ารวมจากข้อความ 'พบจำนวน X รายการ' (10 รายการ/หน้า) """
    m = re.search(r'พบจำนวน\s*([\d,]+)\s*รายการ', html)
    if m:
        total = int(m.group(1).replace(',', ''))
        return max(1, (total + 9) // 10)
    return None

def parse_book_block(block_text, img_src, links):
    """ แยกข้อมูลแต่ละ field จาก text block """
    lines = [l.strip() for l in block_text.split('\n') if l.strip()]
    
    MEDIA_TYPES = {'หนังสือเรียน', 'แบบฝึกหัด', 'แบบฝึกทักษะ',
                   'คู่มือครู', 'คู่มือครูและแผนการจัดการเรียนรู้',
                   'แบบฝึกไวยากรณ์', 'หนังสืออ่านประกอบ', 'ซีดี'}
    
    FIELD_KEYS = {
        'รายวิชา': 'รายวิชา', 'กลุ่มสาระการเรียนรู้': 'กลุ่มสาระการเรียนรู้',
        'ชั้น': 'ชั้น', 'ผู้จัดพิมพ์': 'ผู้จัดพิมพ์', 'ผู้เรียบเรียง': 'ผู้เรียบเรียง',
        'ปี พ.ศ. ที่เผยแพร่': 'ปีที่เผยแพร่', 'ขนาด': 'ขนาด', 'จำนวนหน้า': 'จำนวนหน้า',
        'กระดาษ': 'กระดาษ', 'พิมพ์': 'พิมพ์', 'น้ำหนัก': 'น้ำหนัก',
    }
    
    book = {col: '' for col in
            ['บัญชี','ประเภทสื่อ','ชื่อหนังสือ','รายวิชา','กลุ่มสาระการเรียนรู้',
             'ชั้น','ผู้จัดพิมพ์','ผู้เรียบเรียง','ปีที่เผยแพร่',
             'ขนาด','จำนวนหน้า','กระดาษ','พิมพ์','น้ำหนัก','ราคา',
             'URL_รูปภาพ','URL_ใบประกาศ','URL_ตัวอย่างเนื้อหา']}
    
    book['URL_รูปภาพ'] = BASE_IMG + img_src if img_src else ''
    
    for href, text in links:
        if not href: continue
        full = BASE_IMG + href if not href.startswith('http') else href
        if any(k in text for k in ['ใบประกาศ', 'ใบประกัน', 'ใบอนุญาต']):
            book['URL_ใบประกาศ'] = full
        elif 'ตัวอย่าง' in text:
            book['URL_ตัวอย่างเนื้อหา'] = full
            
    i = 0
    found_type = False
    while i < len(lines):
        line = lines[i]
        # ประเภทสื่อ
        if not found_type and line in MEDIA_TYPES:
            book['ประเภทสื่อ'] = line
            found_type = True
            if i+1 < len(lines) and lines[i+1] not in FIELD_KEYS and lines[i+1] not in MEDIA_TYPES:
                book['ชื่อหนังสือ'] = lines[i+1]
                i += 2
                continue
                
        # Field key-value pairs
        if line in FIELD_KEYS and i+1 < len(lines):
            book[FIELD_KEYS[line]] = lines[i+1]
            i += 2
            continue
            
        # ราคา
        m = re.search(r'ราคา\s+([\d,]+(?:\.\d+)?)\s*บาท', line)
        if m:
            book['ราคา'] = m.group(1)
        i += 1
        
    return book

def parse_page(html):
    """ parse หน้าเว็บและคืน list ของหนังสือ """
    soup = BeautifulSoup(html, 'lxml')
    books = []
    
    book_imgs = soup.find_all('img', src=re.compile(r'images/book/.*_image'))
    
    for img in book_imgs:
        img_src = img.get('src', '')
        container = None
        for tag in ['tr', 'table', 'div']:
            container = img.find_parent(tag)
            if container:
                text = container.get_text(separator='\n')
                if ('ผู้จัดพิมพ์' in text or 'ราคา' in text) and len(text) > 100:
                    break
                container = None
                
        if not container:
            el = img
            for _ in range(10):
                el = el.parent
                if el is None: break
                t = el.get_text()
                if 'ผู้จัดพิมพ์' in t and 'ราคา' in t:
                    container = el
                    break
                    
        if not container: continue
        
        links = [(a.get('href',''), a.get_text(strip=True)) for a in container.find_all('a', href=True)]
        block_text = container.get_text(separator='\n')
        book = parse_book_block(block_text, img_src, links)
        
        if book.get('ชื่อหนังสือ') or book.get('ประเภทสื่อ'):
            books.append(book)
            
    return books

# ==========================================
# 🚀 เริ่มการทำงานหลัก (Main execution)
# ==========================================
if __name__ == '__main__':
    session = requests.Session()

    # ตั้งค่าระบบ Auto-Retry หากเน็ตหลุดหรือเซิร์ฟเวอร์ล่มชั่วคราว
    retry_strategy = Retry(
        total=3,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
    )
    adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=20, pool_maxsize=20)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    print('✅ Setup เรียบร้อย กำลังตรวจจำนวนหน้าของแต่ละบัญชี...')

    # ตรวจจำนวนหน้าของแต่ละบัญชี
    pages_per_account = {}
    for code in ACCOUNT_CODES:
        html = get_page_html(1, session, bookmain=code)
        total = detect_total_pages(html)
        if total is None:
            # fallback: ลอง parse ดูว่ามีรายการไหม
            total = 1 if parse_page(html) else 0
        pages_per_account[code] = total
        print(f'  บัญชี {code_to_label(code)}: {total} หน้า')

    total_tasks = sum(pages_per_account.values())
    print(f'\n🚀 เริ่ม scrape รวม {total_tasks} หน้า (10 threads)')

    def scrape_single_page(code, p_num):
        try:
            h = get_page_html(p_num, session, bookmain=code)
            bks = parse_page(h)
            for b in bks:
                b['บัญชี'] = code_to_label(code)
            return code, p_num, bks, None
        except Exception as ex:
            return code, p_num, None, ex

    all_books = []
    failed_pages = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(scrape_single_page, code, p): (code, p)
            for code, n in pages_per_account.items()
            for p in range(1, n + 1)
        }

        for future in tqdm(concurrent.futures.as_completed(futures), total=total_tasks, desc='Scraping', unit='page'):
            code, p_num, bks, ex = future.result()
            if ex:
                tqdm.write(f'  ⚠️ บัญชี {code_to_label(code)} หน้า {p_num} error: {ex}')
                failed_pages.append((code, p_num))
            elif bks is not None:
                all_books.extend(bks)

    print(f'\n✅ ดึงข้อมูลเสร็จ! รวม {len(all_books)} รายการ (ก่อนรวมซ้ำ)')

    if failed_pages:
        print(f'หน้าที่ล้มเหลว: {failed_pages}')

    # สร้าง DataFrame และรวมรายการซ้ำ (หนังสือเล่มเดียวกันแต่อยู่หลายบัญชี)
    if all_books:
        df = pd.DataFrame(all_books)

        # key สำหรับ dedupe: ใช้ URL_รูปภาพ (unique per book) ถ้ามี; fallback เป็น combo
        def dedupe_key(row):
            if row.get('URL_รูปภาพ'):
                return row['URL_รูปภาพ']
            return (row.get('ชื่อหนังสือ',''), row.get('ผู้จัดพิมพ์',''),
                    row.get('ชั้น',''), row.get('ประเภทสื่อ',''))

        df['_key'] = df.apply(dedupe_key, axis=1)

        # เรียงบัญชีให้สวย (1.1, 1.2, 2.1, ...) แล้ว group
        account_order = [code_to_label(c) for c in ACCOUNT_CODES]
        def join_accounts(series):
            uniq = sorted(set(series.dropna().astype(str)),
                          key=lambda x: account_order.index(x) if x in account_order else 999)
            return ','.join(uniq)

        # เอาแถวแรกของแต่ละ key (ทุก field ของหนังสือเหมือนกัน) แล้วแทน บัญชี ด้วยค่ารวม
        merged_accounts = df.groupby('_key')['บัญชี'].apply(join_accounts)
        df_clean = df.drop_duplicates('_key', keep='first').copy()
        df_clean['บัญชี'] = df_clean['_key'].map(merged_accounts)
        df_clean = df_clean.drop(columns=['_key']).reset_index(drop=True)
        df_clean.index += 1

        print(f'จำนวนหนังสือหลังรวมซ้ำ: {len(df_clean):,} รายการ')

        # คลีนช่องราคาให้เป็นตัวเลขล้วน สำหรับนำไปคำนวณในระบบค้นหาหลัก
        df_clean['ราคา'] = df_clean['ราคา'].astype(str).str.replace('บาท', '').str.replace(',', '').str.strip()
        
        # บันทึกเป็นไฟล์ Excel 
        output_file = 'textbooks.xlsx'
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_clean.to_excel(writer, sheet_name='หนังสือทั้งหมด', index_label='ลำดับ')
                
                # ปรับความกว้างคอลัมน์อัตโนมัติ
                ws = writer.sheets['หนังสือทั้งหมด']
                for col in ws.columns:
                    max_w = max((len(str(c.value or '')) for c in col), default=10)
                    ws.column_dimensions[col[0].column_letter].width = min(max_w + 2, 50)
                    
            print(f'💾 บันทึกไฟล์สำเร็จ: {output_file} พร้อมใช้งาน!')
            
        except PermissionError:
            print(f'\n❌ บันทึกไม่สำเร็จ: สิทธิ์ถูกปฏิเสธ (Permission denied)')
            print(f'💡 สาเหตุ: คุณกำลังเปิดไฟล์ "{output_file}" ค้างไว้ในโปรแกรม Excel')
            print(f'🛠️ วิธีแก้: กรุณาปิดโปรแกรม Excel ที่เปิดไฟล์นี้อยู่ แล้วกดรันอัปเดตใหม่อีกครั้งครับ')