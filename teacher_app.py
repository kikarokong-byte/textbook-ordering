import streamlit as st
import pandas as pd
import os
import db
import requests
from io import BytesIO

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_image(url):
    """ดึงรูปผ่าน Server-side เพื่อหลีกเลี่ยงการ Block จากเซิร์ฟเวอร์สำนักพิมพ์"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Referer': 'https://www.aksorn.com/',
            'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8'
        }
        r = requests.get(url, headers=headers, timeout=5)
        if r.status_code == 200 and 'image' in r.headers.get('Content-Type', ''):
            return BytesIO(r.content)
    except:
        pass
    return None

st.set_page_config(page_title="ระบบสั่งหนังสือคุณครู", page_icon="🛍️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }
    .title-col { font-weight: 600; font-size: 1.1em; color: #1e3a8a; }
    .price-col { color: #e11d48; font-weight: bold; }
    .stNumberInput input { text-align: center; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

db.init_db()
budgets = db.load_budgets()

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "textbooks.xlsx")
@st.cache_data(show_spinner=False)
def load_catalog():
    if not os.path.exists(DB_PATH):
        return pd.DataFrame()
    df = pd.read_excel(DB_PATH, sheet_name=0)
    df.columns = df.columns.str.strip()
    if 'ราคา' in df.columns:
        df['ราคา_num'] = pd.to_numeric(df['ราคา'].astype(str).str.replace('บาท','').str.replace(',','').str.strip(), errors='coerce').fillna(0)
    else:
        df['ราคา_num'] = 0.0
    return df

catalog = load_catalog()

if 'cart' not in st.session_state:
    st.session_state.cart = {} 

st.title("🛒 ระบบเบิกจ่ายและสั่งหนังสือเรียนสำหรับคุณครู")

# --- โซนล็อกอินจำแลง ---
with st.container(border=True):
    st.markdown("#### 👤 ข้อมูลผู้สั่งซื้อ (โปรดกรอกก่อนเริ่มใช้งาน)")
    c1, c2 = st.columns(2)
    with c1:
        teacher_name = st.text_input("ชื่อ-นามสกุล คุณครูผู้ทำรายการ:")
        st.caption("💡 หากสอนหลายชั้น กรุณาทำรายการและกดยืนยันให้จบทีละชั้น เพื่อแยกงบประมาณออกจากกัน")
    with c2:
        class_name = st.selectbox("ระดับชั้นที่ต้องการเบิกงบประมาณ:", options=["--- กรุณาเลือก ---"] + list(budgets.keys()))

if not teacher_name or class_name == "--- กรุณาเลือก ---":
    st.warning("⚠️ กรุณากรอกชื่อและเลือกระดับชั้นให้ครบถ้วนก่อน เพื่อเริ่มใช้งานระบบและดูงบประมาณครับ")
    st.stop()

limit = budgets.get(class_name, 0.0)
submitted_total = db.get_submitted_total(class_name)

current_cart_total = 0.0
for k, v in st.session_state.cart.items():
    current_cart_total += v['qty'] * v['price']

remaining = limit - submitted_total - current_cart_total
is_over_budget = remaining < 0

# --- Dashboard ---
st.markdown(f"### 📊 สถานะงบประมาณของ **{class_name}**")
mc1, mc2, mc3, mc4 = st.columns(4)
mc1.metric("💰 งบจัดสรร", f"฿ {limit:,.2f}")
mc2.metric("📦 ยืนยันสั่งไปแล้ว", f"฿ {submitted_total:,.2f}")
mc3.metric("🛒 ยอดในตะกร้าขณะนี้", f"฿ {current_cart_total:,.2f}")

if is_over_budget:
    mc4.metric("🚨 เกินยอดงบไปแล้ว", f"฿ {abs(remaining):,.2f}", "- ทะลุเพดานงบ", delta_color="inverse")
    st.warning(f"⚠️ คำเตือน: ขณะนี้การสั่งซื้อของคุณเกินงบประมาณที่จัดสรรไปแล้ว **{abs(remaining):,.2f} บาท** (แต่คุณยังคงสามารถกดยืนยันการสั่งซื้อได้ ระบบจะบันทึกยอดไว้เผื่อพิจารณา)")
    progress_val = 1.0
else:
    mc4.metric("✅ งบประมาณคงเหลือ", f"฿ {remaining:,.2f}")
    total_spent = submitted_total + current_cart_total
    progress_val = total_spent / limit if limit > 0 else 0.0
    if progress_val > 1.0: progress_val = 1.0
    
st.progress(progress_val, text=f"สัดส่วนการใช้งบประมาณ: {progress_val*100:.1f}%")
st.markdown("---")

# --- Catalog / Cart ---
col_sidebar, col_main = st.columns([1, 3])

with col_sidebar:
    st.markdown("#### 🛒 ตะกร้าสินค้าของคุณ")
    if len(st.session_state.cart) == 0:
        st.info("ยังไม่มีสินค้าในตะกร้า")
    else:
        for k, v in list(st.session_state.cart.items()):
            with st.container(border=True):
                st.markdown(f"**{v['name']}**")
                st.caption(f"จำนวน: {v['qty']} เล่ม | รวม: {v['qty']*v['price']:,.2f} บ.")
                if st.button("❌ นำออก", key=f"del_{k}", use_container_width=True):
                    del st.session_state.cart[k]
                    st.rerun()
                
        if st.button("🗑️ ล้างตะกร้าทั้งหมด", use_container_width=True):
            st.session_state.cart = {}
            st.rerun()

    st.markdown("#### ✅ ยืนยันคำสั่งซื้อ")
    if current_cart_total == 0:
        st.button("📄 ยืนยันการสั่งซื้อ", disabled=True, use_container_width=True)
    else:
        btn_label = "📄 ยืนยันการสั่งซื้อ (ทะลุงบ)" if is_over_budget else "📄 ยืนยันการสั่งซื้อ"
        btn_type = "secondary" if is_over_budget else "primary"
        
        if st.button(btn_label, type=btn_type, use_container_width=True):
            items = [{'book_name': v['name'], 'publisher': v['pub'], 'price': v['price'], 'qty': v['qty']} for k,v in st.session_state.cart.items()]
            success, msg = db.save_order(teacher_name, class_name, items)
            if success:
                st.session_state.cart = {}
                st.success("🎉 ระบบได้รับใบคำสั่งซื้อเรียบร้อยแล้ว! ข้อมูลถูกส่งไปยังฝ่ายจัดซื้อทันที")
                st.balloons()
            else:
                st.error(f"❌ {msg}")
            
    st.markdown("---")
    st.markdown("#### 🔍 ตัวกรองค้นหาแคตตาล็อก")
    
    filter_my_class = st.checkbox(f"🎓 โชว์เฉพาะหนังสือชั้น {class_name}", value=True)
    search_q = st.text_input("ชื่อหนังสือ หรือ รหัส:")
    
    subjects = ["แสดงทั้งหมด"] + sorted(catalog['กลุ่มสาระการเรียนรู้'].dropna().astype(str).unique().tolist())
    search_subject = st.selectbox("หมวดหมู่ / กลุ่มสาระฯ:", options=subjects)
    
    pubs = ["แสดงทั้งหมด"] + sorted(catalog['ผู้จัดพิมพ์'].dropna().astype(str).unique().tolist())
    search_pub = st.selectbox("สำนักพิมพ์:", options=pubs)
    

with col_main:
    st.markdown("#### 📚 แคตตาล็อกหนังสือ")
    
    res = catalog.copy()
    if search_q: res = res[res['ชื่อหนังสือ'].astype(str).str.contains(search_q, case=False, na=False)]
    if search_subject != "แสดงทั้งหมด": res = res[res['กลุ่มสาระการเรียนรู้'].astype(str) == search_subject]
    if search_pub != "แสดงทั้งหมด": res = res[res['ผู้จัดพิมพ์'].astype(str) == search_pub]
    
    if filter_my_class and class_name != "--- กรุณาเลือก ---":
        def extract_year_number(text):
            """ดึงตัวเลขสุดท้ายจากชื่อชั้นเช่น 'ประถมศึกษาปีที่ 3' -> 3"""
            import re
            m = re.search(r'(\d+)', str(text))
            return int(m.group(1)) if m else None
        
        def match_class(c_user, c_db):
            c_u = str(c_user).replace(" ", "")
            c_d = str(c_db).replace(" ", "")
            if not c_d or c_d == 'nan' or c_d == '-': return False
            
            # \u0e15รงตรง
            if c_u in c_d or c_d in c_u: return True
            
            # จับคู่ชื่อย่อ \u0e1b./\u0e21./\u0e2d.
            abbr = c_u.replace("อนุบาลปีที่","อ.").replace("ประถมศึกษาปีที่","ป.").replace("มัธยมศึกษาปีที่","ม.")
            if abbr != c_u and (abbr in c_d or c_d in abbr): return True
            
            # จับคู่ช่วงชั้น \u0e40ช่น "มัธยมศึกษาปีที่ 4-6" กับ \u0e1bีที่ 4, 5, 6
            import re
            range_match = re.search(r'(\d+)-(\d+)', c_d)
            if range_match:
                start_y = int(range_match.group(1))
                end_y = int(range_match.group(2))
                user_y = extract_year_number(c_user)
                if user_y and start_y <= user_y <= end_y:
                    return True
            
            # จับคู่ช่วงที่ปนอยู่ในสตริงแบบ "ประถมศึกษาปีที่ 1 เล่ม2" (\u0e08ากเล่ม 1 \u0e41ละ \u0e40ล่ม 2)
            multi_match = re.search(r'(\d+)(\u0e41ละ|,|/|-)(\d+)', c_d)
            if multi_match:
                nums = re.findall(r'\d+', c_d)
                user_y = extract_year_number(c_user)
                if user_y and str(user_y) in nums:
                    return True
            
            return False
        
        res = res[res['ชั้น'].apply(lambda x: match_class(class_name, x))]
    
    if len(res) > 90:
        st.caption(f"พบหนังสือเข้าข่าย {len(res)} เล่ม (ระบบจำกัดการโชว์ภาพแค่ 90 เล่มแรก โปรดพิมพ์ตัวกรองค้นหาให้แคบลงเพื่อป้องกันเว็บอืด)")
        res = res.head(90)
    else:
        st.caption(f"พบหนังสือ {len(res)} เล่ม")
        
    for i in range(0, len(res), 4):
        cols = st.columns(4)
        for j in range(4):
            if i + j < len(res):
                row = res.iloc[i + j]
                book_id = str(row.name)
                
                with cols[j]:
                    with st.container(border=True):
                        # แสดงหน้าปกตัวอย่างแบบไม่ให้ภาพแตก
                        url_img = row.get('URL_รูปภาพ', '')
                        if pd.notnull(url_img) and str(url_img).startswith('http'):
                            img_data = fetch_image(str(url_img))
                            if img_data:
                                st.image(img_data, use_container_width=False, width=160)
                            else:
                                # Fallback: ลองให้ browser โหลดเอง
                                st.markdown(f'<div style="text-align:center;margin-bottom:10px;"><img src="{url_img}" style="max-height:160px;max-width:100%;object-fit:contain;" onerror="this.parentElement.innerHTML=\'<div style=background:#f1f5f9;height:160px;display:flex;align-items:center;justify-content:center;color:#94a3b8;border-radius:4px>\'\u0e44\u0e21\'\u0e21\'\u0e35\u0e20\u0e32\u0e1e</div>\'"></div>', unsafe_allow_html=True)
                        else:
                            st.markdown("""<div style="background:#f1f5f9; height:160px; display:flex; 
                                       align-items:center; justify-content:center; color:#94a3b8; 
                                       border-radius:4px; margin-bottom:10px;">ไม่มีภาพ</div>""", unsafe_allow_html=True)
                            
                        st.markdown(f"<div class='title-col'>{row.get('ชื่อหนังสือ', '-')}</div>", unsafe_allow_html=True)
                        st.caption(f"ระดับ: {row.get('ชั้น', '-')} | หมวด: {row.get('กลุ่มสาระการเรียนรู้', '-')}")
                        st.markdown(f"<small>🏢 พิมพ์โดย: {row.get('ผู้จัดพิมพ์', '-')}</small>", unsafe_allow_html=True)
                        
                        price = float(row.get('ราคา_num', 0))
                        st.markdown(f"<div class='price-col'>ราคาปก: ฿ {price:,.2f}</div>", unsafe_allow_html=True)
                        
                        cur_qty = st.session_state.cart.get(book_id, {}).get('qty', 0)
                        qty = st.number_input("จำนวนสั่งซื้อ:", min_value=0, max_value=2000, value=cur_qty, step=1, key=f"qty_{book_id}")
                        
                        if st.button("➕ ตกลง / ยกเลิก" if cur_qty > 0 else "➕ เพิ่มลงตะกร้า", key=f"add_{book_id}", use_container_width=True):
                            if qty > 0:
                                st.session_state.cart[book_id] = {
                                    'name': row.get('ชื่อหนังสือ', '-'),
                                    'pub': row.get('ผู้จัดพิมพ์', '-'),
                                    'price': price,
                                    'qty': qty
                                }
                                st.rerun()
                            elif qty == 0 and book_id in st.session_state.cart:
                                del st.session_state.cart[book_id]
                                st.rerun()
