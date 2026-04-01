import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import io
import os
import subprocess
import db

db.init_db()

# ==================== ตั้งค่า ====================
st.set_page_config(
    page_title="ระบบตรวจสอบและจัดทำรายการสื่อการเรียนรู้",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }
    
    .app-header {
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 70%, #60a5fa 100%);
        padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(37,99,235,0.15); color: white;
    }
    .app-header h1 { font-size: 1.8rem; font-weight: 700; margin: 0; color: white;}
    .app-header p  { font-size: 0.9rem; margin: 0.3rem 0 0 0; opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ==================== โหลดข้อมูล ====================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "textbooks.xlsx")

@st.cache_data(show_spinner=False)
def load_data(filepath):
    if not os.path.exists(filepath):
        return pd.DataFrame(), f"ไม่พบไฟล์ฐานข้อมูลหลักที่ {filepath}"
    try:
        df = pd.read_excel(filepath, sheet_name=0)
        df.columns = df.columns.str.strip()
        if 'ราคา' in df.columns:
            df['ราคา_num'] = pd.to_numeric(
                df['ราคา'].astype(str).str.replace('บาท','').str.replace(',','').str.strip(), errors='coerce'
            )
        df['_id'] = range(len(df))
        return df, None
    except Exception as e:
        return pd.DataFrame(), f"เกิดข้อผิดพลาดในการโหลดไฟล์: {e}"

df_db, error_msg = load_data(DB_PATH)

def map_class_name(raw_class):
    """แปลงตัวย่อระดับชั้นเป็นชื่อเต็ม"""
    if not isinstance(raw_class, str) or not raw_class.strip(): return ""
    c = raw_class.strip().lower()
    mapping = {
        'อ.1': 'อนุบาลปีที่ 1', 'อ.2': 'อนุบาลปีที่ 2', 'อ.3': 'อนุบาลปีที่ 3',
        'ป.1': 'ประถมศึกษาปีที่ 1', 'ป.2': 'ประถมศึกษาปีที่ 2', 'ป.3': 'ประถมศึกษาปีที่ 3',
        'ป.4': 'ประถมศึกษาปีที่ 4', 'ป.5': 'ประถมศึกษาปีที่ 5', 'ป.6': 'ประถมศึกษาปีที่ 6',
        'ม.1': 'มัธยมศึกษาปีที่ 1', 'ม.2': 'มัธยมศึกษาปีที่ 2', 'ม.3': 'มัธยมศึกษาปีที่ 3',
        'ม.4': 'มัธยมศึกษาปีที่ 4', 'ม.5': 'มัธยมศึกษาปีที่ 5', 'ม.6': 'มัธยมศึกษาปีที่ 6'
    }
    for k, v in mapping.items():
        if k in c:
            return v
    return raw_class

# ==================== Header ====================
st.markdown("""
<div class="app-header">
    <h1>📚 ระบบตรวจสอบและจัดทำรายการสื่อการเรียนรู้</h1>
    <p>ระบบนำเข้า จับคู่ และส่งออกข้อมูลหนังสือเรียนตามบัญชีสื่อการเรียนรู้ขั้นพื้นฐาน</p>
</div>
""", unsafe_allow_html=True)

if error_msg:
    st.error(f"❌ {error_msg}")
    st.stop()
# ==========================================
# เมนูด้านข้าง (Sidebar) สำหรับจัดการระบบ
# ==========================================
with st.sidebar:
    st.markdown("### 🎯 ตั้งค่าความแม่นยำ (Fuzzy Search)")
    st.session_state.price_match_threshold = st.slider("ความแม่นยำกรณีราคาตรงกัน", min_value=50, max_value=100, value=80, step=5)
    st.session_state.name_match_threshold = st.slider("ความแม่นยำกรณีค้นชื่อเพียวๆ", min_value=50, max_value=100, value=95, step=5)
    
    st.markdown("---")
    st.markdown("### ⚙️ การจัดการระบบ")
    st.info("หากมีการรันบอทอัปเดตข้อมูลหลังบ้านเสร็จแล้ว ให้กดปุ่มด้านล่างนี้เพื่อให้ระบบโหลดข้อมูลชุดใหม่")
    if st.button("🔄 โหลดฐานข้อมูลใหม่ (Refresh)", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.success("ล้างความจำและโหลดข้อมูลใหม่เรียบร้อย!")
        st.rerun()

# ==================== TABS ====================
tab1, tab2, tab3, tab4 = st.tabs(["📋 อัปโหลดตรวจสอบบัญชี", "🔍 ค้นหาในฐานข้อมูล", "⚙️ อัปเดตฐานข้อมูลแม่", "💰 แจกแจงงบและคำสั่งซื้อครู"])

# --------------------------------------------------
# TAB 1: จับคู่และส่งออก (Core Feature)
# --------------------------------------------------
with tab1:
    st.markdown("### 1. นำเข้าไฟล์รายชื่อหนังสือ (Excel หรือ CSV)")
    
    tmpl = pd.DataFrame({
        'ชื่อหนังสือ': ['ภาษาพาที', 'คณิตศาสตร์ เล่ม 1', 'วิทยาศาสตร์'], 
        'ชั้น': ['ประถมศึกษาปีที่ 1', 'มัธยมศึกษาปีที่ 4', ''],
        'สำนักพิมพ์': ['องค์การค้าของ สกสค.', 'สสวท.', ''],
        'ราคา': [65, 80, 120]
    })
    st.download_button("⬇️ ดาวน์โหลด Template ตัวอย่าง", data=tmpl.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'), file_name="template_รายชื่อหนังสือ.csv", mime="text/csv")
    
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ (รองรับ .xlsx, .xls, .csv)", type=['csv', 'xlsx', 'xls'])
    
    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        
        # --- อ่านไฟล์และเลือกชีต ---
        if file_ext in ['xlsx', 'xls']:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            if len(sheet_names) > 1:
                selected_sheets = st.multiselect("📑 ตรวจพบหลายชีต เลือกชีตที่ต้องการดึง (เลือกพร้อมกันได้)", options=sheet_names, default=[sheet_names[0]])
            else:
                selected_sheets = sheet_names
            
            if not selected_sheets:
                st.warning("⚠️ กรุณาเลือกอย่างน้อย 1 ชีต")
                st.stop()
                
            dfs = []
            for sheet in selected_sheets:
                df_sheet = pd.read_excel(xls, sheet_name=sheet)
                df_sheet['ชีตที่นำเข้า'] = sheet
                dfs.append(df_sheet)
            df_up_raw = pd.concat(dfs, ignore_index=True)
            
        else:
            try:
                df_up_raw = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            except:
                uploaded_file.seek(0)
                df_up_raw = pd.read_csv(uploaded_file, encoding='cp874')
            df_up_raw['ชีตที่นำเข้า'] = 'CSV'
            
        df_up_raw.columns = df_up_raw.columns.astype(str).str.strip()
        all_cols = list(df_up_raw.columns)
        
        st.markdown("---")
        st.markdown("### 2. ตั้งค่าจับคู่คอลัมน์ (Column Mapping)")
        st.info("📌 กรุณาจิ้มเลือกคอลัมน์จากไฟล์ของคุณ ให้ตรงกับข้อมูลที่ระบบต้องการ (ถ้าหาไม่เจอให้เลือก 'ปล่อยว่าง')")
        
        def find_default_col(keywords, cols):
            for kw in keywords:
                for col in cols:
                    if kw in col.lower():
                        return cols.index(col)
            return len(cols) # index ของ "--- ปล่อยว่าง ---"
            
        options = all_cols + ["--- ปล่อยว่าง ---"]
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            idx_name = find_default_col(['ชื่อ', 'รายการ', 'หนังสือ'], all_cols)
            if idx_name == len(all_cols):
                options_name = all_cols
                idx_name = 0
            else:
                options_name = all_cols
            map_name = st.selectbox("📖 คอลัมน์ 'ชื่อหนังสือ'*", options=options_name, index=idx_name)
            
        with col2:
            idx_class = find_default_col(['ชั้น', 'ระดับ', 'ห้อง'], all_cols)
            map_class = st.selectbox("🎓 คอลัมน์ 'ระดับชั้น'", options=options, index=idx_class)
            
        with col3:
            idx_price = find_default_col(['ราคา', 'บาท'], all_cols)
            map_price = st.selectbox("💰 คอลัมน์ 'ราคา'", options=options, index=idx_price)
            
        with col4:
            idx_pub = find_default_col(['สำนักพิมพ์', 'ผู้จัดพิมพ์', 'สนพ'], all_cols)
            map_pub = st.selectbox("🏢 คอลัมน์ 'สำนักพิมพ์'", options=options, index=idx_pub)
            
        if st.button("🔍 เริ่มจับคู่ข้อมูล", type="primary"):
            # แมพข้อมูลลง dataframe โครงสร้างใหม่ที่ระบบเข้าใจ
            df_up = pd.DataFrame()
            df_up['ชื่อหนังสือ'] = df_up_raw[map_name]
            df_up['ชั้น'] = df_up_raw[map_class] if map_class in all_cols else ""
            df_up['ราคา'] = df_up_raw[map_price] if map_price in all_cols else ""
            df_up['สำนักพิมพ์'] = df_up_raw[map_pub] if map_pub in all_cols else ""
            df_up['ชีตที่นำเข้า'] = df_up_raw['ชีตที่นำเข้า']
            
            # กรองแถวที่ชื่อหนังสือเป็นว่างทิ้ง
            df_up = df_up.dropna(subset=['ชื่อหนังสือ'])
            progress = st.progress(0, text="กำลังประมวลผลจับคู่...")
            results = []
            
            total_rows = len(df_up)
            update_step = max(1, total_rows // 20) # Update every 5%
            
            for idx, row in df_up.iterrows():
                if idx % update_step == 0 or idx == total_rows - 1:
                    progress.progress((idx + 1) / total_rows, text=f"กำลังตรวจสอบรายการที่ {idx+1}/{total_rows}")
                
                q_name = str(row['ชื่อหนังสือ']).strip()
                q_class_raw = str(row['ชั้น']).strip() if 'ชั้น' in df_up.columns and pd.notnull(row['ชั้น']) else ""
                q_class = map_class_name(q_class_raw)
                q_price_raw = row.get('ราคา', '')
                
                q_pub = ""
                if 'สำนักพิมพ์' in df_up.columns and pd.notnull(row['สำนักพิมพ์']):
                    q_pub = str(row['สำนักพิมพ์']).strip()
                elif 'ผู้จัดพิมพ์' in df_up.columns and pd.notnull(row['ผู้จัดพิมพ์']):
                    q_pub = str(row['ผู้จัดพิมพ์']).strip()
                
                try:
                    q_price = float(str(q_price_raw).replace(',','').strip())
                except:
                    q_price = None
                
                match_found = False
                matched_db_idx = None
                options = []
                
                df_search_pool = df_db.copy()
                
                # Step 1: ค้นหาจาก ราคา + ชื่อ
                if pd.notnull(q_price):
                    price_match = df_search_pool[df_search_pool['ราคา_num'] == q_price]
                    if not price_match.empty:
                        scores = price_match['ชื่อหนังสือ'].astype(str).apply(lambda x: fuzz.token_sort_ratio(q_name, x))
                        best_idx = scores.idxmax()
                        if scores[best_idx] >= st.session_state.get('price_match_threshold', 80):
                            db_class = str(price_match.loc[best_idx].get('ชั้น', ''))
                            if not q_class or q_class in db_class or db_class in q_class:
                                match_found = True
                                matched_db_idx = best_idx
                
                # Step 2: หาด้วยชื่อ และจำกัดให้อยู่ในระดับชั้นเดียวกัน
                if not match_found:
                    search_df = df_search_pool.copy()
                    
                    if q_class:
                        class_filtered = search_df[search_df['ชั้น'].astype(str).str.contains(q_class, case=False, na=False)]
                        if not class_filtered.empty:
                            search_df = class_filtered
                    
                    scores = search_df['ชื่อหนังสือ'].astype(str).apply(lambda x: fuzz.token_sort_ratio(q_name, x))
                    
                    n_options = min(10, len(search_df))
                    top_match_idx = scores.nlargest(n_options).index
                    
                    if len(top_match_idx) > 0 and scores[top_match_idx[0]] >= st.session_state.get('name_match_threshold', 95):
                        db_class = str(search_df.loc[top_match_idx[0]].get('ชั้น', ''))
                        if not q_class or q_class in db_class or db_class in q_class:
                            match_found = True
                            matched_db_idx = top_match_idx[0]
                    
                    if not match_found:
                        for i in top_match_idx:
                            db_r = df_db.loc[i]
                            opt_label = f"✨ [{db_r.get('ชั้น','-')}] {db_r['ชื่อหนังสือ']} | ราคา: {db_r.get('ราคา','-')} | สนพ: {db_r.get('ผู้จัดพิมพ์','-')}"
                            options.append({'idx': i, 'label': opt_label})
                
                results.append({
                    'id': idx,
                    'q_name': q_name,
                    'q_class': q_class,
                    'q_pub': q_pub,
                    'q_price': q_price_raw,
                    'q_sheet': row.get('ชีตที่นำเข้า', 'CSV'),
                    'status': '🟢 ตรง' if match_found else '🔴 รอแก้ไข',
                    'db_idx': matched_db_idx,
                    'options': options
                })
            
            progress.empty()
            st.session_state.match_results = results
            
    # --- แสดงผลตารางหลัก ---
    if 'match_results' in st.session_state:
        st.markdown("---")
        results = st.session_state.match_results
        
        view_data = []
        unmatched_items = []
        for r in results:
            if r['status'] in ['🟢 ตรง', '🟡 แก้ไขแล้ว']:
                db_r = df_db.loc[r['db_idx']]
                view_data.append({
                    'สถานะ': r['status'],
                    'ชีตที่นำเข้า': r.get('q_sheet', '-'),
                    'ชื่อที่อัปโหลด': r['q_name'],
                    'ชั้นที่อัปโหลด': r['q_class'],
                    'สนพ.ที่อัปโหลด': r['q_pub'],
                    'ราคาที่อัปโหลด': r['q_price'],
                    'ชื่อในระบบ': db_r['ชื่อหนังสือ'],
                    'ชั้นในระบบ': db_r.get('ชั้น', '-'),
                    'สนพ.ในระบบ': db_r.get('ผู้จัดพิมพ์', '-'),
                    'ราคาในระบบ': db_r.get('ราคา', '-')
                })
            elif r['status'] == '⚪ ข้าม':
                view_data.append({
                    'สถานะ': '⚪ ข้าม',
                    'ชีตที่นำเข้า': r.get('q_sheet', '-'),
                    'ชื่อที่อัปโหลด': r['q_name'],
                    'ชั้นที่อัปโหลด': r['q_class'],
                    'สนพ.ที่อัปโหลด': r['q_pub'],
                    'ราคาที่อัปโหลด': r['q_price'],
                    'ชื่อในระบบ': '--- ใช้ข้อมูลเดิม ---',
                    'ชั้นในระบบ': '-',
                    'สนพ.ในระบบ': '-',
                    'ราคาในระบบ': '-'
                })
            else: # 🔴 รอแก้ไข
                unmatched_items.append(r)
                view_data.append({
                    'สถานะ': '🔴 รอแก้ไข',
                    'ชีตที่นำเข้า': r.get('q_sheet', '-'),
                    'ชื่อที่อัปโหลด': r['q_name'],
                    'ชั้นที่อัปโหลด': r['q_class'],
                    'สนพ.ที่อัปโหลด': r['q_pub'],
                    'ราคาที่อัปโหลด': r['q_price'],
                    'ชื่อในระบบ': '--- ยังไม่พบ ---',
                    'ชั้นในระบบ': '-',
                    'สนพ.ในระบบ': '-',
                    'ราคาในระบบ': '-'
                })
                
        # --- Dashboard ---
        st.markdown("### 2. สรุปผลการตรวจสอบข้อมูล")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📚 จำนวนที่นำเข้า", f"{len(results)} รายการ")
        col2.metric("🟢 จับคู่ตรงเป้า", f"{len([r for r in results if r['status'] == '🟢 ตรง'])} รายการ")
        col3.metric("🔴 รอผู้ใช้แก้ไข", f"{len(unmatched_items)} รายการ")
        col4.metric("🟡 ข้าม/แก้ไขแล้ว", f"{len([r for r in results if r['status'] in ['🟡 แก้ไขแล้ว', '⚪ ข้าม']])} รายการ")
        
        df_view = pd.DataFrame(view_data)
        st.dataframe(df_view, use_container_width=True, hide_index=True)
        
        # --- ส่วนแก้ไขรายการที่ไม่ตรง ---
        if unmatched_items:
            st.markdown(f"### 3. ตรวจสอบและแก้ไขรายการที่ไม่ตรง (เหลือ {len(unmatched_items)} รายการ)")
            st.info("💡 เลือกข้อที่ถูกต้องให้ครบทุกรายการ แล้วกดปุ่ม \"บันทึกการแก้ไขทั้งหมด\" ที่ด้านล่างสุดเพียงครั้งเดียวครับ")
            
            with st.form("resolve_form"):
                for item in unmatched_items:
                    with st.container(border=True):
                        st.markdown(f"##### 📙 {item['q_name']}")
                        
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"🎓 **ชั้น:** `{item['q_class'] if item['q_class'] else '-'}`")
                        c2.markdown(f"🏢 **สำนักพิมพ์:** `{item['q_pub'] if item['q_pub'] else '-'}`")
                        c3.markdown(f"💰 **ราคา:** `{item['q_price'] if item['q_price'] else '-'}`")
                        
                        # ตัวเลือกเริ่มต้น
                        dd_options = [
                            {'idx': -2, 'label': '--- ข้ามไปก่อน ยังไม่ตั้งค่า ---'},
                            {'idx': -1, 'label': '❌ ข้ามรายการนี้ (ใช้ข้อมูลเดิมที่อัปโหลดมา)'}
                        ] + item['options']
                        
                        st.selectbox(
                            "📌 เลือกหนังสือจากฐานข้อมูล หรือเลือกข้าม:",
                            options=dd_options,
                            format_func=lambda x: x['label'],
                            key=f"sel_{item['id']}"
                        )
                
                submitted = st.form_submit_button("💾 บันทึกการแก้ไขทั้งหมด", type="primary", use_container_width=True)
                if submitted:
                    updated_count = 0
                    for item in unmatched_items:
                        sel = st.session_state[f"sel_{item['id']}"]
                        if sel['idx'] != -2:
                            updated_count += 1
                            # อัปเดตสถานะในตัวแปรหลัก
                            for r in st.session_state.match_results:
                                if r['id'] == item['id']:
                                    if sel['idx'] == -1:
                                        r['status'] = '⚪ ข้าม'
                                    else:
                                        r['status'] = '🟡 แก้ไขแล้ว'
                                        r['db_idx'] = sel['idx']
                                    break
                    if updated_count > 0:
                        st.rerun() # สั่งรีเฟรชหน้าเว็บ เพื่อตัดรายการออก
                    else:
                        st.warning("ไม่มีการเปลี่ยนแปลงข้อมูลครับ")
        else:
            st.success("✨ ยอดเยี่ยม! จัดการรายการข้อมูลครบทุกรายการแล้ว")
        
            
        # --- ส่งออกข้อมูล ---
        st.markdown("---")
        st.markdown("### 4. ส่งออกข้อมูล (Export)")
        
        export_rows = []
        for r in results:
            if r['status'] in ['🟢 ตรง', '🟡 แก้ไขแล้ว']:
                db_row = df_db.loc[r['db_idx']].copy()
                db_row['ชื่อที่นำเข้า'] = r['q_name']
                db_row['ชีตที่นำเข้า'] = r.get('q_sheet', '-')
                export_rows.append(db_row)
            elif r['status'] == '⚪ ข้าม':
                # กรณีที่กดข้าม ให้ใช้ข้อมูลเดิมที่อัปโหลดมาส่งออกไปเลย
                skip_row = {
                    'ชีตที่นำเข้า': r.get('q_sheet', '-'),
                    'ชื่อที่นำเข้า': r['q_name'],
                    'ชื่อหนังสือ': r['q_name'],
                    'ชั้น': r['q_class'],
                    'ผู้จัดพิมพ์': r['q_pub'],
                    'ราคา': r['q_price']
                }
                export_rows.append(skip_row)
                    
        if export_rows:
            df_export = pd.DataFrame(export_rows)
            
            # จัดเรียงคอลัมน์ เอาข้อมูลอ้างอิงนำเข้าขึ้นก่อน 
            # ตามด้วยข้อมูลทั้งหมดจากฐานข้อมูล (ดึงมาครบทุกคอลัมน์)
            db_cols = [c for c in df_db.columns if c not in ['_id', 'ราคา_num', 'ลำดับ']]
            final_cols = ['ชีตที่นำเข้า', 'ชื่อที่นำเข้า'] + db_cols
            
            # ประกอบร่างข้อมูล
            final_cols = [c for c in final_cols if c in df_export.columns]
            df_export = df_export[final_cols].fillna("-")
            
            st.info(f"พร้อมส่งออกข้อมูลจำนวน **{len(df_export)}** รายการ (รวมรายการที่กดข้าม โดยใช้ข้อมูลเดิมด้วย)")
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='รายการสั่งซื้อ')
                
            st.download_button(
                "📥 ดาวน์โหลดไฟล์ Excel (พร้อมใช้งาน)",
                data=out.getvalue(),
                file_name="รายการบัญชีหนังสือ_อัปเดตแล้ว.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("ไม่มีรายการข้อมูลที่สมบูรณ์สำหรับส่งออก")

# --------------------------------------------------
# TAB 2: ค้นหาทั่วไป
# --------------------------------------------------
with tab2:
    st.markdown("### ค้นหาหนังสือในฐานข้อมูล")
    c1, c2, c3, c4 = st.columns(4)
    with c1: search_q = st.text_input("ชื่อหนังสือ:")
    with c2: search_c = st.text_input("ระดับชั้น:")
    with c3: search_p = st.text_input("ราคา:")
    with c4: search_pub = st.text_input("สำนักพิมพ์:")
    
    if search_q or search_c or search_p or search_pub:
        res = df_db.copy()
        if search_q: res = res[res['ชื่อหนังสือ'].astype(str).str.contains(search_q, case=False, na=False)]
        if search_c: res = res[res['ชั้น'].astype(str).str.contains(search_c, case=False, na=False)]
        if search_p: res = res[res['ราคา'].astype(str).str.contains(search_p, case=False, na=False)]
        if search_pub: res = res[res['ผู้จัดพิมพ์'].astype(str).str.contains(search_pub, case=False, na=False)]
        
        st.dataframe(res[[c for c in ['ชื่อหนังสือ','ชั้น','ผู้จัดพิมพ์','ราคา'] if c in res.columns]], use_container_width=True, hide_index=True)

# --------------------------------------------------
# TAB 3: อัปเดตฐานข้อมูล (Master Data)
# --------------------------------------------------
with tab3:
    st.markdown("### ⚙️ อัปเดตฐานข้อมูลหนังสือเรียนหลัก")
    st.warning(f"ไฟล์ฐานข้อมูลปัจจุบันมีจำนวน: **{len(df_db):,}** รายการ")
    
    st.markdown("---")
    st.markdown("#### 🚀 ดึงข้อมูลล่าสุดจากเว็บไซต์อัตโนมัติ")
    st.info("💡 เมื่อกดปุ่ม ระบบจะเปิดหน้าต่างการทำงานสีดำ (Terminal) ขึ้นมาใหม่ เพื่อไม่ให้หน้าเว็บค้าง ระหว่างนี้คุณสามารถใช้งานระบบค้นหาต่อได้เลยครับ")
    
    if st.button("▶️ เปิดโปรแกรมอัปเดตฐานข้อมูล (รันไฟล์ .bat)", type="primary"):
        try:
            # ใช้ subprocess สั่งเปิดไฟล์ .bat ขึ้นมาในหน้าต่างใหม่ (สำหรับ Windows)
            subprocess.Popen(['cmd.exe', '/c', 'start', 'อัปเดตฐานข้อมูล.bat'])
            
            st.success("✅ เปิดหน้าต่างดึงข้อมูลสำเร็จ! กรุณาดูความคืบหน้าที่หน้าต่างสีดำครับ")
            st.warning("⚠️ สำคัญ: เมื่อหน้าต่างสีดำทำงานเสร็จและปิดไปแล้ว ให้คุณกดปุ่ม '🔄 โหลดฐานข้อมูลใหม่ (Refresh)' ที่เมนูด้านซ้ายมือ เพื่อให้ระบบดึงหนังสือชุดใหม่มาแสดงครับ")
            
        except Exception as e:
            st.error(f"❌ ไม่สามารถเปิดไฟล์ได้ กรุณาตรวจสอบว่ามีไฟล์ 'อัปเดตฐานข้อมูล.bat' อยู่ในโฟลเดอร์เดียวกับระบบหรือไม่ (Error: {e})")

    st.markdown("---")
    st.markdown("#### 📂 กรณีต้องการอัปโหลดไฟล์ด้วยตัวเอง (Manual)")
    new_db_file = st.file_uploader("เลือกไฟล์ Excel บัญชีสื่อฯ (นามสกุล .xlsx)", type=['xlsx'])
    
    if new_db_file:
        if st.button("💾 บันทึกและแทนที่ฐานข้อมูลเดิม", type="secondary"):
            try:
                with open(DB_PATH, "wb") as f:
                    f.write(new_db_file.getbuffer())
                st.cache_data.clear()
                st.success("✅ อัปเดตฐานข้อมูลสำเร็จ! ระบบกำลังรีเฟรช...")
                st.rerun()
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการบันทึกไฟล์: {e}")

# --------------------------------------------------
# TAB 4: ผู้ดูแลระบบสั่งหนังสือคุณครู
# --------------------------------------------------
with tab4:
    st.markdown("### 🛒 ศูนย์ควบคุมคำสั่งซื้อคุณครูรายชั้น (Dashboard)")
    
    st.markdown("#### 💰 1. จัดสรรงบประมาณต่อห้องเรียน")
    st.info("💡 นำไปใช้ในหน้าเว็บร้านค้าของคุณครู เพื่อควบคุมยอดและบีบให้สั่งซื้อไม่เกินยอดเงินอัตโนมัติ")
    
    df_budgets = db.load_budgets_df()
    edited_budgets = st.data_editor(df_budgets, num_rows="dynamic", use_container_width=True, hide_index=True)
    if st.button("💾 บันทึกการเปลี่ยนแปลงงบประมาณ", type="primary"):
        db.save_budgets(edited_budgets)
        st.success("บันทึกค่างบประมาณใหม่เรียบร้อยแล้ว!")
        
    st.markdown("---")
    st.markdown("#### 📦 2. คำสั่งซื้อที่ส่งเข้ามาจากคุณครูทั้งหมด")
    
    df_orders = db.load_orders()
    if len(df_orders) == 0:
        st.caption("ยังไม่มีข้อมูลคำสั่งซื้อจากคุณครูครับ แนะนำให้ส่งไฟล์เปิดเว็บให้คุณครูใช้งาน")
    else:
        st.markdown("---")
        # --- UI Filters ---
        c_filter, _ = st.columns([1, 2])
        all_classes = ["--- ดูรวมทั้งโรงเรียน ---"] + sorted(df_orders['class_name'].unique().tolist())
        selected_class = c_filter.selectbox("🎯 กรองดูเฉพาะชั้นเรียน (ดูผลบนเว็บ):", options=all_classes)
        
        display_df = df_orders.copy()
        if selected_class != "--- ดูรวมทั้งโรงเรียน ---":
            display_df = display_df[display_df['class_name'] == selected_class]
            
        view_type = st.radio("รูปแบบการแสดงผลหน้าจอเว็บ:", ["รายการแยกตามครูผู้สั่ง (ดูรายคน)", "สรุปยอดรวมหนังสือเพื่อเตรียมจัดซื้อ (อ้างอิงรายการ)"], horizontal=True)
        
        if view_type == "รายการแยกตามครูผู้สั่ง (ดูรายคน)":
            st.info("💡 คุณแอดมินสามารถคลิกแก้จำนวนเล่ม แก้ชั้น ลบแถวทิ้ง (กดตรงช่องซ้ายสุดของแถวแล้วกดปุ่ม Delete บนคีย์บอร์ด) หรือกดเครื่องหมาย + ด้านล่างตารางเพื่อเพิ่มหนังสือเองได้เลยครับ")
            
            edited_df = st.data_editor(
                display_df, 
                use_container_width=True, 
                hide_index=True, 
                num_rows="dynamic",
                column_config={
                    "id": None, # ซ่อน ID ไม่ให้แก้
                    "timestamp": None # ซ่อน Timestamp
                },
                key="orders_editor"
            )
            
            if st.button("💾 ยืนยัน บันทึกการเพิ่ม/ลบ/แก้ไข รายการลงฐานข้อมูล", type="primary"):
                success, msg = db.sync_orders(edited_df, display_df)
                if success:
                    st.success("บันทึกการปรับปรุงข้อมูลคำสั่งซื้อเรียบร้อยแล้ว!")
                    st.rerun()
                else:
                    st.error(f"เกิดข้อผิดพลาด: {msg}")
        else:
            agg_df = display_df.groupby(['class_name', 'book_name', 'publisher', 'unit_price']).agg({'quantity': 'sum', 'total_price': 'sum'}).reset_index()
            st.dataframe(agg_df, use_container_width=True, hide_index=True)
            
        st.markdown("---")
        st.markdown("#### 📥 ระบบดาวน์โหลดสถิติอัจฉริยะ (Multi-Sheet Excel)")
        st.info("💡 ไฟล์ที่ดาวน์โหลดจะประกอบด้วย: 1. ชีตสรุปยอดสิริรวมทั้งโรงเรียน และ 2. ชีตแยกย่อยของแต่ละชั้นเรียน พร้อมดึงข้อมูลวิชา ปีที่เผยแพร่ ขนาด ฯลฯ กลับมาให้ครบเต็มรูปแบบ!")
            
        c1, c2 = st.columns([1, 1])
        with c1:
            # 1. การฟิวชันข้อมูล (Left Join) กับ df_db
            df_db_copy = df_db.copy()
            cols_to_add = [c for c in df_db_copy.columns if c not in ['_id', 'ราคา_num', 'ลำดับ', 'ราคา', 'ชื่อหนังสือ', 'ผู้จัดพิมพ์']] 
            
            db_subset = df_db_copy[['ชื่อหนังสือ', 'ผู้จัดพิมพ์'] + cols_to_add].drop_duplicates(subset=['ชื่อหนังสือ', 'ผู้จัดพิมพ์'])
            
            master_orders = pd.merge(df_orders, db_subset, left_on=['book_name', 'publisher'], right_on=['ชื่อหนังสือ', 'ผู้จัดพิมพ์'], how='left')
            
            # 2. จัดเรียงและแปลงชื่อคอลัมน์
            new_db_cols = [c for c in cols_to_add if c in master_orders.columns]
            rename_map = {
                'book_name': 'ชื่อหนังสือ (ที่สั่งรอบนี้)',
                'publisher': 'สำนักพิมพ์/ผู้จัดพิมพ์ (ที่สั่ง)',
                'unit_price': 'ราคาต่อหน่วย',
                'quantity': 'จำนวนเล่ม',
                'total_price': 'ราคารวมทั้งหมด',
                'class_name': 'ระดับชั้น',
                'teacher_name': 'ผู้ทำรายการคำสั่งซื้อ',
                'timestamp': 'เวลาที่ยืนยันคำสั่ง'
            }
            master_orders.rename(columns=rename_map, inplace=True)
            
            # สร้างตัวแปร agg_dict สำหรับกลุ่มผลรวม
            agg_dict = {'จำนวนเล่ม': 'sum', 'ราคารวมทั้งหมด': 'sum'}
            for c in new_db_cols:
                agg_dict[c] = 'first' # ดึงค่าลักษณะจำเพาะของหนังสือมา 1 แถวเพื่ออ้างอิง
            
            # เตรียมโครงสร้างข้อมูลสรุปรวม
            summary_df = master_orders.groupby(['ชื่อหนังสือ (ที่สั่งรอบนี้)', 'สำนักพิมพ์/ผู้จัดพิมพ์ (ที่สั่ง)', 'ราคาต่อหน่วย'], dropna=False).agg(agg_dict).reset_index()
            
            # -------------------------------------------------------------
            # UI สำหรับเลือกสร้างประโยคคุณลักษณะหนังสือ
            st.markdown("##### ⚙️ แจกแจงคุณลักษณะหนังสือ (สำหรับไฟล์ฝ่ายพัสดุ)")
            st.caption("เลือกข้อมูลที่ต้องการนำมาต่อท้ายกันเป็นประโยค 'คุณลักษณะ' (เลือกคลิกเรียงตามลำดับก่อนหลังได้เลย)")
            
            default_specs = [c for c in new_db_cols if c in ['ประเภทสื่อ', 'กลุ่มสาระการเรียนรู้']]
            if not default_specs: default_specs = new_db_cols[:1] if new_db_cols else []
            
            selected_specs = st.multiselect("เลือกคุณลักษณะที่ต้องการ:", options=new_db_cols, default=default_specs)
            include_header = st.checkbox("ใส่ชื่อหัวข้อกำกับประโยคด้วย (เช่น ประเภทสื่อ: หนังสือเรียน | จำนวนหน้า: 120)", value=True)
            # -------------------------------------------------------------
            
            # ฟังก์ชันแปลง Dataframe สำหรับฝ่ายการเงิน (เรียงคอลัมน์และบวกผลรวม)
            def create_finance_df(source_df):
                f_df = pd.DataFrame()
                f_df['ลำดับ'] = range(1, len(source_df) + 1)
                f_df['ประเภทสื่อ'] = source_df['ประเภทสื่อ'] if 'ประเภทสื่อ' in source_df.columns else '-'
                f_df['ชื่อหนังสือ'] = source_df['ชื่อหนังสือ (ที่สั่งรอบนี้)']
                f_df['สำนักพิมพ์'] = source_df['สำนักพิมพ์/ผู้จัดพิมพ์ (ที่สั่ง)']
                f_df['จำนวนเล่ม'] = source_df['จำนวนเล่ม']
                f_df['ราคาต่อหน่วย'] = source_df['ราคาต่อหน่วย']
                f_df['ราคารวมทั้งหมด'] = source_df['ราคารวมทั้งหมด']
                
                if len(source_df) > 0:
                    total_row = pd.DataFrame([{
                        'ลำดับ': '',
                        'ประเภทสื่อ': '',
                        'ชื่อหนังสือ': '=== รวมยอดทั้งสิ้น ===',
                        'สำนักพิมพ์': '',
                        'จำนวนเล่ม': source_df['จำนวนเล่ม'].sum(),
                        'ราคาต่อหน่วย': '',
                        'ราคารวมทั้งหมด': source_df['ราคารวมทั้งหมด'].sum()
                    }])
                    # กำจัด warning behavior เวลา concat
                    f_df = pd.concat([f_df, total_row], ignore_index=True)
                    
                return f_df
            
            # ฟังก์ชันแปลง Dataframe ให้เหลือ 2 คอลัมน์ (ชื่อ กับ คุณลักษณะต่อกัน)
            def create_specs_df(source_df):
                specs_df = pd.DataFrame()
                
                def make_book_name(row):
                    media = str(row.get('ประเภทสื่อ', '')).strip()
                    name = str(row.get('ชื่อหนังสือ (ที่สั่งรอบนี้)', '')).strip()
                    if media and media not in ['nan', '-', 'None']:
                        # เช็กเล็กน้อยว่าถ้าในชื่อหนังสือมีคำว่า "หนังสือ..." นำหน้าอยู่แล้วจะได้ไม่ซ้ำซ้อน
                        if not name.startswith(media):
                            return f"{media}{name}"
                    return name
                
                specs_df['ชื่อหนังสือ'] = source_df.apply(make_book_name, axis=1)
                
                def make_spec_text(row):
                    parts = []
                    for c in selected_specs:
                        val = str(row.get(c, '')).strip()
                        if val and val != 'nan' and val != '-':
                            if include_header:
                                parts.append(f"{c}: {val}")
                            else:
                                parts.append(val)
                    return " , ".join(parts) if parts else "-"
                
                specs_df['คุณลักษณะ'] = source_df.apply(make_spec_text, axis=1)
                return specs_df
            
            out_finance = io.BytesIO()
            out_specs = io.BytesIO()
            
            # ส่งข้อมูลลงดิสก์จำลองพร้อมกัน 2 ไฟล์ด้วยชุดคอลัมน์ที่เล็มแล้ว
            with pd.ExcelWriter(out_finance, engine='openpyxl') as writer_f, \
                 pd.ExcelWriter(out_specs, engine='openpyxl') as writer_s:
                 
                # 3.1 บิลรวมตัวแม่ทั้งโรงเรียน (Summary Sheet)
                create_finance_df(summary_df).to_excel(writer_f, index=False, sheet_name='🔥 สรุปเป้ายอดรวมการเงิน')
                create_specs_df(summary_df).to_excel(writer_s, index=False, sheet_name='🔥 สรุปเป้าคุณลักษณะรวม')
                
                # 3.2 บิลแยกลูกย่อยทีละชั้นเรียน (Multi-Class Sheets)
                for c_name in sorted(master_orders['ระดับชั้น'].dropna().unique()):
                    class_df = master_orders[master_orders['ระดับชั้น'] == c_name]
                    class_summary = class_df.groupby(['ชื่อหนังสือ (ที่สั่งรอบนี้)', 'สำนักพิมพ์/ผู้จัดพิมพ์ (ที่สั่ง)', 'ราคาต่อหน่วย'], dropna=False).agg(agg_dict).reset_index()
                    
                    # เลี่ยงอักขระพิเศษในชื่อ Sheet และตัดความยาวห้ามเกิน 31
                    safe_sn = str(c_name)[:30]
                    for char in ['/', '\\', '?', '*', ':', '[', ']']: safe_sn = safe_sn.replace(char, '')
                    if not safe_sn.strip(): safe_sn = "Unknown_Class"
                    
                    create_finance_df(class_summary).to_excel(writer_f, index=False, sheet_name=safe_sn)
                    create_specs_df(class_summary).to_excel(writer_s, index=False, sheet_name=safe_sn)
                    
            st.markdown("##### 💸 สำหรับฝ่ายการเงิน (บัญชีเบิกจ่าย)")
            st.download_button("📥 บัญชีสั่งซื้อ-การเงิน", data=out_finance.getvalue(), file_name="บัญชีสั่งซื้อจัดจ้าง_ฝ่ายบัญชีการเงิน.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            
            st.markdown("##### 📝 สำหรับฝ่ายพัสดุ (อ้างอิงคุณลักษณะ)")
            st.download_button("📥 ประกาศคุณลักษณะสเปคพัสดุ", data=out_specs.getvalue(), file_name="ใบกำหนดคุณลักษณะเสริม_ฝ่ายพัสดุ.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="secondary", use_container_width=True)
            
        with c2:
            st.markdown("<br><br>", unsafe_allow_html=True) # ดันปุ่มให้ตรงกัน
            if st.button("🗑️ ล้างคำสั่งซื้อที่ค้างทิ้งทั้งหมด (คลิกเพื่อเริ่มเทอมใหม่)", use_container_width=True):
                db.clear_orders()
                st.success("ข้อมูลคำสั่งซื้อถูกทำความสะอาดหมดจดเรียบร้อยแล้ว!")
                st.rerun()