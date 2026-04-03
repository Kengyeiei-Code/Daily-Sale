import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta

# --- 1. ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Sam Roasters - Back Office", page_icon="☕", layout="centered")
st.title("☕ ระบบจัดการยอดขาย Sam Roasters")
st.markdown("อัปโหลดไฟล์จาก POS เพื่อบันทึกลงระบบฐานข้อมูลและดูรายงาน")

# --- 2. สร้างระบบฐานข้อมูล (Database) ซ่อนไว้หลังบ้าน ---
# ข้อมูลจะถูกเก็บในไฟล์ sam_roasters.db อย่างปลอดภัย
conn = sqlite3.connect('sam_roasters.db')
c = conn.cursor()
c.execute('CREATE TABLE IF NOT EXISTS sales_history (date TEXT, category TEXT, amount REAL)')
c.execute('CREATE TABLE IF NOT EXISTS beverage_history (date TEXT, menu TEXT, qty REAL)')
conn.commit()

# --- 3. หน้าจออัปโหลดไฟล์ ---
with st.form("upload_form"):
    st.subheader("📥 อัปโหลดไฟล์ประจำวัน")
    sale_file = st.file_uploader("1. ไฟล์ยอดรวม และ ส่วนลด (SaleReport)", type=['csv', 'xlsx'])
    bev_file = st.file_uploader("2. ไฟล์จำนวนแก้ว (SaleByBehavior)", type=['csv', 'xlsx'])
    
    # วันที่ที่ต้องการบันทึกยอด (ค่าเริ่มต้นคือเมื่อวาน)
    record_date = st.date_input("เลือกวันที่ของยอดขาย", datetime.today() - timedelta(days=1))
    
    submit_button = st.form_submit_button("ประมวลผลและบันทึกลงฐานข้อมูล")

# --- 4. ระบบประมวลผลเมื่อกดปุ่ม ---
if submit_button:
    if sale_file and bev_file:
        with st.spinner("กำลังประมวลผลข้อมูล..."):
            try:
                # อ่านไฟล์ด้วย Pandas (ฉลาดและเร็วกว่า Sheet)
                if bev_file.name.endswith('.csv'):
                    df_bev = pd.read_csv(bev_file)
                else:
                    df_bev = pd.read_excel(bev_file)
                
                # --- ตัวอย่างการสกัดข้อมูลจำนวนแก้ว ---
                # ระบบจะดึงเฉพาะคอลัมน์ สินค้า และ จำนวน มาเช็ก
                if 'สินค้า' in df_bev.columns and 'จำนวน' in df_bev.columns:
                    # รวมยอดแก้วแต่ละเมนู
                    menu_qty = df_bev.groupby('สินค้า')['จำนวน'].sum().reset_index()
                    
                    # บันทึกลงฐานข้อมูล SQLite
                    for index, row in menu_qty.iterrows():
                        c.execute("INSERT INTO beverage_history (date, menu, qty) VALUES (?, ?, ?)", 
                                  (str(record_date), str(row['สินค้า']), float(row['จำนวน'])))
                    conn.commit()

                st.success(f"✅ บันทึกข้อมูลของวันที่ {record_date.strftime('%d/%m/%Y')} ลงฐานข้อมูลสำเร็จ!")
                
                # --- แสดง Dashboard เบื้องต้นให้ดูบนเว็บเลย ---
                st.subheader("📊 สรุปจำนวนแก้ว (อัปเดตล่าสุด)")
                st.dataframe(menu_qty, use_container_width=True)

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    else:
        st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ไฟล์ครับ")

# ปิดการเชื่อมต่อฐานข้อมูล
conn.close()
