import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Sam Roasters - Auto Report", page_icon="☕", layout="centered")

st.title("☕ ระบบสร้างรายงานอัตโนมัติ")
st.markdown("โยนไฟล์รายงานของเมื่อวาน และไฟล์ POS ของวันนี้ ระบบจะหยอดข้อมูลลงช่องและคำนวณให้ทันที")

with st.form("upload_form"):
    st.info("📝 1. ไฟล์ตั้งต้น (Template)")
    template_file = st.file_uploader("โยนไฟล์รายงาน Excel ของเมื่อวาน (ไฟล์ที่มีฟอร์ม)", type=['xlsx'])
    
    st.info("📊 2. ไฟล์ข้อมูลจาก POS ของวันนี้")
    sale_file = st.file_uploader("ไฟล์ยอดรวมและส่วนลด (SaleReport)", type=['csv', 'xlsx'])
    bev_file = st.file_uploader("ไฟล์จำนวนแก้ว (SaleByBehavior)", type=['csv', 'xlsx'])
    
    submit_button = st.form_submit_button("ประมวลผลและสร้างไฟล์ Excel 🚀")

if submit_button:
    if not template_file or not sale_file or not bev_file:
        st.error("❌ กรุณาโยนไฟล์ให้ครบทั้ง 3 ช่องครับ")
    else:
        with st.spinner("กำลังประมวลผล..."):
            try:
                data_map = {}
                log_messages = [] # ตัวเก็บข้อความรายงานผล

                # --- 1. สกัดข้อมูลจาก POS ของวันนี้ ---
                # 1.1 ไฟล์ SaleByBehavior
                df_bev = pd.read_csv(bev_file) if bev_file.name.endswith('.csv') else pd.read_excel(bev_file)
                item_col = next((c for c in df_bev.columns if 'สินค้า' in str(c)), None)
                qty_col = next((c for c in df_bev.columns if 'จำนวน' in str(c) and 'บิล' not in str(c)), None)
                type_col = next((c for c in df_bev.columns if 'ปรเะภท' in str(c) or 'ประเภท' in str(c) and 'สินค้า' not in str(c)), None)
                del_col = next((c for c in df_bev.columns if 'Delivery' in str(c)), None)
                net_col = next((c for c in df_bev.columns if 'ยอดสุทธิ' in str(c) and 'Incentive' not in str(c)), None)

                for _, row in df_bev.iterrows():
                    if item_col and pd.notna(row[item_col]):
                        item_name = str(row[item_col]).strip().lower()
                        qty = float(str(row[qty_col]).replace(',', '')) if qty_col and pd.notna(row[qty_col]) else 0
                        data_map[item_name] = data_map.get(item_name, 0) + qty

                    if type_col and pd.notna(row[type_col]):
                        type_name = str(row[type_col]).strip().lower()
                        del_name = str(row[del_col]).strip().lower() if del_col and pd.notna(row.get(del_col)) else ""
                        net_val = float(str(row[net_col]).replace(',', '')) if net_col and pd.notna(row[net_col]) else 0
                        
                        data_map[type_name] = data_map.get(type_name, 0) + net_val
                        if type_name in ['delivery', 'grab', 'lineman'] and del_name:
                            data_map[del_name] = data_map.get(del_name, 0) + net_val

                # 1.2 ไฟล์ SaleReport (ส่วนลด)
                df_sale = pd.read_csv(sale_file, header=None) if sale_file.name.endswith('.csv') else pd.read_excel(sale_file, header=None)
                discount_section = False
                for _, row in df_sale.iterrows():
                    col0 = str(row[0]).strip()
                    if col0 == "Discount Summary":
                        discount_section = True
                        continue
                    if discount_section and col0 and col0 != "Name" and str(col0).lower() != "nan":
                        amt_str = str(row[1]).replace(',', '') if pd.notna(row[1]) else "0"
                        try:
                            data_map[col0.lower()] = float(amt_str)
                        except:
                            pass

                # --- 2. นำข้อมูลไปหยอดลง Template ---
                wb = openpyxl.load_workbook(template_file)
                
                # ระบบค้นหาหน้าชีตอัตโนมัติ
                ws_bev = next((wb[sn] for sn in wb.sheetnames if 'bev' in sn.lower()), None)
                ws_sale = next((wb[sn] for sn in wb.sheetnames if 'sale' in sn.lower()), None)

                # 2.1 หยอดลงหน้า จำนวนแก้ว
                if ws_bev:
                    log_messages.append("✅ **พบหน้าจำนวนแก้ว:** กำลังอัปเดตข้อมูล...")
                    updated_bev = 0
                    for r in range(2, ws_bev.max_row + 1):
                        menu_cell = ws_bev.cell(row=r, column=2) # B
                        today_cell = ws_bev.cell(row=r, column=3) # C
                        yest_cell = ws_bev.cell(row=r, column=4) # D

                        if menu_cell.value and not str(menu_cell.value).startswith('='):
                            raw_name = str(menu_cell.value).strip()
                            m_name = raw_name.lower()
                            
                            if raw_name in ["Menu", "QTY"] or "Total" in raw_name:
                                continue
                            if today_cell.data_type == 'f' or yest_cell.data_type == 'f':
                                continue 

                            current_today = today_cell.value if isinstance(today_cell.value, (int, float)) else 0
                            yest_cell.value = current_today
                            
                            new_val = data_map.get(m_name, 0)
                            today_cell.value = new_val
                            
                            if new_val > 0:
                                updated_bev += 1
                    
                    log_messages.append(f"-> อัปเดตเมนูเครื่องดื่มไปทั้งหมด {updated_bev} รายการ")
                else:
                    log_messages.append("❌ **ไม่พบหน้าจำนวนแก้ว** (หาชื่อชีตที่มีคำว่า Bev ไม่เจอ)")

                # 2.2 หยอดลงหน้า ยอดขาย
                if ws_sale:
                    log_messages.append("✅ **พบหน้ายอดขาย:** กำลังอัปเดตข้อมูล...")
                    updated_sale = 0
                    for r in range(2, ws_sale.max_row + 1):
                        cat_cell = ws_sale.cell(row=r, column=1) # A
                        today_cell = ws_sale.cell(row=r, column=2) # B
                        yest_cell = ws_sale.cell(row=r, column=4) # D

                        if cat_cell.value and not str(cat_cell.value).startswith('='):
                            raw_name = str(cat_cell.value).strip()
                            c_name = raw_name.lower()
                            
                            if raw_name in ["Category", "Sales - Discount"] or "Total" in raw_name or "Diff" in raw_name:
                                continue
                            if today_cell.data_type == 'f' or yest_cell.data_type == 'f':
                                continue 

                            current_today = today_cell.value if isinstance(today_cell.value, (int, float)) else 0
                            yest_cell.value = current_today
                            
                            new_val = data_map.get(c_name, 0)
                            today_cell.value = new_val
                            
                            if new_val > 0:
                                updated_sale += 1
                                
                    log_messages.append(f"-> อัปเดตยอดขายและส่วนลดไปทั้งหมด {updated_sale} หมวดหมู่")
                else:
                    log_messages.append("❌ **ไม่พบหน้ายอดขาย** (หาชื่อชีตที่มีคำว่า Sale ไม่เจอ)")

                # --- 3. เตรียมไฟล์ Excel ให้ดาวน์โหลด ---
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("🎉 ประมวลผลและสร้างไฟล์ Excel เสร็จสมบูรณ์แล้ว!")
                
                # แสดงข้อความรายงานผลให้เห็นชัดๆ
                for msg in log_messages:
                    st.write(msg)

                st.download_button(
                    label="📥 กดดาวน์โหลดไฟล์ Excel ตรงนี้ครับ",
                    data=output,
                    file_name="SamRoasters_Daily_Report_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")
