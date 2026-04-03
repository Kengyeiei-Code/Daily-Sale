import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Sam Roasters - Auto Report", page_icon="☕", layout="centered")

st.title("☕ ระบบสร้างรายงานอัตโนมัติ")
st.markdown("โยนไฟล์รายงานของเมื่อวาน และไฟล์ POS ของวันนี้ ระบบจะหยอดข้อมูลลงช่องและคำนวณให้ทันที")

with st.form("upload_form"):
    st.info("📝 1. ไฟล์ตั้งต้น (Template)")
    template_file = st.file_uploader("โยนไฟล์รายงาน Excel ของเมื่อวาน (ไฟล์ที่มีฟอร์ม Template)", type=['xlsx'])
    
    st.info("📊 2. ไฟล์ข้อมูลจาก POS ของวันนี้")
    sale_file = st.file_uploader("ไฟล์ยอดรวมและส่วนลด (SaleReport)", type=['csv', 'xlsx'])
    bev_file = st.file_uploader("ไฟล์จำนวนแก้ว (SaleByBehavior)", type=['csv', 'xlsx'])
    
    submit_button = st.form_submit_button("ประมวลผลและสร้างไฟล์ Excel 🚀")

if submit_button:
    if not template_file or not sale_file or not bev_file:
        st.error("❌ กรุณาโยนไฟล์ให้ครบทั้ง 3 ช่องครับ")
    else:
        with st.spinner("กำลังแกะไฟล์และหยอดข้อมูลลง Template..."):
            try:
                # --- 1. สกัดข้อมูลจาก POS ของวันนี้ ---
                data_map = {}
                
                # 1.1 ไฟล์ SaleByBehavior
                if bev_file.name.endswith('.csv'):
                    df_bev = pd.read_csv(bev_file)
                else:
                    df_bev = pd.read_excel(bev_file)
                
                item_col = next((c for c in df_bev.columns if 'สินค้า' in str(c)), None)
                qty_col = next((c for c in df_bev.columns if 'จำนวน' in str(c) and 'บิล' not in str(c)), None)
                type_col = next((c for c in df_bev.columns if 'ปรเะภท' in str(c) or 'ประเภท' in str(c) and 'สินค้า' not in str(c)), None)
                del_col = next((c for c in df_bev.columns if 'Delivery' in str(c)), None)
                net_col = next((c for c in df_bev.columns if 'ยอดสุทธิ' in str(c) and 'Incentive' not in str(c)), None)

                for _, row in df_bev.iterrows():
                    # รวมยอดจำนวนแก้ว
                    if item_col and pd.notna(row[item_col]):
                        item_name = str(row[item_col]).strip().lower()
                        qty = float(str(row[qty_col]).replace(',', '')) if qty_col and pd.notna(row[qty_col]) else 0
                        data_map[item_name] = data_map.get(item_name, 0) + qty

                    # รวมยอดเงินตามประเภท (EatIn, Grab, etc.)
                    if type_col and pd.notna(row[type_col]):
                        type_name = str(row[type_col]).strip().lower()
                        del_name = str(row[del_col]).strip().lower() if del_col and pd.notna(row.get(del_col)) else ""
                        net_val = float(str(row[net_col]).replace(',', '')) if net_col and pd.notna(row[net_col]) else 0
                        
                        data_map[type_name] = data_map.get(type_name, 0) + net_val
                        if type_name in ['delivery', 'grab', 'lineman'] and del_name:
                            data_map[del_name] = data_map.get(del_name, 0) + net_val

                # 1.2 ไฟล์ SaleReport (ดึงเฉพาะส่วนลด)
                if sale_file.name.endswith('.csv'):
                    df_sale = pd.read_csv(sale_file, header=None)
                else:
                    df_sale = pd.read_excel(sale_file, header=None)
                    
                discount_section = False
                for _, row in df_sale.iterrows():
                    col0 = str(row[0]).strip()
                    if col0 == "Discount Summary":
                        discount_section = True
                        continue
                    if discount_section:
                        if col0 == "Name" or col0 == "" or str(col0).lower() == "nan":
                            continue
                        amt_str = str(row[1]).replace(',', '') if pd.notna(row[1]) else "0"
                        try:
                            data_map[col0.lower()] = float(amt_str)
                        except:
                            pass

                # --- 2. นำข้อมูลไปหยอดลง Template ที่อัปโหลดมา ---
                wb = openpyxl.load_workbook(template_file)
                
                # 2.1 หยอดลงหน้า จำนวนแก้ว (Template_Beverage)
                if 'Template_Beverage' in wb.sheetnames:
                    ws_bev = wb['Template_Beverage']
                    for r in range(2, ws_bev.max_row + 1):
                        menu_cell = ws_bev.cell(row=r, column=2) # คอลัมน์ B
                        today_cell = ws_bev.cell(row=r, column=3) # คอลัมน์ C
                        yest_cell = ws_bev.cell(row=r, column=4) # คอลัมน์ D

                        if menu_cell.value and not str(menu_cell.value).startswith('='):
                            raw_name = str(menu_cell.value).strip()
                            m_name = raw_name.lower()
                            
                            if raw_name in ["Menu", "QTY"] or "Total" in raw_name:
                                continue
                            if today_cell.data_type == 'f' or yest_cell.data_type == 'f':
                                continue # ข้ามช่องที่เป็นสูตร

                            # ขยับยอดวันนี้ไปเป็นเมื่อวาน แล้วใส่ยอดใหม่
                            current_today = today_cell.value if isinstance(today_cell.value, (int, float)) else 0
                            yest_cell.value = current_today
                            today_cell.value = data_map.get(m_name, 0)

                # 2.2 หยอดลงหน้า ยอดขายและส่วนลด (Template_Sale)
                if 'Template_Sale' in wb.sheetnames:
                    ws_sale = wb['Template_Sale']
                    for r in range(2, ws_sale.max_row + 1):
                        cat_cell = ws_sale.cell(row=r, column=1) # คอลัมน์ A
                        today_cell = ws_sale.cell(row=r, column=2) # คอลัมน์ B
                        yest_cell = ws_sale.cell(row=r, column=4) # คอลัมน์ D

                        if cat_cell.value and not str(cat_cell.value).startswith('='):
                            raw_name = str(cat_cell.value).strip()
                            c_name = raw_name.lower()
                            
                            if raw_name in ["Category", "Sales - Discount"] or "Total" in raw_name or "Diff" in raw_name:
                                continue
                            if today_cell.data_type == 'f' or yest_cell.data_type == 'f':
                                continue # ข้ามช่องที่เป็นสูตร

                            # ขยับยอดวันนี้ไปเป็นเมื่อวาน แล้วใส่ยอดใหม่
                            current_today = today_cell.value if isinstance(today_cell.value, (int, float)) else 0
                            yest_cell.value = current_today
                            today_cell.value = data_map.get(c_name, 0)

                # --- 3. เตรียมไฟล์ Excel ให้ดาวน์โหลด ---
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("✅ อัปเดตข้อมูลและคำนวณสูตรลง Template เสร็จสมบูรณ์!")
                st.download_button(
                    label="📥 กดดาวน์โหลดไฟล์ Excel (ที่อัปเดตแล้ว)",
                    data=output,
                    file_name="SamRoasters_Daily_Report_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")
