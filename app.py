import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="SKU Generator Pro", 
    page_icon="📦", 
    layout="wide"
)

st.title("📦 ระบบสร้างรหัส SKU สินค้าตามช่องทางการขาย")
st.write("จัดการสร้างรหัสสินค้าแยกตามช่องทางการขายอัตโนมัติ รวดเร็ว แม่นยำ")
st.write("---")

channel_map = {
    'Eat in': {'code': '01', 'abbr': 'EI'},
    'Take away': {'code': '02', 'abbr': 'TA'},
    'Delivery': {'code': '03', 'abbr': 'DL'},
	'Grab food': {'code': '04', 'abbr': 'GF'},
	'Line man': {'code': '05', 'abbr': 'LM'},
	'FoodPanda': {'code': '06', 'abbr': 'FP'},
	'RobinHood': {'code': '07', 'abbr': 'RH'},
	'GoJek': {'code': '08', 'abbr': 'GJ'},
	'Down Payment': {'code': '09', 'abbr': 'DE'},
	'Car Stop': {'code': '10', 'abbr': 'CS'},
	'OHO': {'code': '11', 'abbr': 'OH'},
	'Shopee Food': {'code': '12', 'abbr': 'SF'},
	'Pinto Pum': {'code': '13', 'abbr': 'PP'},
	'E GET': {'code': '14', 'abbr': 'EG'},
	'Airasia Food': {'code': '15', 'abbr': 'AF'},
	'SOOK Shop': {'code': '16', 'abbr': 'SK'}
}

# ฟังก์ชันตกแต่งหัวข้อตาราง ขยายขนาดคอลัมน์อัตโนมัติ และตรึงแนวแถวบนสุด
def style_excel_sheet(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in worksheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            
    for idx, col in enumerate(worksheet.columns, start=1):
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(idx)
        worksheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

tab1, tab2 = st.tabs(["📝 กรอกข้อมูลเองแบบหลายสินค้า (Manual Entry)", "📥 อัปโหลดไฟล์ Excel (Import)"])

# ==========================================
# แท็บที่ 1: กรอกข้อมูลเอง
# ==========================================
with tab1:
    st.subheader("📝 ระบุข้อมูลสินค้าแยกตามรายการ")
    
    if 'product_items' not in st.session_state:
        st.session_state.product_items = [{'id': 0, 'sku': '', 'name': '', 'channels': ['Eat in']}]
    if 'next_id' not in st.session_state:
        st.session_state.next_id = 1

    if st.button("➕ เพิ่มรายการสินค้า", key="add_item_btn"):
        st.session_state.product_items.append({
            'id': st.session_state.next_id,
            'sku': '',
            'name': '',
            'channels': ['Eat in']
        })
        st.session_state.next_id += 1
        st.rerun()

    st.write("")
    
    updated_items = []
    
    for idx, item in enumerate(st.session_state.product_items):
        item_id = item['id']
        
        c1, c2, c3, c4 = st.columns([3, 3, 4, 1])
        
        with c1:
            sku_val = st.text_input(
                f"รหัส SKU เดิม #{idx+1}", 
                value=item['sku'], 
                key=f"sku_in_{item_id}", 
                placeholder="เช่น PROD001",
                label_visibility="collapsed" if idx > 0 else "visible"
            ).strip()
        with c2:
            name_val = st.text_input(
                f"ชื่อสินค้า #{idx+1}", 
                value=item['name'], 
                key=f"name_in_{item_id}", 
                placeholder="เช่น ชานม",
                label_visibility="collapsed" if idx > 0 else "visible"
            ).strip()
        with c3:
            # แก้ไขจุดนี้: ถอดคำว่า "ทุกช่องทางการขาย" ออกจากตัวเลือก ให้เหลือแค่รายการหลักในคลังระบบ
            channels_val = st.multiselect(
                f"ช่องทางการขาย #{idx+1}",
                options=list(channel_map.keys()),
                default=item['channels'],
                key=f"channels_in_{item_id}",
                label_visibility="collapsed" if idx > 0 else "visible"
            )
        with c4:
            if idx == 0:
                st.write('<div style="margin-top:28px;"></div>', unsafe_allow_html=True)
                
            if st.button("❌", key=f"del_btn_{item_id}", help="ลบรายการนี้"):
                st.session_state.product_items = [x for x in st.session_state.product_items if x['id'] != item_id]
                if not st.session_state.product_items:
                    st.session_state.product_items = [{'id': st.session_state.next_id, 'sku': '', 'name': '', 'channels': ['Eat in']}]
                    st.session_state.next_id += 1
                st.rerun()
                
        item['sku'] = sku_val
        item['name'] = name_val
        item['channels'] = channels_val
        updated_items.append(item)

    st.session_state.product_items = updated_items
    st.write("")
    
    if st.button("⚡ เริ่มสร้างรหัส SKU ทั้งหมดจากรายการด้านบน", type="primary", key="btn_process_manual"):
        manual_results = []
        has_error = False
        
        for index, item in enumerate(st.session_state.product_items):
            sku_old = item['sku']
            p_name = item['name']
            channels_input = item['channels']
            
            if not sku_old and not p_name:
                continue
                
            if not sku_old or not p_name:
                st.error(f"⚠️ รายการแถวที่ {index + 1} กรอกข้อมูลไม่ครบ กรุณาระบุทั้ง SKU เดิม และ ชื่อสินค้า")
                has_error = True
                break
                
            if not channels_input:
                st.error(f"⚠️ รายการแถวที่ {index + 1} ยังไม่ได้เลือกช่องทางการขาย")
                has_error = True
                break
                
            for channel in channels_input:
                info = channel_map.get(channel)
                if info:
                    new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                    manual_results.append({
                        "New_Product_Code": new_sku,
                        "ชื่อสินค้า": p_name,
                        "ช่องทางการขาย": channel,
                        "_Sort_Code": info['code'],
                        "SKU เดิม": sku_old
                    })
                    
        if not has_error and len(manual_results) > 0:
            df_manual_res = pd.DataFrame(manual_results)
            df_manual_res = df_manual_res.sort_values(by=['SKU เดิม', '_Sort_Code']).reset_index(drop=True)
            
            output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
            df_manual_final = df_manual_res[output_cols]
            
            st.success(f"🎉 สร้างรหัสสินค้าสำเร็จทั้งหมด {len(df_manual_final)} รายการ!")
            
            st.write("📋 **รหัสสินค้าใหม่ทั้งหมด (คลิกไอคอนสี่เหลี่ยมซ้อนกันที่มุมขวาบนของกล่องเพื่อคัดลอกทั้งหมด):**")
            all_new_skus_text = "\n".join(df_manual_final["New_Product_Code"].tolist())
            st.code(all_new_skus_text, language="text")
            
            st.write("📊 **ตารางสรุปรายละเอียด:**")
            st.dataframe(df_manual_final, use_container_width=True)
            
            buffer_manual = io.BytesIO()
            with pd.ExcelWriter(buffer_manual, engine='openpyxl') as writer:
                df_manual_final.to_excel(writer, index=False, sheet_name='Manual_SKU')
                style_excel_sheet(writer, 'Manual_SKU')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (จากรายการที่กรอก)",
                data=buffer_manual.getvalue(),
                file_name="manual_sku_products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_manual_res"
            )
        elif not has_error and len(manual_results) == 0:
            st.warning("⚠️ ไม่พบข้อมูลสินค้า กรุณากรอกข้อมูลสินค้าในช่องด้านบนอย่างน้อย 1 รายการ")

# ==========================================
# แท็บที่ 2: อัปโหลดไฟล์ Excel
# ==========================================
with tab2:
    st.subheader("📥 ดาวน์โหลดไฟล์เทมเพลตสำหรับใช้งาน")
    
    template_data = {
        "SKU เดิม": ["PROD001", "PROD002"],
        "ชื่อสินค้า": ["ชานม", "กาแฟ"]
    }
    df_template = pd.DataFrame(template_data)
    
    buffer_template = io.BytesIO()
    with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Template')
        style_excel_sheet(writer, 'Template')
        
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์เทมเพลต Excel (.xlsx)",
        data=buffer_template.getvalue(),
        file_name="sku_import_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_template_xlsx"
    )
    
    st.write("---")
    st.subheader("📤 นำเข้าไฟล์ข้อมูลเพื่อแตกรหัสสินค้าเป็นทุกช่องทางการขาย")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ (รองรับ .xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("📋 **ตัวอย่างข้อมูลที่นำเข้า:**")
            st.dataframe(df.head())
            columns = df.columns.tolist()
            
            st.subheader("⚙️ ตั้งค่าคอลัมน์ข้อมูล")
            col1, col2 = st.columns(2)
            with col1:
                sku_col = st.selectbox("เลือกคอลัมน์ที่เป็น [SKUเดิม]", options=columns, key="import_sku")
            with col2:
                name_col = st.selectbox("เลือกคอลัมน์ที่เป็น [ชื่อสินค้า]", options=columns, key="import_name")
                
            if st.button("⚡ เริ่มแตกไฟล์และสร้างรหัส SKU ทั้งหมด", type="primary"):
                expanded_rows = []
                
                for _, row in df.iterrows():
                    sku_old = str(row[sku_col]).strip()
                    p_name = str(row[name_col]).strip()
                    
                    for channel, info in channel_map.items():
                        new_sku = f"{sku_old}{info['code']}{info['abbr']}"
                        
                        new_row = {
                            'New_Product_Code': new_sku,
                            'ชื่อสินค้า': p_name,
                            'ช่องทางการขาย': channel,
                            '_Sort_Code': info['code'],
                            '_Sku_Old': sku_old
                        }

                        expanded_rows.append(new_row)
                
                df_result = pd.DataFrame(expanded_rows)
                df_result = df_result.sort_values(by=['_Sku_Old', '_Sort_Code']).reset_index(drop=True)
                
                output_cols = ["New_Product_Code", "ชื่อสินค้า", "ช่องทางการขาย"]
                df_final_export = df_result[output_cols]
                
                st.success(f"🎉 ระบบแตกไฟล์ข้อมูลและสร้างรหัสทุกช่องทางการขายเสร็จเรียบร้อย ทั้งหมด {len(df_final_export)} รายการ!")
                st.write("📋 **ผลลัพธ์ข้อมูลใหม่:**")
                st.dataframe(df_final_export, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final_export.to_excel(writer, index=False, sheet_name='All_Channels_SKU')
                    style_excel_sheet(writer, 'All_Channels_SKU')
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์ (ครบทุกช่องทาง)",
		        data=buffer.getvalue(),
                    file_name="multi_channel_sku_products.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_file_res"
                )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ที่มีคอลัมน์สำหรับ SKU เดิม และ ชื่อสินค้า (ระบบจะนำแต่ละแถวไปสร้างครบทั้ง 3 ช่องทางให้เองอัตโนมัติ)")
