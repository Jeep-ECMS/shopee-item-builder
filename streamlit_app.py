import streamlit as st
import openpyxl
import io
import itertools

st.set_page_config(page_title="Shopee Mass Upload Generator", layout="wide")
st.title("📦 เครื่องมือสร้างไฟล์ Mass Upload สำหรับ Shopee")

# 1. ฟอร์มกรอกข้อมูลสินค้าหลัก
with st.form("product_form"):
    st.subheader("1. ข้อมูลสินค้าหลัก")
    col1, col2, col3 = st.columns(3)
    with col1:
        category_id = st.text_input("รหัสหมวดหมู่ (Category ID)", value="120039")
        parent_sku = st.text_input("Parent SKU / รหัสอ้างอิงหลัก", value="SHIRT-001")
        brand = st.text_input("แบรนด์ (Brand ID หรือ No Brand)", value="2200345")
    with col2:
        product_name = st.text_input("ชื่อสินค้า", value="เสื้อยืดคอตตอนผ้านุ่มพิเศษ")
        weight = st.number_input("น้ำหนักสินค้า (kg)", value=0.2, step=0.01)
        cover_image = st.text_input("URL รูปภาพหลัก (Cover Image)", value="https://example.com/cover.jpg")
    with col3:
        product_desc = st.text_area("รายละเอียดสินค้า", value="เสื้อยืดคุณภาพดี ใส่สบาย ระบายอากาศได้ดี")

    st.markdown("---")
    st.subheader("2. ตัวเลือกสินค้า (Variations)")
    
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        v1_name = st.text_input("ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)", value="สี")
        v1_options = st.text_input("รายการตัวเลือกที่ 1 (คั่นด้วยเครื่องหมายจุลภาค ,)", value="แดง, ดำ, ขาว")
    
    with col_v2:
        v2_name = st.text_input("ชื่อตัวเลือกที่ 2 (เช่น ไซส์) [เว้นว่างไว้ได้]", value="ไซส์")
        v2_options = st.text_input("รายการตัวเลือกที่ 2 (คั่นด้วยเครื่องหมายจุลภาค ,)", value="S, M, L")

    st.markdown("---")
    st.subheader("3. ราคา สต๊อก และค่าขนส่ง")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        default_price = st.number_input("ราคาเริ่มต้น (บาท)", value=150.0)
    with col_p2:
        default_stock = st.number_input("จำนวนสต๊อกเริ่มต้น", value=50)
    with col_p3:
        channel_on = st.selectbox("เปิดช่องทางการจัดส่งหลัก", ["On", "Off"], index=0)

    submit = st.form_submit_button("🚀 สร้างไฟล์ Excel สำหรับ Shopee")

# 2. ส่วนการประมวลผลเมื่อกดปุ่ม
if submit:
    list_v1 = [x.strip() for x in v1_options.split(",") if x.strip()]
    list_v2 = [x.strip() for x in v2_options.split(",") if x.strip()] if v2_name else [""]

    variations = list(itertools.product(list_v1, list_v2))
    
    template_path = "Shopee_template.xlsx"
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Template"]

    start_row = 6 

    for idx, (opt1, opt2) in enumerate(variations):
        current_row = start_row + idx
        sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")
        
        ws.cell(row=current_row, column=1, value=category_id)
        ws.cell(row=current_row, column=2, value=product_name if idx == 0 else "")
        ws.cell(row=current_row, column=3, value=product_desc if idx == 0 else "")
        ws.cell(row=current_row, column=9, value=parent_sku)
        ws.cell(row=current_row, column=10, value=parent_sku)
        
        ws.cell(row=current_row, column=11, value=v1_name)
        ws.cell(row=current_row, column=12, value=opt1)
        
        if v2_name and opt2:
            ws.cell(row=current_row, column=14, value=v2_name)
            ws.cell(row=current_row, column=15, value=opt2)
            
        ws.cell(row=current_row, column=16, value=default_price)
        ws.cell(row=current_row, column=17, value=default_stock)
        ws.cell(row=current_row, column=18, value=f"{parent_sku}{sku_suffix}")
        
        ws.cell(row=current_row, column=21, value=cover_image)
        ws.cell(row=current_row, column=30, value=weight)
        
        ws.cell(row=current_row, column=34, value=channel_on)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"✅ สร้างไฟล์สำเร็จ! สร้างรายการตัวเลือกทั้งหมด {len(variations)} รายการ")
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel พร้อมอัปโหลด",
        data=output,
        file_name=f"Shopee_Mass_Upload_{parent_sku}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )