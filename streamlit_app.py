import streamlit as st
import openpyxl
import io
import itertools

st.set_page_config(page_title="Shopee Mass Upload Generator", layout="wide")
st.title("📦 เครื่องมือสร้างไฟล์ Mass Upload สำหรับ Shopee (กำหนด SKU เองได้)")

with st.form("product_form"):
    st.subheader("1. ข้อมูลสินค้าหลัก")
    col1, col2, col3 = st.columns(3)
    with col1:
        category_id = st.text_input("รหัสหมวดหมู่ (Category ID)", value="120039")
        parent_sku = st.text_input("Parent SKU / รหัสอ้างอิงหลัก", value="SHIRT-001")
        brand = st.text_input("แบรนด์ (Brand ID หรือ No Brand)", value="No Brand")
    with col2:
        product_name = st.text_input("ชื่อสินค้า", value="เสื้อยืดคอตตอนผ้านุ่มพิเศษ")
        weight = st.number_input("น้ำหนักสินค้า (kg)", value=0.2, step=0.01)
    with col3:
        product_desc = st.text_area("รายละเอียดสินค้า", value="เสื้อยืดคุณภาพดี ใส่สบาย ระบายอากาศได้ดี")

    st.markdown("---")
    st.subheader("2. รูปภาพสินค้าหลัก (สูงสุด 9 รูป)")
    col_img1, col_img2 = st.columns(2)
    with col_img1:
        cover_image = st.text_input("URL รูปภาพปกหลัก (Cover Image - บังคับ)", value="https://example.com/cover.jpg")
        img1 = st.text_input("URL รูปภาพประกอบ 1 (Item Image 1)", value="")
        img2 = st.text_input("URL รูปภาพประกอบ 2 (Item Image 2)", value="")
        img3 = st.text_input("URL รูปภาพประกอบ 3 (Item Image 3)", value="")
        img4 = st.text_input("URL รูปภาพประกอบ 4 (Item Image 4)", value="")
    with col_img2:
        img5 = st.text_input("URL รูปภาพประกอบ 5 (Item Image 5)", value="")
        img6 = st.text_input("URL รูปภาพประกอบ 6 (Item Image 6)", value="")
        img7 = st.text_input("URL รูปภาพประกอบ 7 (Item Image 7)", value="")
        img8 = st.text_input("URL รูปภาพประกอบ 8 (Item Image 8)", value="")

    st.markdown("---")
    st.subheader("3. ตัวเลือกสินค้า (Variations) & SKU")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        v1_name = st.text_input("ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)", value="สี")
        v1_options = st.text_input("รายการตัวเลือกที่ 1 (คั่นด้วย ,)", value="แดง, ดำ, ขาว")
        v1_images = st.text_input("URL รูปภาพตัวเลือกที่ 1 (คั่นด้วย , ตรงตามจำนวนตัวเลือก)", value="https://example.com/red.jpg, https://example.com/black.jpg, https://example.com/white.jpg")
    
    with col_v2:
        v2_name = st.text_input("ชื่อตัวเลือกที่ 2 (เช่น ไซส์) [เว้นว่างไว้ได้]", value="ไซส์")
        v2_options = st.text_input("รายการตัวเลือกที่ 2 (คั่นด้วย ,)", value="S, M, L")

    st.subheader("🔑 ระบุ SKU ของแต่ละ Variation (แยกตามลำดับด้วยเครื่องหมาย ,)")
    custom_skus_input = st.text_area(
        "กรอกรายการ SKU ของแต่ละ Variation (คั่นด้วย ,)",
        value="SHIRT-RED-S, SHIRT-RED-M, SHIRT-RED-L, SHIRT-BLK-S, SHIRT-BLK-M, SHIRT-BLK-L, SHIRT-WHT-S, SHIRT-WHT-M, SHIRT-WHT-L",
        help="เรียงลำดับ SKU ตามรายการ Variation ที่จะเกิดขึ้น (ถ้าเว้นว่างไว้ ระบบจะใช้ Parent SKU เป็นหลัก)"
    )

    st.markdown("---")
    st.subheader("4. ราคา สต๊อก และค่าขนส่ง")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        default_price = st.number_input("ราคาเริ่มต้น (บาท)", value=150.0)
    with col_p2:
        default_stock = st.number_input("จำนวนสต๊อกเริ่มต้น", value=50)
    with col_p3:
        channel_on = st.selectbox("เปิดช่องทางการจัดส่งหลัก", ["On", "Off"], index=0)

    submit = st.form_submit_button("🚀 สร้างไฟล์ Excel สำหรับ Shopee")

if submit:
    list_v1 = [x.strip() for x in v1_options.split(",") if x.strip()]
    list_v1_imgs = [x.strip() for x in v1_images.split(",") if x.strip()]
    list_v2 = [x.strip() for x in v2_options.split(",") if x.strip()] if v2_name else [""]
    
    # รายการ Custom SKU ที่ผู้ใช้ระบุเอง
    custom_skus = [x.strip() for x in custom_skus_input.split(",") if x.strip()]

    # สร้าง Dictionary จับคู่รูปกับ Variation 1
    v1_img_dict = {}
    for i, opt in enumerate(list_v1):
        if i < len(list_v1_imgs):
            v1_img_dict[opt] = list_v1_imgs[i]
        else:
            v1_img_dict[opt] = ""

    variations = list(itertools.product(list_v1, list_v2))
    
    # โหลดไฟล์ Template
    try:
        wb = openpyxl.load_workbook("Shopee_template.xlsx")
        ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
    except Exception:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template"

    # กำหนดบรรทัดเริ่มต้นสำหรับลงข้อมูล (เริ่มบรรทัดที่ 6)
    start_row = 6 

    for idx, (opt1, opt2) in enumerate(variations):
        current_row = start_row + idx
        
        # ดึง SKU ที่กรอกมาใช้ (ถ้ากรอกไม่ครบหรือเว้นว่างไว้ จะใช้ Auto-fallback ให้)
        if idx < len(custom_skus):
            variation_sku = custom_skus[idx]
        else:
            sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")
            variation_sku = f"{parent_sku}{sku_suffix}"
        
        # 1. ข้อมูลพื้นฐานสินค้า (ใส่เฉพาะบรรทัดแรก)
        ws.cell(row=current_row, column=1, value=category_id)
        ws.cell(row=current_row, column=2, value=product_name if idx == 0 else "")
        ws.cell(row=current_row, column=3, value=product_desc if idx == 0 else "")
        
        # 2. Variation Integration No. / Parent SKU (คอลัมน์ J)
        ws.cell(row=current_row, column=10, value=parent_sku)
        
        # 3. ข้อมูล Variations (คอลัมน์ K ถึง R)
        ws.cell(row=current_row, column=11, value=v1_name)                     # K: Variation Name 1
        ws.cell(row=current_row, column=12, value=opt1)                       # L: Option for Variation 1
        ws.cell(row=current_row, column=13, value=v1_img_dict.get(opt1, "")) # M: Image per Variation
        
        if v2_name and opt2:
            ws.cell(row=current_row, column=14, value=v2_name)                 # N: Variation Name 2
            ws.cell(row=current_row, column=15, value=opt2)                   # O: Option for Variation 2
            
        ws.cell(row=current_row, column=16, value=default_price)              # P: Price
        ws.cell(row=current_row, column=17, value=default_stock)              # Q: Stock
        ws.cell(row=current_row, column=18, value=variation_sku)              # R: SKU แต่ละ Variation ที่ใส่เอง
        
        # 4. ข้อมูลรูปภาพหลัก 9 รูป (คอลัมน์ U ถึง AC - ใส่บรรทัดแรก)
        if idx == 0:
            main_images = [cover_image, img1, img2, img3, img4, img5, img6, img7, img8]
            for col_idx, img_url in enumerate(main_images):
                if img_url:
                    ws.cell(row=current_row, column=21 + col_idx, value=img_url)
                    
            # น้ำหนักสินค้า (คอลัมน์ AD หรือ 30)
            ws.cell(row=current_row, column=30, value=weight)
            # ช่องทางการจัดส่ง (คอลัมน์ AH หรือ 34)
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