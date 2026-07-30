import streamlit as st
import openpyxl
import io
import itertools
import pandas as pd

st.set_page_config(page_title="Shopee Mass Upload Generator (Multi-Product)", layout="wide")
st.title("📦 เครื่องมือสร้างไฟล์ Mass Upload Shopee (หลายสินค้า & กำหนดราคาแยกได้)")

# จัดเก็บรายการสินค้าใน Session State
if "products" not in st.sessions:
    st.session_state.products = [
        {
            "category_id": "120039",
            "parent_sku": "SHIRT-001",
            "brand": "No Brand",
            "product_name": "เสื้อยืดคอตตอนผ้านุ่มพิเศษ",
            "weight": 0.2,
            "product_desc": "เสื้อยืดคุณภาพดี ใส่สบาย",
            "cover_image": "https://example.com/shirt_cover.jpg",
            "v1_name": "สี",
            "v1_options": "แดง, ดำ",
            "v1_images": "https://example.com/red.jpg, https://example.com/black.jpg",
            "v2_name": "ไซส์",
            "v2_options": "S, M",
        }
    ]

# ปุ่มเพิ่ม / ลบ รายการสินค้า
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button("➕ เพิ่มสินค้าชิ้นใหม่"):
        st.session_state.products.append({
            "category_id": "100000",
            "parent_sku": f"ITEM-{len(st.session_state.products)+1:03d}",
            "brand": "No Brand",
            "product_name": f"สินค้าชิ้นที่ {len(st.session_state.products)+1}",
            "weight": 0.5,
            "product_desc": "รายละเอียดสินค้า...",
            "cover_image": "https://example.com/cover.jpg",
            "v1_name": "ตัวเลือก",
            "v1_options": "แบบ A, แบบ B",
            "v1_images": "",
            "v2_name": "",
            "v2_options": "",
        })
        st.rerun()

# ฟอร์มกรอกข้อมูลของสินค้าแต่ละตัว
updated_products_data = []

for idx, p in enumerate(st.session_state.products):
    st.markdown("---")
    col_title, col_del = st.columns([8, 2])
    with col_title:
        st.subheader(f"🛒 สินค้าชิ้นที่ {idx + 1}: {p['product_name']}")
    with col_del:
        if len(st.session_state.products) > 1:
            if st.button(f"🗑️ ลบสินค้านี้", key=f"del_{idx}"):
                st.session_state.products.pop(idx)
                st.rerun()

    # 1. ข้อมูลหลัก
    c1, c2, c3 = st.columns(3)
    with c1:
        cat_id = st.text_input("Category ID", value=p["category_id"], key=f"cat_{idx}")
        p_sku = st.text_input("Parent SKU", value=p["parent_sku"], key=f"psku_{idx}")
        brand = st.text_input("แบรนด์ (Brand)", value=p["brand"], key=f"brand_{idx}")
    with c2:
        p_name = st.text_input("ชื่อสินค้า", value=p["product_name"], key=f"name_{idx}")
        weight = st.number_input("น้ำหนัก (kg)", value=p["weight"], step=0.01, key=f"w_{idx}")
    with c3:
        p_desc = st.text_area("รายละเอียดสินค้า", value=p["product_desc"], key=f"desc_{idx}")

    # 2. รูปภาพปก
    cover_img = st.text_input("URL รูปภาพปกหลัก (Cover Image)", value=p["cover_image"], key=f"cimg_{idx}")

    # 3. ตัวเลือก Variation
    cv1, cv2 = st.columns(2)
    with cv1:
        v1_name = st.text_input("ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)", value=p["v1_name"], key=f"v1n_{idx}")
        v1_opts = st.text_input("รายการตัวเลือกที่ 1 (คั่นด้วย ,)", value=p["v1_options"], key=f"v1o_{idx}")
        v1_imgs = st.text_input("URL รูปภาพตัวเลือกที่ 1 (คั่นด้วย ,)", value=p["v1_images"], key=f"v1i_{idx}")
    with cv2:
        v2_name = st.text_input("ชื่อตัวเลือกที่ 2 (เช่น ไซส์) [เว้นว่างได้]", value=p["v2_name"], key=f"v2n_{idx}")
        v2_opts = st.text_input("รายการตัวเลือกที่ 2 (คั่นด้วย ,)", value=p["v2_options"], key=f"v2o_{idx}")

    # คำนวณตาราง Variation เพื่อให้ผู้ใช้กรอกราคา/สต๊อก/SKU แยกตามรายการ
    list_v1 = [x.strip() for x in v1_opts.split(",") if x.strip()]
    list_v2 = [x.strip() for x in v2_opts.split(",") if x.strip()] if v2_name else [""]
    variations = list(itertools.product(list_v1, list_v2))

    # สร้างข้อมูลเริ่มต้นใส่ DataFrame ให้แก้ไขง่ายๆ
    grid_data = []
    for opt1, opt2 in variations:
        var_title = f"{opt1}" + (f" / {opt2}" if opt2 else "")
        sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")
        grid_data.append({
            "Variation": var_title,
            "SKU": f"{p_sku}{sku_suffix}",
            "Price (บาท)": 150.0,
            "Stock (ชิ้น)": 50,
            "Opt1": opt1,
            "Opt2": opt2
        })
    
    df_var = pd.DataFrame(grid_data)
    
    st.write("💰 **กำหนดราคา สต๊อก และ SKU สำหรับแต่ละ Variation:**")
    edited_df = st.data_editor(
        df_var,
        column_config={
            "Variation": st.column_config.Column(disabled=True),
            "Opt1": None,  # ซ่อนคอลัมน์ระบบ
            "Opt2": None
        },
        hide_index=True,
        key=f"editor_{idx}"
    )

    updated_products_data.append({
        "cat_id": cat_id,
        "p_sku": p_sku,
        "brand": brand,
        "p_name": p_name,
        "weight": weight,
        "p_desc": p_desc,
        "cover_img": cover_img,
        "v1_name": v1_name,
        "v1_imgs": [x.strip() for x in v1_imgs.split(",") if x.strip()],
        "v2_name": v2_name,
        "variations_table": edited_df
    })

# --- ส่วนของการสร้างไฟล์ EXCEL ---
st.markdown("---")
if st.button("🚀 สร้างไฟล์ Excel รวมทุกสินค้าสำหรับ Shopee", type="primary", use_container_width=True):
    try:
        wb = openpyxl.load_workbook("Shopee_template.xlsx")
        ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
    except Exception:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template"

    start_row = 6  # บรรทัดแรกที่จะเริ่มเขียนข้อมูลใน Template

    for p_data in updated_products_data:
        cat_id = p_data["cat_id"]
        p_sku = p_data["p_sku"]
        p_name = p_data["p_name"]
        p_desc = p_data["p_desc"]
        weight = p_data["weight"]
        cover_img = p_data["cover_img"]
        v1_name = p_data["v1_name"]
        v2_name = p_data["v2_name"]
        v1_imgs = p_data["v1_imgs"]
        df_vars = p_data["variations_table"]

        # ดึงลิสต์ Option 1 เพื่อแมปกับรูปภาพ
        unique_v1 = df_vars["Opt1"].unique().tolist()
        v1_img_dict = {}
        for i, opt in enumerate(unique_v1):
            v1_img_dict[opt] = v1_imgs[i] if i < len(v1_imgs) else ""

        for idx, row in df_vars.iterrows():
            current_row = start_row
            
            # 1. ข้อมูลพื้นฐานสินค้า (ลงบรรทัดแรกของสินค้านั้นๆ)
            ws.cell(row=current_row, column=1, value=cat_id)
            ws.cell(row=current_row, column=2, value=p_name if idx == 0 else "")
            ws.cell(row=current_row, column=3, value=p_desc if idx == 0 else "")
            ws.cell(row=current_row, column=10, value=p_sku) # Parent SKU
            
            # 2. ข้อมูล Variation & Custom Price/Stock/SKU
            ws.cell(row=current_row, column=11, value=v1_name)
            ws.cell(row=current_row, column=12, value=row["Opt1"])
            ws.cell(row=current_row, column=13, value=v1_img_dict.get(row["Opt1"], ""))
            
            if v2_name and row["Opt2"]:
                ws.cell(row=current_row, column=14, value=v2_name)
                ws.cell(row=current_row, column=15, value=row["Opt2"])
                
            ws.cell(row=current_row, column=16, value=row["Price (บาท)"])
            ws.cell(row=current_row, column=17, value=row["Stock (ชิ้น)"])
            ws.cell(row=current_row, column=18, value=row["SKU"])
            
            # 3. รูปภาพหลักและค่าจัดส่ง (ลงบรรทัดแรก)
            if idx == 0:
                ws.cell(row=current_row, column=21, value=cover_img) # Item Cover Image
                ws.cell(row=current_row, column=30, value=weight)
                ws.cell(row=current_row, column=34, value="On")

            start_row += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"✅ สร้างไฟล์สำเร็จ! รวมสินค้าทั้งหมด {len(updated_products_data)} รายการ")
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel พร้อมอัปโหลด Shopee",
        data=output,
        file_name="Shopee_Mass_Upload_MultiProduct.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )