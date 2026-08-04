import streamlit as st
import openpyxl
import io
import itertools
import math
import pandas as pd

st.set_page_config(page_title="Shopee Auto Price & Mass Upload Builder", layout="wide")

# --- 1. LOGIC การคำนวณค่าขนส่ง SLS & ภาษีศุลกากร (CUSTOMS & TRANSPORTATION) ---

SLS_RATES_THB = [
    (100, 52), (200, 76), (300, 100), (400, 124), (500, 148),
    (600, 172), (700, 196), (800, 220), (900, 244), (1000, 268),
    (1500, 388), (2000, 508), (2500, 628), (3000, 748), (3500, 868), (4000, 988)
]

def get_sls_shipping_fee_thb(weight_g):
    for max_weight, fee in SLS_RATES_THB:
        if weight_g <= max_weight:
            return fee
    return 988 + math.ceil((weight_g - 4000) / 500) * 120

def get_sls_shipping_fee_php(weight_g):
    if weight_g <= 50:
        normal_fee = 6
    else:
        steps = math.ceil((weight_g - 50) / 50)
        normal_fee = 6 + (steps * 25)
    esf_zone_a = 50
    return esf_zone_a + normal_fee

def calculate_net_price(buying_price_jpy, weight_kg, currency="THB", rate_jpy=None):
    try:
        buying_price_jpy = float(buying_price_jpy)
        weight_g = float(weight_kg) * 1000
    except (ValueError, TypeError):
        return 0

    if buying_price_jpy <= 0:
        return 0
    if weight_g <= 0:
        weight_g = 100

    if currency == "THB":
        rate = rate_jpy if rate_jpy else 4.868196
        buying_price_thb = buying_price_jpy / rate
        transportation_jp_thb = 70
        sls_fee = get_sls_shipping_fee_thb(weight_g)
        net_price = (sls_fee + buying_price_thb + transportation_jp_thb) / 0.65
        return round(net_price)

    elif currency == "PHP":
        rate = rate_jpy if rate_jpy else 2.590111
        buying_price_php = buying_price_jpy / rate
        transportation_jp_php = 300 / rate
        sls_fee = get_sls_shipping_fee_php(weight_g)
        
        temp_price = (sls_fee + buying_price_php + transportation_jp_php) / 0.70
        cif_value = temp_price * 1.01 + (1369 * (weight_g / 1000))
        
        # กฎภาษีศุลกากร De Minimis 10,000 PHP
        if cif_value >= 10000:
            net_price = (sls_fee + (1369 * (weight_g / 1000) * 0.12) + buying_price_php + transportation_jp_php) / 0.5788
        else:
            net_price = temp_price
            
        return round(net_price)

    return 0


# --- 2. MULTI-LANGUAGE DICTIONARY ---
LANG_TEXTS = {
    "TH": {
        "title": "📦 เครื่องมือสร้างไฟล์ Mass Upload Shopee & คำนวณราคาขายอัตโนมัติ",
        "calc_setting": "⚙️ ตั้งค่าการคำนวณราคา (Target Market & Exchange Rate)",
        "currency_select": "เลือกตลาดเป้าหมาย",
        "rate_label": "อัตราแลกเปลี่ยน (JPY)",
        "add_product": "➕ เพิ่มสินค้าชิ้นใหม่",
        "del_product": "🗑️ ลบสินค้านี้",
        "product_num": "🛒 สินค้าชิ้นที่",
        "cat_id": "Category ID / รหัสหมวดหมู่",
        "parent_sku": "Parent SKU / รหัสอ้างอิงหลัก",
        "brand": "แบรนด์ (Brand)",
        "p_name": "ชื่อสินค้า",
        "weight": "น้ำหนักสินค้า (kg)",
        "p_desc": "รายละเอียดสินค้า",
        "cover_img": "URL รูปภาพปกหลัก (Cover Image)",
        "v1_name": "ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)",
        "v1_opts": "รายการตัวเลือกที่ 1 (คั่นด้วย ,)",
        "v1_imgs": "URL รูปภาพตัวเลือกที่ 1 (คั่นด้วย ,)",
        "v2_name": "ชื่อตัวเลือกที่ 2 (เช่น ไซส์) [เว้นว่างได้]",
        "v2_opts": "รายการตัวเลือกที่ 2 (คั่นด้วย ,)",
        "grid_title": "💰 กำหนดราคาซื้อ (JPY) สต๊อก และคำนวณราคาขายอัตโนมัติ:",
        "cost_col": "ราคาซื้อ (JPY)",
        "price_col": "ราคาขายอัตโนมัติ",
        "stock_col": "Stock (ชิ้น)",
        "btn_generate": "🚀 สร้างไฟล์ Excel รวมทุกสินค้าสำหรับ Shopee",
        "success_msg": "✅ สร้างไฟล์สำเร็จ! รวมสินค้าทั้งหมด {count} รายการ",
        "btn_download": "📥 ดาวน์โหลดไฟล์ Excel พร้อมอัปโหลด Shopee",
        "default_pname": "เสื้อยืดคอตตอนผ้านุ่มพิเศษ",
        "default_pdesc": "เสื้อยืดคุณภาพดี ใส่สบาย",
        "default_v1_name": "สี",
        "default_v1_opts": "แดง, ดำ",
        "default_v2_name": "ไซส์",
        "default_v2_opts": "S, M",
    },
    "EN": {
        "title": "📦 Shopee Mass Upload Generator & Price Calculator",
        "calc_setting": "⚙️ Price Calculation Settings",
        "currency_select": "Target Market",
        "rate_label": "Exchange Rate (JPY)",
        "add_product": "➕ Add New Product",
        "del_product": "🗑️ Delete Product",
        "product_num": "🛒 Product #",
        "cat_id": "Category ID",
        "parent_sku": "Parent SKU",
        "brand": "Brand",
        "p_name": "Product Name",
        "weight": "Weight (kg)",
        "p_desc": "Product Description",
        "cover_img": "Cover Image URL",
        "v1_name": "Variation 1 Name (e.g., Color)",
        "v1_opts": "Variation 1 Options (comma separated)",
        "v1_imgs": "Variation 1 Image URLs (comma separated)",
        "v2_name": "Variation 2 Name (e.g., Size) [Optional]",
        "v2_opts": "Variation 2 Options (comma separated)",
        "grid_title": "💰 Buying Price (JPY), Stock & Auto Calculated Selling Price:",
        "cost_col": "Buying Price (JPY)",
        "price_col": "Auto Selling Price",
        "stock_col": "Stock",
        "btn_generate": "🚀 Generate Combined Shopee Excel File",
        "success_msg": "✅ Successfully generated! Total {count} product(s).",
        "btn_download": "📥 Download Excel File for Shopee",
        "default_pname": "Premium Cotton T-Shirt",
        "default_pdesc": "High quality soft t-shirt, comfortable to wear.",
        "default_v1_name": "Color",
        "default_v1_opts": "Red, Black",
        "default_v2_name": "Size",
        "default_v2_opts": "S, M",
    },
    "JA": {
        "title": "📦 Shopee 一括出品ファイル生成 & 自動価格計算ツール",
        "calc_setting": "⚙️ 価格計算設定 (ターゲット市場 & 為替)",
        "currency_select": "ターゲット市場",
        "rate_label": "為替レート (JPY)",
        "add_product": "➕ 新しい商品を追加",
        "del_product": "🗑️ この商品を削除",
        "product_num": "🛒 商品 #",
        "cat_id": "カテゴリーID",
        "parent_sku": "親SKU (Parent SKU)",
        "brand": "ブランド (Brand)",
        "p_name": "商品名",
        "weight": "重量 (kg)",
        "p_desc": "商品説明",
        "cover_img": "メインカバー画像URL",
        "v1_name": "バリエーション1名称 (例: 色)",
        "v1_opts": "バリエーション1の選択肢 (カンマ区切り)",
        "v1_imgs": "バリエーション1の画像URL (カンマ区切り)",
        "v2_name": "バリエーション2名称 (例: サイズ) [任意]",
        "v2_opts": "バリエーション2の選択肢 (カンマ区切り)",
        "grid_title": "💰 仕入れ値(JPY)・在庫・自動計算販売価格:",
        "cost_col": "仕入れ値 (JPY)",
        "price_col": "自動計算販売価格",
        "stock_col": "在庫数",
        "btn_generate": "🚀 全商品まとめてShopee用Excelファイルを生成",
        "success_msg": "✅ 生成成功！ 合計 {count} 件の商品。",
        "btn_download": "📥 Shopeeアップロード用Excelをダウンロード",
        "default_pname": "プレミアムコットンTシャツ",
        "default_pdesc": "高品質で着心地の良いTシャツです。",
        "default_v1_name": "カラー",
        "default_v1_opts": "レッド, ブラック",
        "default_v2_name": "サイズ",
        "default_v2_opts": "S, M",
    }
}

# Sidebar ภาษา
st.sidebar.title("🌐 Language / 言語")
selected_lang = st.sidebar.radio("Select Language", ["TH (ไทย)", "EN (English)", "JA (日本語)"], index=0)

if "EN" in selected_lang:
    lang_code = "EN"
elif "JA" in selected_lang:
    lang_code = "JA"
else:
    lang_code = "TH"

T = LANG_TEXTS[lang_code]

st.title(T["title"])

# --- 3. GLOBAL CONTROL PANEL ---
st.subheader(T["calc_setting"])
col_cur, col_rate = st.columns(2)

with col_cur:
    currency = st.selectbox(T["currency_select"], ["THB", "PHP"])

with col_rate:
    default_rate = 4.868196 if currency == "THB" else 2.590111
    rate_jpy = st.number_input(f"{T['rate_label']} ({currency}/JPY)", value=default_rate, format="%.6f")

st.markdown("---")

if "products" not in st.session_state:
    st.session_state.products = [
        {
            "category_id": "120039",
            "parent_sku": "SHIRT-001",
            "brand": "No Brand",
            "product_name": T["default_pname"],
            "weight": 0.1,  # ค่าเริ่มต้น 100g (0.1 kg)
            "product_desc": T["default_pdesc"],
            "cover_image": "https://example.com/shirt_cover.jpg",
            "v1_name": T["default_v1_name"],
            "v1_options": T["default_v1_opts"],
            "v1_images": "https://example.com/red.jpg, https://example.com/black.jpg",
            "v2_name": T["default_v2_name"],
            "v2_options": T["default_v2_opts"],
        }
    ]

# ปุ่มเพิ่มรายการสินค้า
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button(T["add_product"]):
        st.session_state.products.append({
            "category_id": "100000",
            "parent_sku": f"ITEM-{len(st.session_state.products)+1:03d}",
            "brand": "No Brand",
            "product_name": f"{T['product_num']} {len(st.session_state.products)+1}",
            "weight": 0.1,
            "product_desc": "...",
            "cover_image": "https://example.com/cover.jpg",
            "v1_name": "Option",
            "v1_options": "A, B",
            "v1_images": "",
            "v2_name": "",
            "v2_options": "",
        })
        st.rerun()

updated_products_data = []

for idx, p in enumerate(st.session_state.products):
    st.markdown("---")
    col_title, col_del = st.columns([8, 2])
    with col_title:
        st.subheader(f"{T['product_num']}{idx + 1}: {p['product_name']}")
    with col_del:
        if len(st.session_state.products) > 1:
            if st.button(f"{T['del_product']}", key=f"del_{idx}"):
                st.session_state.products.pop(idx)
                st.rerun()

    # 1. ข้อมูลหลัก
    c1, c2, c3 = st.columns(3)
    with c1:
        cat_id = st.text_input(T["cat_id"], value=p["category_id"], key=f"cat_{idx}")
        p_sku = st.text_input(T["parent_sku"], value=p["parent_sku"], key=f"psku_{idx}")
        brand = st.text_input(T["brand"], value=p["brand"], key=f"brand_{idx}")
    with c2:
        p_name = st.text_input(T["p_name"], value=p["product_name"], key=f"name_{idx}")
        weight = st.number_input(T["weight"], value=float(p["weight"]), step=0.05, format="%.2f", key=f"w_{idx}")
    with c3:
        p_desc = st.text_area(T["p_desc"], value=p["product_desc"], key=f"desc_{idx}")

    # 2. รูปภาพปก
    cover_img = st.text_input(T["cover_img"], value=p["cover_image"], key=f"cimg_{idx}")

    # 3. ตัวเลือก Variation
    cv1, cv2 = st.columns(2)
    with cv1:
        v1_name = st.text_input(T["v1_name"], value=p["v1_name"], key=f"v1n_{idx}")
        v1_opts = st.text_input(T["v1_opts"], value=p["v1_options"], key=f"v1o_{idx}")
        v1_imgs = st.text_input(T["v1_imgs"], value=p["v1_images"], key=f"v1i_{idx}")
    with cv2:
        v2_name = st.text_input(T["v2_name"], value=p["v2_name"], key=f"v2n_{idx}")
        v2_opts = st.text_input(T["v2_opts"], value=p["v2_options"], key=f"v2o_{idx}")

    # คำนวณตาราง Variation
    list_v1 = [x.strip() for x in v1_opts.split(",") if x.strip()]
    list_v2 = [x.strip() for x in v2_opts.split(",") if x.strip()] if v2_name else [""]
    variations = list(itertools.product(list_v1, list_v2))

    grid_data = []
    price_col_label = f"{T['price_col']} ({currency})"

    for opt1, opt2 in variations:
        var_title = f"{opt1}" + (f" / {opt2}" if opt2 else "")
        sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")
        default_cost_jpy = 1000  # ค่าเริ่มต้น 1000 Yen ตามต้องการ

        grid_data.append({
            "Variation": var_title,
            "SKU": f"{p_sku}{sku_suffix}",
            T["cost_col"]: default_cost_jpy,
            price_col_label: 0,
            T["stock_col"]: 50,
            "Opt1": opt1,
            "Opt2": opt2
        })
        
    df_var = pd.DataFrame(grid_data)

    # คำนวณราคาอัตโนมัติ Real-time จากต้นทุน JPY และ น้ำหนัก kg
    df_var[price_col_label] = df_var.apply(
        lambda row: calculate_net_price(
            buying_price_jpy=row[T["cost_col"]],
            weight_kg=weight,
            currency=currency,
            rate_jpy=rate_jpy
        ), axis=1
    )

    st.write(T["grid_title"])
    edited_df = st.data_editor(
        df_var,
        column_config={
            "Variation": st.column_config.Column(disabled=True),
            T["cost_col"]: st.column_config.NumberColumn(T["cost_col"], min_value=0, format="%d ¥"),
            price_col_label: st.column_config.NumberColumn(price_col_label, disabled=True, format="%d " + currency),
            T["stock_col"]: st.column_config.NumberColumn(T["stock_col"], min_value=0, format="%d"),
            "Opt1": None,
            "Opt2": None
        },
        hide_index=True,
        key=f"editor_{idx}"
    )

    # Recalculate หลังการแก้ไขราคาซื้อ JPY ในตาราง
    edited_df[price_col_label] = edited_df.apply(
        lambda row: calculate_net_price(
            buying_price_jpy=row[T["cost_col"]],
            weight_kg=weight,
            currency=currency,
            rate_jpy=rate_jpy
        ), axis=1
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
        "variations_table": edited_df,
        "price_col_label": price_col_label
    })

# --- 4. EXPORT EXCEL ---
st.markdown("---")
if st.button(T["btn_generate"], type="primary", use_container_width=True):
    try:
        wb = openpyxl.load_workbook("Shopee_template.xlsx")
        ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
    except Exception:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template"

    start_row = 6

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
        price_col_label = p_data["price_col_label"]

        unique_v1 = df_vars["Opt1"].unique().tolist()
        v1_img_dict = {}
        for i, opt in enumerate(unique_v1):
            v1_img_dict[opt] = v1_imgs[i] if i < len(v1_imgs) else ""

        for idx, row in df_vars.iterrows():
            current_row = start_row
            
            ws.cell(row=current_row, column=1, value=cat_id)
            ws.cell(row=current_row, column=2, value=p_name if idx == 0 else "")
            ws.cell(row=current_row, column=3, value=p_desc if idx == 0 else "")
            ws.cell(row=current_row, column=10, value=p_sku)
            
            ws.cell(row=current_row, column=11, value=v1_name)
            ws.cell(row=current_row, column=12, value=row["Opt1"])
            ws.cell(row=current_row, column=13, value=v1_img_dict.get(row["Opt1"], ""))
            
            if v2_name and row["Opt2"]:
                ws.cell(row=current_row, column=14, value=v2_name)
                ws.cell(row=current_row, column=15, value=row["Opt2"])
                
            ws.cell(row=current_row, column=16, value=row[price_col_label])
            ws.cell(row=current_row, column=17, value=row[T["stock_col"]])
            ws.cell(row=current_row, column=18, value=row["SKU"])
            
            if idx == 0:
                ws.cell(row=current_row, column=21, value=cover_img)
                ws.cell(row=current_row, column=30, value=weight)
                ws.cell(row=current_row, column=34, value="On")

            start_row += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(T["success_msg"].format(count=len(updated_products_data)))
    st.download_button(
        label=T["btn_download"],
        data=output,
        file_name=f"Shopee_Mass_Upload_{currency}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )