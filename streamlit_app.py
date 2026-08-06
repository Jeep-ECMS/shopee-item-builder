import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
import itertools
import pandas as pd
import urllib.request
import json

st.set_page_config(page_title="Shopee Auto Price & Mass Upload Builder", layout="wide")

# --- 0. API ดึงอัตราแลกเปลี่ยน REALTIME (THB -> JPY และ PHP -> JPY) ---
@st.cache_data(ttl=3600)
def fetch_base_rates():
    default_rates = {"THB": 4.73, "PHP": 2.65}
    rates_out = {}
    
    try:
        url_thb = "https://open.er-api.com/v6/latest/THB"
        req = urllib.request.Request(url_thb, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            thb_jpy = data.get("rates", {}).get("JPY")
            if thb_jpy:
                rates_out["THB"] = round(thb_jpy, 4)
    except Exception:
        rates_out["THB"] = default_rates["THB"]

    try:
        url_php = "https://open.er-api.com/v6/latest/PHP"
        req = urllib.request.Request(url_php, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            php_jpy = data.get("rates", {}).get("JPY")
            if php_jpy:
                rates_out["PHP"] = round(php_jpy, 4)
    except Exception:
        rates_out["PHP"] = default_rates["PHP"]

    return rates_out


# --- 1. ตาราง SLS TRANSPORTATION & คำนวณราคาปลอดภัย ---
SLS_RATES_THB = [
    (100, 93), (200, 105), (300, 129), (400, 153), (500, 177),
    (600, 200), (700, 230), (800, 250), (900, 270), (1000, 300),
    (1100, 321), (1500, 388), (2000, 508), (2100, 550), (2200, 585),
    (2300, 595), (2400, 615), (2500, 628), (3000, 748), (3500, 868),
    (4000, 988), (4500, 1108), (5000, 1228), (5500, 1348), (6000, 1468),
    (6500, 1588), (7000, 1708), (7500, 1828), (8000, 1948), (8500, 2068),
    (9000, 2188), (9500, 2308), (10000, 2428), (10500, 2548), (11000, 2668),
    (11500, 2788), (12000, 2908), (12500, 3028), (13000, 3148), (13500, 3268),
    (14000, 3388), (14500, 3508), (15000, 3628), (15500, 3748), (16000, 3868),
    (16500, 3988), (17000, 4108), (17500, 4228), (18000, 4348), (18500, 4468),
    (19000, 4588), (19500, 4708), (20000, 4828)
]

SLS_RATES_PHP = [
    (50, 76), (100, 101), (150, 126), (200, 151), (250, 176),
    (300, 201), (350, 226), (400, 251), (450, 276), (500, 301),
    (550, 326), (600, 351), (650, 376), (700, 401), (750, 426),
    (800, 451), (850, 476), (900, 501), (950, 526), (1000, 551),
    (1050, 576), (1100, 601), (1150, 626), (1200, 651), (1250, 676),
    (1300, 701), (1350, 726), (1400, 751), (1450, 776), (1500, 801),
    (1550, 826)
]

def get_sls_shipping_fee_thb(weight_g):
    for max_weight, fee in SLS_RATES_THB:
        if weight_g <= max_weight:
            return fee
    last_weight, last_fee = SLS_RATES_THB[-1]
    extra_steps = ((weight_g - last_weight) + 499) // 500
    return last_fee + (extra_steps * 120)

def get_sls_shipping_fee_php(weight_g):
    for max_weight, fee in SLS_RATES_PHP:
        if weight_g <= max_weight:
            return fee
    last_weight, last_fee = SLS_RATES_PHP[-1]
    extra_steps = ((weight_g - last_weight) + 49) // 50
    return last_fee + (extra_steps * 25)

def calculate_net_price(buying_price_jpy, weight_g, profit_rate_pct=30.0, currency="THB", rate_to_jpy=None):
    try:
        buying_price_jpy = float(buying_price_jpy) if buying_price_jpy is not None else 0.0
        weight_g = float(weight_g) if weight_g is not None else 0.0
        profit_rate_pct = float(profit_rate_pct) if profit_rate_pct is not None else 0.0
    except (ValueError, TypeError):
        return 0

    if buying_price_jpy <= 0 or weight_g <= 0:
        return 0

    margin_factor = 1.0 - (profit_rate_pct / 100.0)
    if margin_factor <= 0:
        margin_factor = 0.01

    if currency == "THB":
        rate = rate_to_jpy if rate_to_jpy else 4.73
        buying_price_thb = buying_price_jpy / rate if rate > 0 else 0
        sls_fee = get_sls_shipping_fee_thb(weight_g)
        transportation_jp_thb = 70.0
        net_price = (sls_fee + buying_price_thb + transportation_jp_thb) / margin_factor
        return round(net_price)

    elif currency == "PHP":
        rate = rate_to_jpy if rate_to_jpy else 2.65
        buying_price_php = buying_price_jpy / rate if rate > 0 else 0
        transportation_jp_php = 116.0
        sls_fee = get_sls_shipping_fee_php(weight_g)
        
        base_price = (sls_fee + buying_price_php + transportation_jp_php) / margin_factor
        cif_check = (base_price * 1.01) + (1369.0 * (weight_g / 1000.0))
        
        if cif_check >= 10000:
            custom_tax_part = 1369.0 * (weight_g / 1000.0) * 0.12
            net_price = (sls_fee + custom_tax_part + buying_price_php + transportation_jp_php) / 0.5788
        else:
            net_price = base_price
            
        return round(net_price)

    return 0


# --- 2. MULTI-LANGUAGE DICTIONARY ---
LANG_TEXTS = {
    "TH": {
        "title": "📦 เครื่องมือสร้างไฟล์ Mass Upload Shopee & คำนวณราคาขายอัตโนมัติ",
        "calc_setting": "⚙️ ตั้งค่าการคำนวณราคา (Target Market & Real-time Exchange Rate)",
        "currency_select": "เลือกตลาดเป้าหมาย",
        "rate_label": "อัตราแลกเปลี่ยน Real-time (1 {curr} -> JPY)",
        "rate_info": "💡 ดึงข้อมูลอัตราแลกเปลี่ยน Real-time ล่าสุดอัตโนมัติ",
        "add_product": "➕ เพิ่มสินค้าชิ้นใหม่",
        "del_product": "🗑️ ลบสินค้านี้",
        "product_num": "🛒 สินค้าชิ้นที่",
        "cat_id": "Category ID / รหัสหมวดหมู่",
        "parent_sku": "Parent SKU / รหัสอ้างอิงหลัก",
        "integration_no": "Variation Integration No.",
        "brand": "แบรนด์ (Brand)",
        "p_name": "ชื่อสินค้า",
        "weight": "น้ำหนักสินค้าเริ่มต้น (g)",
        "p_desc": "รายละเอียดสินค้า",
        "cover_img": "URL รูปภาพปกหลัก & รูปสินค้า (คั่นด้วย , จะกระจายไป Column U-AC อัตโนมัติ)",
        "cover_img_help": "ใส่ URL คั่นด้วยเครื่องหมายจุลภาค รูปแรก = Cover Image, รูปถัดไป = Item Image 1-8",
        "v1_name": "ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)",
        "v1_opts": "รายการตัวเลือกที่ 1 (คั่นด้วย ,)",
        "v1_imgs": "URL รูปภาพตัวเลือกที่ 1 (1 รูปต่อ 1 ตัวเลือก คั่นด้วย ,)",
        "v1_imgs_help": "ระบุ URL รูปตามลำดับตัวเลือกที่ 1",
        "v2_name": "ชื่อตัวเลือกที่ 2 (เช่น ไซส์) [เว้นว่างได้]",
        "v2_opts": "รายการตัวเลือกที่ 2 (คั่นด้วย ,)",
        "batch_title": "⚡ ตั้งค่าด่วน (แยกปรับแต่ละค่าไปยังทุก Variation):",
        "btn_apply": "⚡ นำไปใช้",
        "grid_title": "💰 ตารางกำหนดราคาซื้อ (JPY), น้ำหนัก (g), Profit Rate (%), สต๊อก และราคาขาย:",
        "cost_col": "Buying Price (JPY)",
        "weight_col": "Weight (g)",
        "profit_col": "Profit Rate (%)",
        "price_col": "Selling Price",
        "stock_col": "Stock (ชิ้น)",
        "sku_col": "SKU (แก้ไขได้)",
        "btn_generate": "🚀 สร้างไฟล์ Excel รวมทุกสินค้าสำหรับ Shopee",
        "success_msg": "✅ สร้างไฟล์สำเร็จ! รวมสินค้าทั้งหมด {count} รายการ",
        "btn_download": "📥 ดาวน์โหลดไฟล์ Excel พร้อมอัปโหลด Shopee",
        "default_pname": "รองเท้าสปอร์ตผ้าใบคุณภาพสูง",
        "default_pdesc": "รองเท้าสปอร์ต นุ่ม สวมใส่สบาย",
        "default_v1_name": "สี",
        "default_v1_opts": "WINE, WHITE",
        "default_v2_name": "ไซส์",
        "default_v2_opts": "23.0cm., 24.0cm., 25.0cm., 26.0cm.",
    },
    "EN": {
        "title": "📦 Shopee Mass Upload Generator & Price Calculator",
        "calc_setting": "⚙️ Price Calculation Settings (Target Market & Real-time Rate)",
        "currency_select": "Target Market",
        "rate_label": "Real-time Exchange Rate (1 {curr} -> JPY)",
        "rate_info": "💡 Auto-fetched latest real-time exchange rates",
        "add_product": "➕ Add New Product",
        "del_product": "🗑️ Delete Product",
        "product_num": "🛒 Product #",
        "cat_id": "Category ID",
        "parent_sku": "Parent SKU",
        "integration_no": "Variation Integration No.",
        "brand": "Brand",
        "p_name": "Product Name",
        "weight": "Default Weight (g)",
        "p_desc": "Product Description",
        "cover_img": "Cover Image & Item Images URLs (comma-separated, auto-splits to U-AC)",
        "cover_img_help": "First URL = Cover Image, subsequent = Item Image 1-8",
        "v1_name": "Variation 1 Name (e.g., Color)",
        "v1_opts": "Variation 1 Options (comma separated)",
        "v1_imgs": "Variation 1 Image URLs (1 per option, comma separated)",
        "v1_imgs_help": "Provide 1 image URL per Variation 1 option.",
        "v2_name": "Variation 2 Name (e.g., Size) [Optional]",
        "v2_opts": "Variation 2 Options (comma separated)",
        "batch_title": "⚡ Quick Setup (Apply individual value to all variations):",
        "btn_apply": "⚡ Apply",
        "grid_title": "💰 Buying Price (JPY), Weight (g), Profit Rate (%), Stock & Auto Price Table:",
        "cost_col": "Buying Price (JPY)",
        "weight_col": "Weight (g)",
        "profit_col": "Profit Rate (%)",
        "price_col": "Selling Price",
        "stock_col": "Stock",
        "sku_col": "SKU (Editable)",
        "btn_generate": "🚀 Generate Combined Shopee Excel File",
        "success_msg": "✅ Successfully generated! Total {count} product(s).",
        "btn_download": "📥 Download Excel File for Shopee",
        "default_pname": "High Quality Sneakers",
        "default_pdesc": "Soft and comfortable sport shoes.",
        "default_v1_name": "Color",
        "default_v1_opts": "WINE, WHITE",
        "default_v2_name": "Size",
        "default_v2_opts": "23.0cm., 24.0cm., 25.0cm., 26.0cm.",
    },
    "JA": {
        "title": "📦 Shopee 一括出品ファイル生成 & 自動価格計算ツール",
        "calc_setting": "⚙️ 価格計算設定 (ターゲット市場 & リアルタイム為替レート)",
        "currency_select": "ターゲット市場",
        "rate_label": "リアルタイム為替レート (1 {curr} -> JPY)",
        "rate_info": "💡 最新のリアルタイム為替レートを自動取得中",
        "add_product": "➕ 新しい商品を追加",
        "del_product": "🗑️ この商品を削除",
        "product_num": "🛒 商品 #",
        "cat_id": "カテゴリーID",
        "parent_sku": "親SKU (Parent SKU)",
        "integration_no": "Variation Integration No.",
        "brand": "ブランド (Brand)",
        "p_name": "商品名",
        "weight": "デフォルト重量 (g)",
        "p_desc": "商品説明",
        "cover_img": "カバー画像URL (カンマ区切りでU-AC列に自動展開)",
        "cover_img_help": "1つ目＝カバー画像、2つ目以降＝商品画像1〜8",
        "v1_name": "バリエーション1名称 (例: 色)",
        "v1_opts": "バリエーション1の選択肢 (カンマ区切り)",
        "v1_imgs": "バリエーション1の画像URL (各選択肢1枚、カンマ区切り)",
        "v1_imgs_help": "バリエーション1の各項目ごとに1枚のURLを入力してください。",
        "v2_name": "バリエーション2名称 (例: サイズ) [任意]",
        "v2_opts": "バリエーション2の選択肢 (カンマ区切り)",
        "batch_title": "⚡ 個別一括設定 (全バリエーションに適用):",
        "btn_apply": "⚡ 適用",
        "grid_title": "💰 仕入れ値(JPY)・重量(g)・利益率(%)・在庫・自動計算販売価格:",
        "cost_col": "Buying Price (JPY)",
        "weight_col": "Weight (g)",
        "profit_col": "Profit Rate (%)",
        "price_col": "Selling Price",
        "stock_col": "在庫数",
        "sku_col": "SKU (編集可能)",
        "btn_generate": "🚀 全商品まとめてShopee用Excelファイルを生成",
        "success_msg": "✅ 生成成功！ 合計 {count} 件の商品。",
        "btn_download": "📥 Shopeeアップロード用Excelをダウンロード",
        "default_pname": "高品質スニーカー",
        "default_pdesc": "履き心地が良く快適なスポーツシューズ。",
        "default_v1_name": "カラー",
        "default_v1_opts": "WINE, WHITE",
        "default_v2_name": "サイズ",
        "default_v2_opts": "23.0cm., 24.0cm., 25.0cm., 26.0cm.",
    }
}

st.sidebar.title("🌐 Language Settings")
selected_lang = st.sidebar.selectbox(
    "Choose Language / 言語選択",
    ["TH (ไทย)", "EN (English)", "JA (日本語)"],
    index=0
)

if "EN" in selected_lang:
    lang_code = "EN"
elif "JA" in selected_lang:
    lang_code = "JA"
else:
    lang_code = "TH"

T = LANG_TEXTS[lang_code]
st.title(T["title"])

# --- 3. GLOBAL CONTROL PANEL (REAL-TIME EXCHANGE RATE) ---
realtime_rates = fetch_base_rates()

st.subheader(T["calc_setting"])
col_cur, col_rate = st.columns(2)

with col_cur:
    currency = st.selectbox(T["currency_select"], ["THB", "PHP"], key="global_currency")

with col_rate:
    current_realtime_rate = realtime_rates.get(currency, 4.73 if currency == "THB" else 2.65)
    rate_jpy = st.number_input(
        T["rate_label"].format(curr=currency), 
        value=current_realtime_rate, 
        format="%.4f",
        help=T["rate_info"],
        key="global_rate"
    )
    st.caption(f"{T['rate_info']}: **1 {currency} = {rate_jpy} JPY**")

st.markdown("---")

# --- INITIALIZE SESSION STATE FOR PRODUCTS ---
if "products_list" not in st.session_state:
    st.session_state.products_list = [0]
    st.session_state.next_prod_id = 1

col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button(T["add_product"]):
        new_id = st.session_state.next_prod_id
        st.session_state.next_prod_id += 1
        st.session_state.products_list.append(new_id)
        st.rerun()

updated_products_data = []
prod_id_to_remove = None

for idx, p_id in enumerate(st.session_state.products_list):
    st.markdown("---")
    col_title, col_del = st.columns([8, 2])
    
    # กำหนดค่าเริ่มต้นของสินค้า หากยังไม่มีใน Session State
    if f"name_{p_id}" not in st.session_state:
        st.session_state[f"cat_{p_id}"] = "120039"
        st.session_state[f"psku_{p_id}"] = f"361086-{p_id+18}"
        st.session_state[f"integ_{p_id}"] = ""
        st.session_state[f"brand_{p_id}"] = "No Brand"
        st.session_state[f"name_{p_id}"] = T["default_pname"] if p_id == 0 else f"{T['product_num']}{idx+1}"
        st.session_state[f"w_{p_id}"] = 300.0
        st.session_state[f"desc_{p_id}"] = T["default_pdesc"]
        st.session_state[f"cimg_{p_id}"] = "https://example.com/cover.jpg, https://example.com/img1.jpg, https://example.com/img2.jpg"
        st.session_state[f"v1n_{p_id}"] = T["default_v1_name"]
        st.session_state[f"v1o_{p_id}"] = T["default_v1_opts"]
        st.session_state[f"v1i_{p_id}"] = "https://example.com/wine.jpg, https://example.com/white.jpg"
        st.session_state[f"v2n_{p_id}"] = T["default_v2_name"]
        st.session_state[f"v2o_{p_id}"] = T["default_v2_opts"]

    with col_title:
        st.subheader(f"{T['product_num']}{idx + 1}: {st.session_state.get(f'name_{p_id}', '')}")
    with col_del:
        if len(st.session_state.products_list) > 1:
            if st.button(f"{T['del_product']}", key=f"del_btn_{p_id}"):
                prod_id_to_remove = p_id

    c1, c2, c3 = st.columns(3)
    with c1:
        cat_id_val = st.text_input(T["cat_id"], key=f"cat_{p_id}")
        col_psku, col_integ = st.columns(2)
        with col_psku:
            psku_val = st.text_input(T["parent_sku"], key=f"psku_{p_id}")
        with col_integ:
            integ_val = st.text_input(T["integration_no"], key=f"integ_{p_id}")
        brand_val = st.text_input(T["brand"], key=f"brand_{p_id}")
    with c2:
        pname_val = st.text_input(T["p_name"], key=f"name_{p_id}")
        weight_val = st.number_input(T["weight"], step=10.0, format="%.1f", key=f"w_{p_id}")
    with c3:
        pdesc_val = st.text_area(T["p_desc"], key=f"desc_{p_id}")

    cimg_val = st.text_input(T["cover_img"], help=T["cover_img_help"], key=f"cimg_{p_id}")

    cv1, cv2 = st.columns(2)
    with cv1:
        v1n_val = st.text_input(T["v1_name"], key=f"v1n_{p_id}")
        v1o_val = st.text_input(T["v1_opts"], key=f"v1o_{p_id}")
        v1i_val = st.text_input(T["v1_imgs"], help=T["v1_imgs_help"], key=f"v1i_{p_id}")
    with cv2:
        v2n_val = st.text_input(T["v2_name"], key=f"v2n_{p_id}")
        v2o_val = st.text_input(T["v2_opts"], key=f"v2o_{p_id}")

    # ป้องกันการเกิดค่าว่างเปล่า (Empty Options)
    list_v1 = [x.strip() for x in v1o_val.split(",") if x.strip()]
    if not list_v1:
        list_v1 = ["Standard"]

    list_v2 = [x.strip() for x in v2o_val.split(",") if x.strip()] if v2n_val else [""]
    if not list_v2:
        list_v2 = [""]

    variations = list(itertools.product(list_v1, list_v2))

    st.write(T["batch_title"])
    b_col1, b_col2, b_col3, b_col4 = st.columns(4)
    
    with b_col1:
        batch_cost = st.number_input(T["cost_col"], value=1590, step=100, key=f"b_cost_{p_id}")
        apply_cost = st.button(f"{T['btn_apply']} {T['cost_col']}", key=f"btn_apply_cost_{p_id}")
        
    with b_col2:
        batch_weight = st.number_input(T["weight_col"], value=float(weight_val if weight_val else 300.0), step=10.0, format="%.1f", key=f"b_weight_{p_id}")
        apply_weight = st.button(f"{T['btn_apply']} {T['weight_col']}", key=f"btn_apply_weight_{p_id}")
        
    with b_col3:
        batch_profit = st.number_input(T["profit_col"], value=30.0, step=1.0, format="%.1f", key=f"b_profit_{p_id}")
        apply_profit = st.button(f"{T['btn_apply']} {T['profit_col']}", key=f"btn_apply_profit_{p_id}")
        
    with b_col4:
        batch_stock = st.number_input(T["stock_col"], value=2, step=1, key=f"b_stock_{p_id}")
        apply_stock = st.button(f"{T['btn_apply']} {T['stock_col']}", key=f"btn_apply_stock_{p_id}")

    df_state_key = f"df_data_{p_id}"
    
    cost_key = "cost_jpy"
    weight_key = "weight_g"
    profit_key = "profit_rate"
    price_key = "selling_price"
    stock_key = "stock"

    if df_state_key not in st.session_state:
        grid_data = []
        for opt1, opt2 in variations:
            var_title = f"{opt1}" + (f" / {opt2}" if opt2 else "")
            sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")

            grid_data.append({
                "Variation": var_title,
                "SKU": f"{psku_val}{sku_suffix}",
                cost_key: int(batch_cost),
                weight_key: float(batch_weight),
                profit_key: float(batch_profit),
                price_key: 0,
                stock_key: int(batch_stock),
                "Opt1": opt1,
                "Opt2": opt2
            })
        st.session_state[df_state_key] = pd.DataFrame(grid_data)
    else:
        df_existing = st.session_state[df_state_key]
        new_grid_data = []
        for opt1, opt2 in variations:
            var_title = f"{opt1}" + (f" / {opt2}" if opt2 else "")
            sku_default = f"{psku_val}" + (f"-{opt1}" if opt1 else "") + (f"-{opt2}" if opt2 else "")
            
            match = df_existing[df_existing["Variation"] == var_title] if not df_existing.empty and "Variation" in df_existing.columns else pd.DataFrame()
            
            if not match.empty:
                c_val = int(batch_cost) if apply_cost else match.iloc[0].get(cost_key, 1590)
                w_val = float(batch_weight) if apply_weight else match.iloc[0].get(weight_key, float(weight_val if weight_val else 300.0))
                p_val = float(batch_profit) if apply_profit else match.iloc[0].get(profit_key, 30.0)
                s_val = int(batch_stock) if apply_stock else match.iloc[0].get(stock_key, 2)
                sku_val = match.iloc[0].get("SKU", sku_default)
            else:
                c_val = int(batch_cost) if apply_cost else 1590
                w_val = float(batch_weight) if apply_weight else float(weight_val if weight_val else 300.0)
                p_val = float(batch_profit) if apply_profit else 30.0
                s_val = int(batch_stock) if apply_stock else 2
                sku_val = sku_default

            new_grid_data.append({
                "Variation": var_title,
                "SKU": sku_val,
                cost_key: c_val,
                weight_key: w_val,
                profit_key: p_val,
                price_key: 0,
                stock_key: s_val,
                "Opt1": opt1,
                "Opt2": opt2
            })
        st.session_state[df_state_key] = pd.DataFrame(new_grid_data)

    df_var = st.session_state[df_state_key]

    # --- 🛡️ ตรวจสอบความปลอดภัยแบบร้อยเปอร์เซ็นต์ ไม่ให้เกิด Error บน DataFrame ---
    if not df_var.empty:
        df_var[price_key] = df_var.apply(
            lambda row: calculate_net_price(
                buying_price_jpy=row.get(cost_key, 0),
                weight_g=row.get(weight_key, 0),
                profit_rate_pct=row.get(profit_key, 0),
                currency=currency,
                rate_to_jpy=rate_jpy
            ), axis=1
        )
    else:
        df_var[price_key] = []

    st.write(T["grid_title"])
    edited_df = st.data_editor(
        df_var,
        column_config={
            "Variation": st.column_config.Column(disabled=True),
            "SKU": st.column_config.TextColumn(T["sku_col"], disabled=False),
            cost_key: st.column_config.NumberColumn(T["cost_col"], min_value=0, format="%d ¥"),
            weight_key: st.column_config.NumberColumn(T["weight_col"], min_value=0.0, format="%.1f g"),
            profit_key: st.column_config.NumberColumn(T["profit_col"], min_value=0.0, max_value=99.0, format="%.1f %%"),
            price_key: st.column_config.NumberColumn(f"{T['price_col']} ({currency})", disabled=True, format="%d " + currency),
            stock_key: st.column_config.NumberColumn(T["stock_col"], min_value=0, format="%d"),
            "Opt1": None,
            "Opt2": None
        },
        hide_index=True,
        key=f"editor_{p_id}"
    )

    st.session_state[df_state_key] = edited_df

    v1_imgs_list = [x.strip() for x in v1i_val.split(",") if x.strip()]
    updated_products_data.append({
        "cat_id": cat_id_val,
        "p_sku": psku_val,
        "integration_no": integ_val,
        "brand": brand_val,
        "p_name": pname_val,
        "weight": weight_val if weight_val else 0,
        "p_desc": pdesc_val,
        "cover_img": cimg_val,
        "v1_name": v1n_val,
        "v1_opts_list": list_v1,
        "v1_imgs": v1_imgs_list,
        "v2_name": v2n_val,
        "variations_table": edited_df,
        "cost_key": cost_key,
        "weight_key": weight_key,
        "price_key": price_key,
        "stock_key": stock_key
    })

if prod_id_to_remove is not None:
    st.session_state.products_list.remove(prod_id_to_remove)
    if f"df_data_{prod_id_to_remove}" in st.session_state:
        del st.session_state[f"df_data_{prod_id_to_remove}"]
    st.rerun()

# --- 4. EXPORT EXCEL ---
st.markdown("---")
if st.button(T["btn_generate"], type="primary", use_container_width=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"

    header_row1 = [
        "Category", "Product Name", "Product Description", "Maximum Purchase Quantity", 
        "Maximum Purchase Quantity - Start Date", "Maximum Purchase Quantity - Time Period (in Days)", 
        "Maximum Purchase Quantity - End Date", "Minimum Purchase Quantity", "Parent SKU", 
        "Variation Integration No.", "Variation Name1", "Option for Variation 1", "Image per Variation", 
        "Variation Name2", "Option for Variation 2", "Price", "Stock", "SKU", "Size Chart Template", 
        "Size Chart Image", "Cover image", "Item Image 1", "Item Image 2", "Item Image 3", 
        "Item Image 4", "Item Image 5", "Item Image 6", "Item Image 7", "Item Image 8", "Weight", 
        "Length", "Width", "Height", "International Express - ส่งจากต่างประเทศ (Japan)", 
        "Pre-order DTS", "Fail Reason"
    ]

    header_row2 = [
        "Optional", "Mandatory", "Mandatory", "Optional", "Conditional Mandatory", 
        "Conditional Mandatory", "Conditional Mandatory", "Optional", "Optional", 
        "Conditional Mandatory", "Conditional Mandatory", "Conditional Mandatory", "Conditional Mandatory", 
        "Conditional Mandatory", "Conditional Mandatory", "Mandatory", "Conditional Mandatory", 
        "Optional", "Conditional Mandatory", "Conditional Mandatory", "Mandatory", "Optional", 
        "Optional", "Optional", "Optional", "Optional", "Optional", "Optional", "Optional", 
        "Conditional Mandatory", "Conditional Mandatory", "Conditional Mandatory", "Conditional Mandatory", 
        "Optional", "Conditional Mandatory", "Optional"
    ]

    header_row3 = [
        "Indicate the appropriate category ID for each product.", "Product name should include product brand and model.", 
        "A good product description enhances the quality of your listing.", "[Per Order + Per Time Period]", 
        "Please select a MaxPQ start date.", "[Per Time Period only]", "Please select a MaxPQ end date", 
        "MPQ is an item level field.", "Parent SKU is used to identify parent products.", 
        "Mandatory for products with variations.", "Please indicate the first variation name.", 
        "Indicate the first variation value.", "Upload an image per variation.", "Indicate the second variation name.", 
        "Please indicate the second variation value.", "Input your product price.", "Input your product stock.", 
        "SKU is a unique identifier.", "Please enter the size chart template ID.", "You only need to fill in either size chart.", 
        "Upload the URL of your main product image.", "Enter the URL of this product image.", "Enter the URL of this product image.", 
        "Enter the URL of this product image.", "Enter the URL of this product image.", "Enter the URL of this product image.", 
        "Enter the URL of this product image.", "Enter the URL of this product image.", "Enter the URL of this product image.", 
        "Input your product weight.", "Fill up all dimensions.", "Input your product width.", "Input your product height.", 
        "Please toggle 'on'", "Pre-order days to ship (DTS)", ""
    ]

    header_row4 = [
        "Choose your desired category ID from the Category Tree.", "Please input 20 to 255 characters.", 
        "Please input 60 to 5000 characters.", "Please input from 1 to 999,999.", "YYYY-MM-DD", 
        "Please input from 1 to 365.", "YYYY-MM-DD", "Minimum purchase quantity only can be a positive integer.", 
        "Please input 1-100 characters.", "Input 1 to 100 characters.", "Input from 1 to 14 characters.", 
        "Input from 1 to 30 characters.", "Enter the URL of this product image.", "Input from 1 to 14 characters.", 
        "Input from 1 to 30 characters.", "Input price.", "Input stock.", "Input less than 100 characters.", 
        "Please enter size chart template ID.", "Size: Max 2Mb", "Size: max 2.0mb", "Size: max 2.0mb", 
        "Size: max 2.0mb", "Size: max 2.0mb", "Size: max 2.0mb", "Size: max 2.0mb", "Size: max 2.0mb", 
        "Size: max 2.0mb", "Size: max 2.0mb", "Please input 0.00 to 1000000.00", "Please input 0 to 1000000", 
        "Please input 0 to 1000000", "Please input 0 to 1000000", "On/Off", "Pre-order DTS range", ""
    ]

    header_row5 = [
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
        "Format accepted: PDF, JPG, JPEG, PNG", "Format accepted: JPG, JPEG, PNG.", "Format accepted: JPG, JPEG, PNG.", 
        "Format accepted: JPG, JPEG, PNG.", "Format accepted: JPG, JPEG, PNG.", "Format accepted: JPG, JPEG, PNG.", 
        "Format accepted: JPG, JPEG, PNG.", "Format accepted: JPG, JPEG, PNG.", "Format accepted: JPG, JPEG, PNG.", 
        "Format accepted: JPG, JPEG, PNG.", "", "", "", "", "", "", ""
    ]

    ws.append(header_row1)
    ws.append(header_row2)
    ws.append(header_row3)
    ws.append(header_row4)
    ws.append(header_row5)

    start_row = 6

    for p_data in updated_products_data:
        cat_id = p_data["cat_id"]
        p_sku = p_data["p_sku"]
        integration_no = p_data["integration_no"]
        p_name = p_data["p_name"]
        p_desc = p_data["p_desc"]
        cover_img = p_data["cover_img"]
        v1_name = p_data["v1_name"]
        v2_name = p_data["v2_name"]
        v1_opts_list = p_data["v1_opts_list"]
        v1_imgs = p_data["v1_imgs"]
        df_vars = p_data["variations_table"]
        
        pr_k = p_data["price_key"]
        st_k = p_data["stock_key"]
        w_k = p_data["weight_key"]

        v1_img_dict = {}
        for i, opt1_val in enumerate(v1_opts_list):
            v1_img_dict[opt1_val] = v1_imgs[i] if i < len(v1_imgs) else ""

        for idx, row in df_vars.iterrows():
            current_row = start_row
            
            ws.cell(row=current_row, column=1, value=cat_id)
            ws.cell(row=current_row, column=2, value=p_name if idx == 0 else "")
            ws.cell(row=current_row, column=3, value=p_desc if idx == 0 else "")
            
            ws.cell(row=current_row, column=9, value=p_sku)
            ws.cell(row=current_row, column=10, value=integration_no)
            
            ws.cell(row=current_row, column=11, value=v1_name)
            ws.cell(row=current_row, column=12, value=row["Opt1"])
            ws.cell(row=current_row, column=13, value=v1_img_dict.get(row["Opt1"], ""))
            
            if v2_name and row["Opt2"]:
                ws.cell(row=current_row, column=14, value=v2_name)
                ws.cell(row=current_row, column=15, value=row["Opt2"])
                
            ws.cell(row=current_row, column=16, value=row[pr_k])
            ws.cell(row=current_row, column=17, value=row[st_k])
            ws.cell(row=current_row, column=18, value=row["SKU"])
            
            if idx == 0:
                # แยก URL รูปภาพด้วยเครื่องหมาย (,) และกระจายลง Column U ถึง AC (Column 21 ถึง 29) อัตโนมัติ
                cover_imgs_list = [x.strip() for x in cover_img.split(",") if x.strip()]
                for img_i, img_url in enumerate(cover_imgs_list):
                    if img_i < 9:
                        ws.cell(row=current_row, column=21 + img_i, value=img_url)

                ws.cell(row=current_row, column=30, value=row[w_k] / 1000.0)
                ws.cell(row=current_row, column=34, value="On")

            start_row += 1

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    wrap_alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    header_alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 20

    for r in range(6, start_row):
        ws.row_dimensions[r].height = 28

    max_col = len(header_row1)
    for r in range(1, start_row):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            
            if r <= 5:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
            else:
                cell.font = data_font
                cell.alignment = wrap_alignment

    column_widths = {
        "A": 15, "B": 35, "C": 45, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15,
        "I": 25, "J": 25, "K": 20, "L": 20, "M": 30, "N": 20, "O": 20, "P": 15,
        "Q": 12, "R": 25, "S": 20, "T": 30, "U": 35, "V": 30, "W": 30, "X": 30,
        "Y": 30, "Z": 30, "AA": 30, "AB": 30, "AC": 30, "AD": 15, "AE": 12,
        "AF": 12, "AG": 12, "AH": 20, "AI": 15, "AJ": 20
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

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