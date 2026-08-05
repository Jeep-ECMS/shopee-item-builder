import streamlit as st
import openpyxl
import io
import itertools
import pandas as pd
import urllib.request
import json

st.set_page_config(page_title="Shopee Auto Price & Mass Upload Builder", layout="wide")

# --- 0. API ดึงอัตราแลกเปลี่ยน REALTIME (THB -> JPY และ PHP -> JPY) ---
@st.cache_data(ttl=3600)
def fetch_base_rates():
    """
    ดึงอัตราแลกเปลี่ยน THB -> JPY และ PHP -> JPY แบบ Real-time
    """
    default_rates = {"THB": 4.73, "PHP": 2.65}
    rates_out = {}
    
    # 1. ดึง THB -> JPY
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

    # 2. ดึง PHP -> JPY
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


# --- 1. ตาราง SLS TRANSPORTATION ---
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
        buying_price_jpy = float(buying_price_jpy)
        weight_g = float(weight_g)
        profit_rate_pct = float(profit_rate_pct)
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
        "cover_img": "URL รูปภาพปกหลัก (Cover Image)",
        "v1_name": "ชื่อตัวเลือกที่ 1 (เช่น สี / รุ่น)",
        "v1_opts": "รายการตัวเลือกที่ 1 (คั่นด้วย ,)",
        "v1_imgs": "URL รูปภาพตัวเลือกที่ 1 (1 รูปต่อ 1 ตัวเลือก คั่นด้วย ,)",
        "v1_imgs_help": "ระบุ URL รูปตามลำดับตัวเลือกที่ 1 เช่น ตัวเลือกคือ แดง, ดำ ให้ใส่ url_red, url_black (รูปนี้จะใช้กับทุกไซส์/ตัวเลือกที่ 2 ของสีนั้นๆ)",
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
        "cover_img": "Cover Image URL",
        "v1_name": "Variation 1 Name (e.g., Color)",
        "v1_opts": "Variation 1 Options (comma separated)",
        "v1_imgs": "Variation 1 Image URLs (1 per option, comma separated)",
        "v1_imgs_help": "Provide 1 image URL per Variation 1 option. It will automatically apply across all Variation 2 sub-options.",
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
        "cover_img": "メインカバー画像URL",
        "v1_name": "バリエーション1名称 (例: 色)",
        "v1_opts": "バリエーション1の選択肢 (カンマ区切り)",
        "v1_imgs": "バリエーション1の画像URL (各選択肢1枚、カンマ区切り)",
        "v1_imgs_help": "バリエーション1の各項目ごとに1枚のURLを入力してください（全サイズ/バリエーション2に共通適用されます）。",
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

# --- 3. GLOBAL CONTROL PANEL (REAL-TIME EXCHANGE RATE: THB/PHP -> JPY) ---
realtime_rates = fetch_base_rates()

st.subheader(T["calc_setting"])
col_cur, col_rate = st.columns(2)

with col_cur:
    currency = st.selectbox(T["currency_select"], ["THB", "PHP"])

with col_rate:
    current_realtime_rate = realtime_rates.get(currency, 4.73 if currency == "THB" else 2.65)
    rate_jpy = st.number_input(
        T["rate_label"].format(curr=currency), 
        value=current_realtime_rate, 
        format="%.4f",
        help=T["rate_info"]
    )
    st.caption(f"{T['rate_info']}: **1 {currency} = {rate_jpy} JPY**")

st.markdown("---")

# Initialize persistent session state for products
if "products" not in st.session_state:
    st.session_state.products = [
        {
            "id": 0,
            "category_id": "120039",
            "parent_sku": "361086-18",
            "integration_no": "",
            "brand": "No Brand",
            "product_name": T["default_pname"],
            "weight": 500.0,
            "product_desc": T["default_pdesc"],
            "cover_image": "https://example.com/cover.jpg",
            "v1_name": T["default_v1_name"],
            "v1_options": T["default_v1_opts"],
            "v1_images": "https://example.com/wine.jpg, https://example.com/white.jpg",
            "v2_name": T["default_v2_name"],
            "v2_options": T["default_v2_opts"],
        }
    ]

if "next_prod_id" not in st.session_state:
    st.session_state.next_prod_id = 1

col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button(T["add_product"]):
        new_id = st.session_state.next_prod_id
        st.session_state.next_prod_id += 1
        st.session_state.products.append({
            "id": new_id,
            "category_id": "100000",
            "parent_sku": f"ITEM-{new_id+1:03d}",
            "integration_no": "",
            "brand": "No Brand",
            "product_name": f"{T['product_num']} {len(st.session_state.products)+1}",
            "weight": 300.0,
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

# Loop over products
prod_index_to_remove = None

for idx, p in enumerate(st.session_state.products):
    p_id = p.get("id", idx) # ป้องกัน KeyError 'id' จาก session state เก่า
    st.markdown("---")
    col_title, col_del = st.columns([8, 2])
    with col_title:
        st.subheader(f"{T['product_num']}{idx + 1}: {p.get('product_name', '')}")
    with col_del:
        if len(st.session_state.products) > 1:
            if st.button(f"{T['del_product']}", key=f"del_btn_{p_id}"):
                prod_index_to_remove = idx

    c1, c2, c3 = st.columns(3)
    with c1:
        p["category_id"] = st.text_input(T["cat_id"], value=p.get("category_id", ""), key=f"cat_{p_id}")
        
        # 📌 แบ่งช่อง Parent SKU และ Variation Integration No. อยู่ข้างๆ กัน
        col_psku, col_integ = st.columns(2)
        with col_psku:
            p["parent_sku"] = st.text_input(T["parent_sku"], value=p.get("parent_sku", ""), key=f"psku_{p_id}")
        with col_integ:
            p["integration_no"] = st.text_input(T["integration_no"], value=p.get("integration_no", ""), key=f"integ_{p_id}")
            
        p["brand"] = st.text_input(T["brand"], value=p.get("brand", ""), key=f"brand_{p_id}")
    with c2:
        p["product_name"] = st.text_input(T["p_name"], value=p.get("product_name", ""), key=f"name_{p_id}")
        p["weight"] = st.number_input(T["weight"], value=float(p.get("weight", 300.0)), step=10.0, format="%.1f", key=f"w_{p_id}")
    with c3:
        p["product_desc"] = st.text_area(T["p_desc"], value=p.get("product_desc", ""), key=f"desc_{p_id}")

    p["cover_image"] = st.text_input(T["cover_img"], value=p.get("cover_image", ""), key=f"cimg_{p_id}")

    cv1, cv2 = st.columns(2)
    with cv1:
        p["v1_name"] = st.text_input(T["v1_name"], value=p.get("v1_name", ""), key=f"v1n_{p_id}")
        p["v1_options"] = st.text_input(T["v1_opts"], value=p.get("v1_options", ""), key=f"v1o_{p_id}")
        p["v1_images"] = st.text_input(T["v1_imgs"], value=p.get("v1_images", ""), help=T["v1_imgs_help"], key=f"v1i_{p_id}")
    with cv2:
        p["v2_name"] = st.text_input(T["v2_name"], value=p.get("v2_name", ""), key=f"v2n_{p_id}")
        p["v2_options"] = st.text_input(T["v2_opts"], value=p.get("v2_options", ""), key=f"v2o_{p_id}")

    list_v1 = [x.strip() for x in p["v1_options"].split(",") if x.strip()]
    list_v2 = [x.strip() for x in p["v2_options"].split(",") if x.strip()] if p["v2_name"] else [""]
    variations = list(itertools.product(list_v1, list_v2))

    # --- ส่วนการตั้งค่าด่วน แยกปุ่มแต่ละช่องอย่างเป็นอิสระ ---
    st.write(T["batch_title"])
    b_col1, b_col2, b_col3, b_col4 = st.columns(4)
    
    with b_col1:
        batch_cost = st.number_input(T["cost_col"], value=1590, step=100, key=f"b_cost_{p_id}")
        apply_cost = st.button(f"{T['btn_apply']} {T['cost_col']}", key=f"btn_apply_cost_{p_id}")
        
    with b_col2:
        batch_weight = st.number_input(T["weight_col"], value=float(p["weight"]), step=10.0, format="%.1f", key=f"b_weight_{p_id}")
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

    # สร้างหรืออัปเดตตารางตราบเท่าที่สินค้านี้ยังถูกเก็บไว้
    if df_state_key not in st.session_state:
        grid_data = []
        for opt1, opt2 in variations:
            var_title = f"{opt1}" + (f" / {opt2}" if opt2 else "")
            sku_suffix = f"-{opt1}" + (f"-{opt2}" if opt2 else "")

            grid_data.append({
                "Variation": var_title,
                "SKU": f"{p['parent_sku']}{sku_suffix}",
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
            sku_default = f"{p['parent_sku']}" + (f"-{opt1}" if opt1 else "") + (f"-{opt2}" if opt2 else "")
            
            match = df_existing[df_existing["Variation"] == var_title]
            
            if not match.empty:
                c_val = int(batch_cost) if apply_cost else match.iloc[0].get(cost_key, 1590)
                w_val = float(batch_weight) if apply_weight else match.iloc[0].get(weight_key, float(p["weight"]))
                p_val = float(batch_profit) if apply_profit else match.iloc[0].get(profit_key, 30.0)
                s_val = int(batch_stock) if apply_stock else match.iloc[0].get(stock_key, 2)
                sku_val = match.iloc[0].get("SKU", sku_default)
            else:
                c_val = int(batch_cost) if apply_cost else 1590
                w_val = float(batch_weight) if apply_weight else float(p["weight"])
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

    # คำนวณราคา Selling Price (THB)
    df_var[price_key] = df_var.apply(
        lambda row: calculate_net_price(
            buying_price_jpy=row[cost_key],
            weight_g=row[weight_key],
            profit_rate_pct=row[profit_key],
            currency=currency,
            rate_to_jpy=rate_jpy
        ), axis=1
    )

    st.write(T["grid_title"])
    edited_df = st.data_editor(
        df_var,
        column_config={
            "Variation": st.column_config.Column(disabled=True),
            "SKU": st.column_config.TextColumn(T["sku_col"], disabled=False),
            cost_key: st.column_config.NumberColumn(T["cost_col"], min_value=0, format="%d ¥"),
            weight_key: st.column_config.NumberColumn(T["weight_col"], min_value=1.0, format="%.1f g"),
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

    v1_imgs_list = [x.strip() for x in p["v1_images"].split(",") if x.strip()]
    updated_products_data.append({
        "cat_id": p["category_id"],
        "p_sku": p["parent_sku"],
        "integration_no": p["integration_no"],
        "brand": p["brand"],
        "p_name": p["product_name"],
        "weight": p["weight"],
        "p_desc": p["product_desc"],
        "cover_img": p["cover_image"],
        "v1_name": p["v1_name"],
        "v1_opts_list": list_v1,
        "v1_imgs": v1_imgs_list,
        "v2_name": p["v2_name"],
        "variations_table": edited_df,
        "cost_key": cost_key,
        "weight_key": weight_key,
        "price_key": price_key,
        "stock_key": stock_key
    })

# จัดการลบสินค้าเมื่อมีการกดปุ่มลบ
if prod_index_to_remove is not None:
    removed_p = st.session_state.products.pop(prod_index_to_remove)
    rem_id = removed_p.get("id", prod_index_to_remove)
    if f"df_data_{rem_id}" in st.session_state:
        del st.session_state[f"df_data_{rem_id}"]
    st.rerun()

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
        integration_no = p_data["integration_no"]
        p_name = p_data["p_name"]
        p_desc = p_data["p_desc"]
        cover_img = p_data["cover_img"]
        v1_name = p_data["v1_name"]
        v2_name = p_data["v2_name"]
        v1_opts_list = p_data["v1_opts_list"]
        v1_imgs = p_data["v1_imgs"]
        df_vars = p_data["variations_table"]
        
        c_k = p_data["cost_key"]
        w_k = p_data["weight_key"]
        pr_k = p_data["price_key"]
        st_k = p_data["stock_key"]

        v1_img_dict = {}
        for i, opt1_val in enumerate(v1_opts_list):
            v1_img_dict[opt1_val] = v1_imgs[i] if i < len(v1_imgs) else ""

        for idx, row in df_vars.iterrows():
            current_row = start_row
            
            ws.cell(row=current_row, column=1, value=cat_id)
            ws.cell(row=current_row, column=2, value=p_name if idx == 0 else "")
            ws.cell(row=current_row, column=3, value=p_desc if idx == 0 else "")
            ws.cell(row=current_row, column=9, value=integration_no) # 📌 บันทึก Variation Integration No. ลง Excel Column 9 (I)
            ws.cell(row=current_row, column=10, value=p_sku)
            
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
                ws.cell(row=current_row, column=21, value=cover_img)
                ws.cell(row=current_row, column=30, value=row[w_k] / 1000.0)
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