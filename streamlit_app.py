import streamlit as st
import pandas as pd
import openpyxl
import io
import math

st.set_page_config(page_title="Shopee Mass Upload & Price Calculator", layout="wide")

# --- 1. ระบบจัดการภาษา (TH / EN / JA) ---
LANG_TEXTS = {
    "TH": {
        "title": "🛍️ เครื่องมือสร้างไฟล์ Mass Upload Shopee & คำนวณราคาขาย (THB / PHP)",
        "currency_select": "เลือกตลาดเป้าหมาย (Currency)",
        "rate_label": "อัตราแลกเปลี่ยน (JPY)",
        "default_weight": "น้ำหนักเริ่มต้นสินค้า (กรัม)",
        "add_product": "➕ เพิ่มรายการสินค้า",
        "export_btn": "🚀 ส่งออกไฟล์ Excel สำหรับ Shopee Mass Upload",
        "variation_header": "กำหนดรายการสินค้าและ Variation",
    },
    "EN": {
        "title": "🛍️ Shopee Mass Upload Generator & Price Calculator (THB / PHP)",
        "currency_select": "Select Target Market (Currency)",
        "rate_label": "Exchange Rate (JPY)",
        "default_weight": "Default Weight (grams)",
        "add_product": "➕ Add Product",
        "export_btn": "🚀 Export Excel for Shopee Mass Upload",
        "variation_header": "Product Variations & Pricing",
    },
    "JA": {
        "title": "🛍️ Shopee 一括アップロード作成 & 価格計算ツール (THB / PHP)",
        "currency_select": "ターゲット市場選択 (通貨)",
        "rate_label": "為替レート (JPY)",
        "default_weight": "デフォルト weight (g)",
        "add_product": "➕ 商品を追加",
        "export_btn": "🚀 Shopee用 Excel エกสポート",
        "variation_header": "バリエーション・価格設定",
    }
}

# --- 2. LOGIC การคำนวณค่าขนส่ง SLS & NET PRICE ---

# ตารางค่าขนส่ง SLS สำหรับประเทศไทย (THB)
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

def calculate_net_price(buying_price_jpy, weight_g, currency="THB", rate_jpy=None):
    if not buying_price_jpy or buying_price_jpy <= 0:
        return 0
    if not weight_g or weight_g <= 0:
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
        
        if cif_value >= 10000:
            net_price = (sls_fee + (1369 * (weight_g / 1000) * 0.12) + buying_price_php + transportation_jp_php) / 0.5788
        else:
            net_price = temp_price
            
        return round(net_price)

    return 0


# --- 3. STREAMLIT UI INTERFACE ---

# เลือกภาษาใน Sidebar
lang = st.sidebar.selectbox("🌐 Language / ภาษา", ["TH", "EN", "JA"])
t = LANG_TEXTS[lang]

st.title(t["title"])

st.markdown("---")
# Control Panel
col1, col2, col3 = st.columns(3)

with col1:
    currency = st.selectbox(t["currency_select"], ["THB", "PHP"])

with col2:
    default_rate = 4.868196 if currency == "THB" else 2.590111
    rate_jpy = st.number_input(f"{t['rate_label']} ({currency}/JPY)", value=default_rate, format="%.6f")

with col3:
    default_weight = st.number_input(t["default_weight"], value=300, step=50)

st.subheader(t["variation_header"])

# ข้อมูลตัวอย่างเริ่มต้น
if "product_data" not in st.session_state:
    st.session_state.product_data = pd.DataFrame([
        {"Parent SKU": "MODEL-001", "Variation 1": "Red", "Variation 2": "S", "SKU": "MODEL-001-RED-S", "Buying Price (JPY)": 990, "Weight (g)": default_weight, "Stock": 100},
        {"Parent SKU": "MODEL-001", "Variation 1": "Red", "Variation 2": "M", "SKU": "MODEL-001-RED-M", "Buying Price (JPY)": 990, "Weight (g)": default_weight, "Stock": 100},
        {"Parent SKU": "MODEL-001", "Variation 1": "Black", "Variation 2": "S", "SKU": "MODEL-001-BLK-S", "Buying Price (JPY)": 1200, "Weight (g)": default_weight, "Stock": 100},
    ])

# คำนวณราคาขายอัตโนมัติ
df_display = st.session_state.product_data.copy()
df_display[f"Calculated Price ({currency})"] = df_display.apply(
    lambda row: calculate_net_price(
        buying_price_jpy=row["Buying Price (JPY)"],
        weight_g=row["Weight (g)"],
        currency=currency,
        rate_jpy=rate_jpy
    ), axis=1
)

# ตารางแก้ไขข้อมูล (Data Editor)
edited_df = st.data_editor(
    df_display,
    column_config={
        "Buying Price (JPY)": st.column_config.NumberColumn("Buying Price (JPY)", min_value=0, format="%d ¥"),
        "Weight (g)": st.column_config.NumberColumn("Weight (g)", min_value=1, format="%d g"),
        f"Calculated Price ({currency})": st.column_config.NumberColumn(f"Calculated Price ({currency})", disabled=True, format="%d " + currency),
    },
    use_container_width=True,
    num_rows="dynamic"
)

st.session_state.product_data = edited_df

# --- 4. EXPORT EXCEL GENERATOR ---
st.markdown("---")

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    edited_df.to_excel(writer, index=False, sheet_name='Mass Upload')

st.download_button(
    label=t["export_btn"],
    data=buffer.getvalue(),
    file_name=f"Shopee_Mass_Upload_{currency}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True
)