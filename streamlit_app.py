import streamlit as st
import pandas as pd
import openpyxl
import io
import math

st.set_page_config(page_title="Shopee Mass Upload & Auto Price Calculator", layout="wide")

# --- 1. LOGIC การคำนวณค่าขนส่ง SLS & NET PRICE ---

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
    try:
        buying_price_jpy = float(buying_price_jpy)
        weight_g = float(weight_g)
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
        
        if cif_value >= 10000:
            net_price = (sls_fee + (1369 * (weight_g / 1000) * 0.12) + buying_price_php + transportation_jp_php) / 0.5788
        else:
            net_price = temp_price
            
        return round(net_price)

    return 0

# --- 2. STREAMLIT UI ---

st.title("🛍️ Shopee Mass Upload & Auto Price Calculator (ครบจบในหน้าเดียว)")
st.caption("กรอกข้อมูลสินค้า เลือกตลาดเป้าหมาย คำนวณราคาอัตโนมัติ และส่งออกไฟล์ Excel ได้ทันที")

# Control Panel
st.subheader("⚙️ 1. ตั้งค่าการคำนวณราคา (Target Market & Exchange Rates)")
col1, col2, col3 = st.columns(3)

with col1:
    currency = st.selectbox("เลือกตลาดเป้าหมาย (Currency)", ["THB", "PHP"])

with col2:
    default_rate = 4.868196 if currency == "THB" else 2.590111
    rate_jpy = st.number_input(f"อัตราแลกเปลี่ยน ({currency}/JPY)", value=default_rate, format="%.6f")

with col3:
    default_weight = st.number_input("น้ำหนักเริ่มต้นสินค้า (กรัม)", value=300, step=50)

st.markdown("---")
st.subheader("📝 2. จัดการข้อมูลสินค้า / Variation / ต้นทุน / คำนวณราคาขาย")

# โครงสร้างข้อมูลเริ่มต้น
default_data = [
    {
        "Category ID": 1001,
        "Parent SKU": "SHIRT-001",
        "Product Name": "เสื้อเชิ้ตลายสก๊อต Cotton 100%",
        "Description": "เสื้อเชิ้ตคุณภาพสูง นำเข้าจากญี่ปุ่น",
        "Variation 1 Name": "สี",
        "Variation 1 Option": "Red",
        "Variation 2 Name": "ไซส์",
        "Variation 2 Option": "S",
        "SKU": "SHIRT-001-RED-S",
        "Buying Price (JPY)": 1200,
        "Weight (g)": default_weight,
        "Stock": 50
    },
    {
        "Category ID": 1001,
        "Parent SKU": "SHIRT-001",
        "Product Name": "เสื้อเชิ้ตลายสก๊อต Cotton 100%",
        "Description": "เสื้อเชิ้ตคุณภาพสูง นำเข้าจากญี่ปุ่น",
        "Variation 1 Name": "สี",
        "Variation 1 Option": "Red",
        "Variation 2 Name": "ไซส์",
        "Variation 2 Option": "M",
        "SKU": "SHIRT-001-RED-M",
        "Buying Price (JPY)": 1200,
        "Weight (g)": default_weight,
        "Stock": 50
    }
]

# Reset Data หาก Key ขาดหาย
if "product_data" not in st.session_state or not isinstance(st.session_state.product_data, pd.DataFrame):
    st.session_state.product_data = pd.DataFrame(default_data)

df_display = st.session_state.product_data.copy()

# ตรวจสอบคอลัมน์ที่จำเป็น
required_cols = ["Parent SKU", "Product Name", "Variation 1 Option", "Variation 2 Option", "SKU", "Buying Price (JPY)", "Weight (g)", "Stock", "Category ID", "Description"]
for col in required_cols:
    if col not in df_display.columns:
        df_display[col] = ""

# คำนวณ Net Price
price_col_name = f"Selling Price ({currency})"
df_display[price_col_name] = df_display.apply(
    lambda row: calculate_net_price(
        buying_price_jpy=row.get("Buying Price (JPY)", 0),
        weight_g=row.get("Weight (g)", default_weight),
        currency=currency,
        rate_jpy=rate_jpy
    ), axis=1
)

# ลิสต์คอลัมน์ตามลำดับที่ถูกต้อง
cols_order = [
    "Parent SKU", "Product Name", "Variation 1 Option", "Variation 2 Option", 
    "SKU", "Buying Price (JPY)", "Weight (g)", price_col_name, "Stock", 
    "Category ID", "Description"
]

# แสดงผล Data Editor
edited_df = st.data_editor(
    df_display[cols_order],
    column_config={
        "Buying Price (JPY)": st.column_config.NumberColumn("Buying Price (JPY)", min_value=0, format="%d ¥"),
        "Weight (g)": st.column_config.NumberColumn("Weight (g)", min_value=1, format="%d g"),
        price_col_name: st.column_config.NumberColumn(f"Net Price ({currency})", disabled=True, format="%d " + currency),
        "Stock": st.column_config.NumberColumn("Stock", min_value=0, format="%d"),
    },
    use_container_width=True,
    num_rows="dynamic"
)

st.session_state.product_data = edited_df

# Export Excel
st.markdown("---")
st.subheader("📥 3. ดาวน์โหลดไฟล์ Excel สำหรับ Mass Upload")

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    edited_df.to_excel(writer, index=False, sheet_name='Mass Upload')

st.download_button(
    label=f"🚀 ส่งออกไฟล์ Excel สำหรับ Shopee ({currency})",
    data=buffer.getvalue(),
    file_name=f"Shopee_Mass_Upload_{currency}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True
)