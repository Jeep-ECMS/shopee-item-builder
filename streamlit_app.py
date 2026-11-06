import streamlit as st
import requests
import json
import time
import hmac
import hashlib
import urllib.request
import pandas as pd

st.set_page_config(page_title="Shopee Auto Listing & Price Calculator", layout="wide")

# --- 1. CONFIG & KEYS FROM SECRETS ---
RAINFOREST_API_KEY = st.secrets.get("rainforest", {}).get("api_key", "")
SHOPEE_PARTNER_ID = st.secrets.get("shopee", {}).get("partner_id", 0)
SHOPEE_PARTNER_KEY = st.secrets.get("shopee", {}).get("partner_key", "")
SHOPEE_SHOP_ID = st.secrets.get("shopee", {}).get("shop_id", 0)
SHOPEE_ACCESS_TOKEN = st.secrets.get("shopee", {}).get("access_token", "")

# --- Helper: แลก Code เป็น Access Token ---
def exchange_code_for_token(code_val, shop_id_val):
    path = "/api/v2/auth/token/get"
    timestamp = int(time.time())
    
    base_string = f"{SHOPEE_PARTNER_ID}{path}{timestamp}"
    sign = hmac.new(
        SHOPEE_PARTNER_KEY.encode('utf-8'),
        base_string.encode('utf-8'),
        hashlib.sha256
    ).hexdigest()
    
    url = f"https://partner.shopeemobile.com{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&sign={sign}"
    
    payload = {
        "code": code_val.strip(),
        "partner_id": int(SHOPEE_PARTNER_ID),
        "shop_id": int(shop_id_val)
    }
    
    headers = {"Content-Type": "application/json"}
    res = requests.post(url, headers=headers, data=json.dumps(payload)).json()
    return res

# --- 2. SHOPEE OPEN API SIGNATURE HELPER ---
def generate_shopee_sign(path, timestamp, partner_id, partner_key, access_token="", shop_id=0):
    if access_token and shop_id:
        base_string = f"{partner_id}{path}{timestamp}{access_token}{shop_id}"
    else:
        base_string = f"{partner_id}{path}{timestamp}"
    
    sign = hmac.new(
        partner_key.encode('utf-8'),
        base_string.encode('utf-8'),
        hashlib.sha256
    ).hexdigest()
    return sign

def call_shopee_api(path, payload):
    host = "https://partner.shopeemobile.com"
    timestamp = int(time.time())
    sign = generate_shopee_sign(path, timestamp, SHOPEE_PARTNER_ID, SHOPEE_PARTNER_KEY, SHOPEE_ACCESS_TOKEN, SHOPEE_SHOP_ID)
    
    url = f"{host}{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&access_token={SHOPEE_ACCESS_TOKEN}&shop_id={SHOPEE_SHOP_ID}&sign={sign}"
    headers = {"Content-Type": "application/json"}
    
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    return response.json()

# --- 3. PRICE & EXCHANGE RATE CALCULATION FORMULA ---
@st.cache_data(ttl=3600)
def fetch_base_rates():
    rates_out = {"THB": 4.73, "PHP": 2.65}
    try:
        req = urllib.request.Request("https://open.er-api.com/v6/latest/PHP", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            rates_out["PHP"] = round(data.get("rates", {}).get("JPY", 2.65), 4)
    except:
        pass
    return rates_out

def calculate_net_price(buying_price_jpy, weight_g, profit_rate_pct=30.0, currency="PHP", rate_to_jpy=2.65):
    if buying_price_jpy <= 0 or weight_g <= 0:
        return 0
    margin_factor = 1.0 - (profit_rate_pct / 100.0)
    if margin_factor <= 0:
        margin_factor = 0.01

    if currency == "PHP":
        buying_price_php = buying_price_jpy / rate_to_jpy if rate_to_jpy > 0 else 0
        transportation_jp_php = 116.0
        sls_fee = 76 + max(0, int((weight_g - 50) // 50)) * 25
        base_price = (sls_fee + buying_price_php + transportation_jp_php) / margin_factor
        cif_check = (base_price * 1.01) + (1369.0 * (weight_g / 1000.0))
        if cif_check >= 10000:
            custom_tax_part = 1369.0 * (weight_g / 1000.0) * 0.12
            net_price = (sls_fee + custom_tax_part + buying_price_php + transportation_jp_php) / 0.5788
        else:
            net_price = base_price
        return round(net_price)
    return 0

# --- 4. RAINFOREST AMAZON SCRAPING (METHOD B) ---
def fetch_amazon_data_method_b(parent_asin):
    url = "https://api.rainforestapi.com/request"
    params = {
        'api_key': RAINFOREST_API_KEY,
        'type': 'product',
        'amazon_domain': 'amazon.co.jp',
        'asin': parent_asin
    }
    
    parent_res = requests.get(url, params=params).json()
    product_main = parent_res.get('product', {})
    if not product_main:
        return None, "ไม่พบข้อมูลสินค้าบน Amazon หรือ API Key ไม่ถูกต้อง"

    main_title = product_main.get('title', '')
    main_desc = product_main.get('description', main_title)
    cover_images = [img.get('link') for img in product_main.get('images', [])]
    variants_list = product_main.get('variants', [])

    extracted_vars = []

    if variants_list:
        for var in variants_list:
            child_asin = var.get('asin')
            if not child_asin:
                continue

            child_params = {
                'api_key': RAINFOREST_API_KEY,
                'type': 'product',
                'amazon_domain': 'amazon.co.jp',
                'asin': child_asin
            }
            c_res = requests.get(url, params=child_params).json().get('product', {})
            price_val = c_res.get('buybox_winner', {}).get('price', {}).get('value', 0)
            
            dims = c_res.get('variants_dimensions', [])
            opt1 = dims[0].get('value', '') if len(dims) > 0 else ''
            opt2 = dims[1].get('value', '') if len(dims) > 1 else ''

            extracted_vars.append({
                'asin': child_asin,
                'opt1': opt1 if opt1 else 'Default',
                'opt2': opt2,
                'cost_jpy': int(price_val),
                'image': c_res.get('main_image', {}).get('link', '')
            })
    else:
        price_val = product_main.get('buybox_winner', {}).get('price', {}).get('value', 0)
        extracted_vars.append({
            'asin': parent_asin,
            'opt1': 'Default',
            'opt2': '',
            'cost_jpy': int(price_val),
            'image': product_main.get('main_image', {}).get('link', '')
        })

    return {
        'parent_asin': parent_asin,
        'title': main_title,
        'description': main_desc,
        'images': cover_images,
        'variations': extracted_vars
    }, "Success"

# --- 5. SHOPEE AUTOMATED DIRECT LISTING WORKFLOW ---
def publish_to_shopee(product_data, category_id, weight_g, profit_rate, currency_rate):
    st.info("📦 กำลังส่งข้อมูลสินค้าเข้าร้าน Shopee...")

    # ส่ง URL รูปภาพเข้า Shopee โดยตรง
    images_to_upload = product_data.get('images', [])
    if not images_to_upload and product_data.get('variations'):
        images_to_upload = [v['image'] for v in product_data['variations'] if v.get('image')]

    image_urls = [url for url in images_to_upload[:8] if url]

    item_payload = {
        "category_id": int(category_id),
        "item_name": product_data['title'][:120],
        "description": product_data['description'][:3000] if product_data['description'] else product_data['title'][:120],
        "item_status": "UNLISTED",
        "weight": float(weight_g / 1000.0),
        "image": {"image_url_list": image_urls},
        "brand": {"brand_id": 0, "original_brand_name": "NoBrand"}
    }
    
    item_res = call_shopee_api("/api/v2/product/add_item", item_payload)
    
    # ตรวจสอบ Error Response จาก Shopee API
    if item_res.get("error") != "":
        error_code = item_res.get("error")
        error_msg = item_res.get("message", item_res.get("warning", ""))
        return False, f"สร้างสินค้าไม่สำเร็จ (Shopee API ตอบกลับ: {error_code} - {error_msg})"

    item_id = item_res.get("response", {}).get("item_id")
    if not item_id:
        return False, f"สร้างสินค้าไม่สำเร็จ: {item_res}"

    st.info("⚙️ กำลังตั้งค่า Variation และราคาขาย...")
    
    model_list = []
    opt1_names = list(set([v['opt1'] for v in product_data['variations']]))
    
    for idx, var in enumerate(product_data['variations']):
        selling_price = calculate_net_price(var['cost_jpy'], weight_g, profit_rate, "PHP", currency_rate)
        opt1_idx = opt1_names.index(var['opt1'])
        
        model_list.append({
            "tier_index": [opt1_idx],
            "normal_price": float(selling_price),
            "original_price": float(selling_price),
            "stock": 10,
            "seller_sku": f"{product_data['parent_asin']}-{var['asin']}"
        })

    tier_payload = {
        "item_id": item_id,
        "tier_variation": [
            {
                "name": "Variation",
                "option_list": [{"option": opt} for opt in opt1_names]
            }
        ],
        "model": model_list
    }

    tier_res = call_shopee_api("/api/v2/product/init_tier_variation", tier_payload)
    if tier_res.get("error") == "":
        return True, f"สำเร็จ! ลงสินค้าบน Shopee เรียบร้อยแล้ว (Item ID: {item_id})"
    else:
        return False, f"ตั้งค่า Variation ล้มเหลว: {tier_res.get('message', tier_res.get('error'))}"

# --- 6. STREAMLIT USER INTERFACE ---
st.title("🚀 Amazon ASIN to Shopee Direct Auto-Lister")
st.caption("ระบบดึงข้อมูล ASIN Amazon คำนวณราคาขายอัตโนมัติ และ List เข้าร้านค้า Shopee ผ่าน API")

with st.expander("🔑 เครื่องมือแลก Shopee Access Token (คลิกเพื่อเปิด)"):
    col_c, col_s = st.columns(2)
    with col_c:
        input_code = st.text_input("วางรหัส Code จาก URL:", value="")
    with col_s:
        input_shop = st.text_input("Shop ID:", value="1376701906")
    
    if st.button("🔑 กดเพื่อแลก Access Token"):
        if input_code and input_shop:
            token_res = exchange_code_for_token(input_code, input_shop)
            if token_res.get("error") == "":
                access_t = token_res.get("access_token", "")
                st.success(f"✅ ได้รับ Access Token สำเร็จ!")
                st.code(access_t, language="text")
                st.info("ให้นำ Access Token ข้างบนนี้ไปวางใน Secrets ของ Streamlit Cloud ได้เลยครับ")
            else:
                st.error(f"❌ ล้มเหลว: {token_res.get('message', token_res.get('error'))}")
        else:
            st.warning("กรุณากรอก Code และ Shop ID ให้ครบถ้วน")

st.markdown("---")

rates = fetch_base_rates()
rates_php = rates.get("PHP", 2.65)

col_input, col_config = st.columns([6, 4])

with col_input:
    asin_code = st.text_input("กรอกรหัส Amazon ASIN (Parent ASIN):", value="B073JYC4XM")
    category_id = st.text_input("Shopee Category ID Target:", value="120039")

with col_config:
    weight_input = st.number_input("น้ำหนักสินค้าโดยประมาณ (กรัม):", value=300.0, step=50.0)
    profit_input = st.number_input("อัตรากำไร Profit Margin (%):", value=30.0, step=5.0)
    st.write(f"💱 อัตราแลกเปลี่ยนปัจจุบัน: **1 PHP = {rates_php} JPY**")

if st.button("🔍 1. ดึงข้อมูลสินค้าจาก Amazon (Rainforest API)", type="primary"):
    with st.spinner("กำลังดึงข้อมูลสินค้าและตัวเลือก Child ASIN จาก Amazon..."):
        data, err_msg = fetch_amazon_data_method_b(asin_code)
        if data:
            st.session_state["fetched_product"] = data
            st.success("ดึงข้อมูลสินค้าสำเร็จ!")
        else:
            st.error(err_msg)

if "fetched_product" in st.session_state:
    pdata = st.session_state["fetched_product"]
    st.markdown("---")
    st.subheader(f"📌 {pdata['title']}")
    st.text_area("รายละเอียดสินค้า", pdata['description'], height=100)
    
    preview_rows = []
    for v in pdata['variations']:
        sell_p = calculate_net_price(v['cost_jpy'], weight_input, profit_input, "PHP", rates_php)
        preview_rows.append({
            "Child ASIN": v['asin'],
            "Option 1": v['opt1'],
            "ทุน (JPY)": v['cost_jpy'],
            "ราคาขาย Shopee (PHP)": sell_p
        })
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
    
    st.markdown("---")
    if st.button("🚀 2. อัปโหลดสินค้าเข้าร้าน Shopee ทันที (Direct Publish)", type="secondary"):
        with st.spinner("กำลังส่งข้อมูลเข้า Shopee Open API Console..."):
            success, msg = publish_to_shopee(pdata, category_id, weight_input, profit_input, rates_php)
            if success:
                st.balloons()
                st.success(msg)
            else:
                st.error(msg)
