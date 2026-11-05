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

def upload_image_to_shopee(image_url):
    path = "/api/v2/product/upload_image"
    try:
        img_bytes = requests.get(image_url, timeout=10).content
    except Exception as e:
        return None, f"ไม่สามารถดาวน์โหลดรูปภาพจาก URL ได้: {e}"

    timestamp = int(time.time())
    sign = generate_shopee_sign(path, timestamp, SHOPEE_PARTNER_ID, SHOPEE_PARTNER_KEY, SHOPEE_ACCESS_TOKEN, SHOPEE_SHOP_ID)
    
    url = f"https://partner.shopeemobile.com{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&access_token={SHOPEE_ACCESS_TOKEN}&shop_id={SHOPEE_SHOP_ID}&sign={sign}"
    files = {'image': ('image.jpg', img_bytes, 'image/jpeg')}
    
    res = requests.post(url, files=files).json()
    if res.get("error") == "":
        return res.get("response", {}).get("image_info", {}).get("image_id"), None
    return None, res.get("message", "Upload Image Error")

# --- 3. PRICE & EXCHANGE RATE CALCULATION FORMULA ---
@st.cache_data(ttl=3600)
def fetch_base_rates():
    rates_out = {"THB": 4.73, "PHP": 2.65}
    try:
        req = urllib.request.Request("https://open.er-api.com/v6/latest/PHP", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            rates_out["PHP"] = round(data.get("rates", {}).get("JPY", 2.65), 4)
    except: pass
    return rates_out

def calculate_net_price(buying_price_jpy, weight_g, profit_rate_pct=30.0, currency="PHP", rate_to_jpy=2.65):
    if buying_price_jpy <= 0 or weight_g <= 0: return 0
    margin_factor = 1.0 - (profit_rate_pct / 100.0)
    if margin_factor <= 0: margin_factor = 0.01

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
    if not product_main: return None, "ไม่พบข้อมูลสินค้าบน Amazon หรือ API Key ไม่ถูกต้อง"

    main_title = product_main.get('title', '')
    main_desc = product_main.get('description', main_title)
    cover_images = [img.get('link') for img in product_main.get('images', [])]
    variants_list = product_main.get('variants', [])

    extracted_vars = []

    if variants_list:
        for var in variants_list:
            child_asin = var.get('asin')
            if not child_asin: continue

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
    st.info("🔄 กำลังอัปโหลดรูปภาพเข้าเซิร์ฟเวอร์ Shopee...")
    
    img_ids = []
    last_err = ""
    for img_url in product_data['images'][:8]:
        img_id, err = upload_image_to_shopee(img_url)
        if img_id: 
            img_ids.append(img_id)
        else:
            last_err = err
        
    if not img_ids:
        return False, f"อัปโหลดรูปภาพไม่สำเร็จ: {last_err}"

    st.info("📦 กำลังสร้างสินค้าหลักบน Shopee...")
    item_payload = {
        "category_id": int(category_id),
        "item_name": product_data['title'][:120],
        "description": product_data['description'][:3000],
        "item_status": "UNLISTED",
        "weight": float(weight_g / 1000.0),
        "image": {"image_id_list": img_ids},
        "brand": {"brand_id": 0, "original_brand_name": "NoBrand"}
    }
    
    item_res = call_shopee_api("/api/v2/product/add_item", item_payload)
    item_id = item_res.get("response", {}).get("item_id")
    
    if not item_id:
        return False, f"สร้างสินค้าไม่สำเร็จ: {item_res.get('message')}"

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
        return False, f"ตั้งค่า Variation ล้มเหลว: {tier_res.get('message')}"

# --- 6. STREAMLIT USER INTERFACE ---
st.title("🚀 Amazon ASIN to Shopee Direct Auto-Lister")
st.caption("ระบบดึงข้อมูล ASIN Amazon คำนวณราคาขายอัตโนมัติ และ List เข้าร้านค้า Shopee ผ่าน API")

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
