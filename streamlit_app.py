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

# --- Initialize Session State Token ---
if "live_access_token" not in st.session_state:
    st.session_state["live_access_token"] = st.secrets.get("shopee", {}).get("access_token", "")

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
    return requests.post(url, headers=headers, data=json.dumps(payload)).json()

# --- 2. SHOPEE OPEN API SIGNATURE HELPER ---
def generate_shopee_sign(path, timestamp, partner_id, partner_key, access_token="", shop_id=0):
    if access_token and shop_id:
        base_string = f"{partner_id}{path}{timestamp}{access_token}{shop_id}"
    else:
        base_string = f"{partner_id}{path}{timestamp}"
    
    return hmac.new(
        partner_key.encode('utf-8'),
        base_string.encode('utf-8'),
        hashlib.sha256
    ).hexdigest()

def call_shopee_api(path, payload):
    host = "https://partner.shopeemobile.com"
    timestamp = int(time.time())
    token = st.session_state.get("live_access_token", "")
    sign = generate_shopee_sign(path, timestamp, SHOPEE_PARTNER_ID, SHOPEE_PARTNER_KEY, token, SHOPEE_SHOP_ID)
    url = f"{host}{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&access_token={token}&shop_id={SHOPEE_SHOP_ID}&sign={sign}"
    headers = {"Content-Type": "application/json"}
    return requests.post(url, headers=headers, data=json.dumps(payload)).json()

def upload_image_file_to_shopee(image_url):
    """ ดาวน์โหลดรูปภาพและอัปโหลดเป็น File Multipart เพื่อรับ image_id """
    path = "/api/v2/product/upload_image"
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        img_res = requests.get(image_url, headers=headers, timeout=15)
        if img_res.status_code != 200:
            return None, f"HTTP Error {img_res.status_code}"
        img_bytes = img_res.content
    except Exception as e:
        return None, str(e)

    timestamp = int(time.time())
    token = st.session_state.get("live_access_token", "")
    sign = generate_shopee_sign(path, timestamp, SHOPEE_PARTNER_ID, SHOPEE_PARTNER_KEY, token, SHOPEE_SHOP_ID)
    url = f"https://partner.shopeemobile.com{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&access_token={token}&shop_id={SHOPEE_SHOP_ID}&sign={sign}"
    
    files = {'image': ('image.jpg', img_bytes, 'image/jpeg')}
    try:
        res = requests.post(url, files=files, timeout=30).json()
        if res.get("error") == "":
            return res.get("response", {}).get("image_info", {}).get("image_id"), None
        else:
            return None, f"Shopee Error: {res.get('message', res.get('error'))}"
    except Exception as e:
        return None, str(e)

# --- 3. PRICE & THBJPY EXCHANGE RATE CALCULATION FORMULA ---
@st.cache_data(ttl=3600)
def fetch_thb_jpy_rate():
    default_middle = 4.72
    try:
        req = urllib.request.Request("https://open.er-api.com/v6/latest/THB", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            tt_middle = float(data.get("rates", {}).get("JPY", default_middle))
    except:
        tt_middle = default_middle

    ttb = round(tt_middle - 0.12, 2)
    exchange_rate_final = round(ttb / 1.025, 2)
    return {
        "tt_middle": tt_middle,
        "ttb": ttb,
        "exchange_rate": exchange_rate_final
    }

def calculate_net_price_thb(buying_price_jpy, weight_g, profit_rate_pct=30.0, rate_thb_to_jpy=4.49):
    if buying_price_jpy <= 0: return 0
    margin_factor = 1.0 - (profit_rate_pct / 100.0)
    if margin_factor <= 0: margin_factor = 0.01

    buying_price_thb = buying_price_jpy / rate_thb_to_jpy if rate_thb_to_jpy > 0 else 0
    shipping_thb = max(45, int((weight_g / 100.0) * 25))
    base_price = (buying_price_thb + shipping_thb) / margin_factor
    return round(base_price)

# --- 4. ADVANCED AMAZON DATA SCRAPING ---
def fetch_amazon_data_advanced(parent_asin):
    url = "https://api.rainforestapi.com/request"
    params = {
        'api_key': RAINFOREST_API_KEY,
        'type': 'product',
        'amazon_domain': 'amazon.co.jp',
        'asin': parent_asin,
        'language': 'en_US'
    }
    
    parent_res = requests.get(url, params=params).json()
    product_main = parent_res.get('product', {})
    if not product_main:
        return None, "ไม่พบข้อมูลสินค้าบน Amazon หรือ API Key ไม่ถูกต้อง"

    main_title = product_main.get('title', '')
    main_desc = product_main.get('description', main_title)
    jan_code = product_main.get('ean', '') or product_main.get('gtin', '') or product_main.get('upc', '')
    
    categories = [cat.get('name') for cat in product_main.get('categories', [])]
    category_path = " > ".join(categories) if categories else "Spare Parts and Accessories"

    weight_g = 300.0
    len_cm, width_cm, height_cm = 10.0, 10.0, 10.0
    
    weight_obj = product_main.get('weight_shipping', {}) or product_main.get('weight', {})
    if isinstance(weight_obj, dict):
        if 'raw' in weight_obj:
            if 'g' in weight_obj['raw'].lower() and 'kg' not in weight_obj['raw'].lower():
                try: weight_g = float(weight_obj['raw'].lower().replace('g','').strip())
                except: pass
            elif 'kg' in weight_obj['raw'].lower():
                try: weight_g = float(weight_obj['raw'].lower().replace('kg','').strip()) * 1000.0
                except: pass

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
                'asin': child_asin,
                'language': 'en_US'
            }
            c_res = requests.get(url, params=child_params).json().get('product', {})
            price_val = c_res.get('buybox_winner', {}).get('price', {}).get('value', 0)
            child_jan = c_res.get('ean', jan_code)
            
            dims = c_res.get('variants_dimensions', [])
            opt1 = dims[0].get('value', '') if len(dims) > 0 else ''
            opt2 = dims[1].get('value', '') if len(dims) > 1 else ''

            extracted_vars.append({
                'asin': child_asin,
                'jan': child_jan,
                'opt1': opt1 if opt1 else 'Default',
                'opt2': opt2,
                'cost_jpy': int(price_val),
                'image': c_res.get('main_image', {}).get('link', '')
            })
    else:
        price_val = product_main.get('buybox_winner', {}).get('price', {}).get('value', 0)
        extracted_vars.append({
            'asin': parent_asin,
            'jan': jan_code,
            'opt1': 'Default',
            'opt2': '',
            'cost_jpy': int(price_val),
            'image': product_main.get('main_image', {}).get('link', '')
        })

    return {
        'parent_asin': parent_asin,
        'jan_code': jan_code,
        'title': main_title,
        'description': main_desc,
        'category_path': category_path,
        'weight_g': weight_g,
        'length_cm': len_cm,
        'width_cm': width_cm,
        'height_cm': height_cm,
        'images': cover_images,
        'variations': extracted_vars
    }, "Success"

# --- 5. SHOPEE AUTOMATED DIRECT LISTING WORKFLOW ---
def publish_to_shopee(product_data, category_id, weight_g, profit_rate, rate_thb_to_jpy, len_cm, width_cm, height_cm):
    st.info("🔄 กำลังแปลงไฟล์รูปภาพและอัปโหลดเข้าเซิร์ฟเวอร์ Shopee...")
    
    img_ids = []
    last_err = ""
    
    images_to_upload = product_data.get('images', [])
    if not images_to_upload and product_data.get('variations'):
        images_to_upload = [v['image'] for v in product_data['variations'] if v.get('image')]

    for img_url in images_to_upload[:8]:
        if not img_url: continue
        img_id, err = upload_image_file_to_shopee(img_url)
        if img_id: 
            img_ids.append(img_id)
        else:
            last_err = err
        
    if not img_ids:
        return False, f"อัปโหลดรูปภาพไม่สำเร็จ ({last_err})"

    st.info("📦 กำลังสร้างสินค้าหลักบน Shopee...")
    
    item_payload = {
        "category_id": int(category_id),
        "item_name": product_data['title'][:120],
        "description": f"[P{product_data['parent_asin']}]\n\n{product_data['description'][:2800]}",
        "item_status": "UNLISTED",
        "weight": float(weight_g / 1000.0),
        "dimension": {
            "package_length": int(len_cm),
            "package_width": int(width_cm),
            "package_height": int(height_cm)
        },
        "image": {"image_id_list": img_ids},
        "brand": {"brand_id": 0, "original_brand_name": "NoBrand"}
    }
    
    item_res = call_shopee_api("/api/v2/product/add_item", item_payload)
    item_id = item_res.get("response", {}).get("item_id")
    
    if not item_id:
        error_msg = item_res.get("message", item_res.get("error"))
        return False, f"สร้างสินค้าไม่สำเร็จ: {error_msg}"

    st.info("⚙️ กำลังตั้งค่า Variation และราคาขาย (THB)...")
    
    model_list = []
    opt1_names = list(set([v['opt1'] for v in product_data['variations']]))
    
    for idx, var in enumerate(product_data['variations']):
        selling_price_thb = calculate_net_price_thb(var['cost_jpy'], weight_g, profit_rate, rate_thb_to_jpy)
        opt1_idx = opt1_names.index(var['opt1'])
        
        model_list.append({
            "tier_index": [opt1_idx],
            "normal_price": float(selling_price_thb),
            "original_price": float(selling_price_thb),
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
st.title("🚀 Amazon ASIN to Shopee Direct Auto-Lister (THB Price)")
st.caption("ดึงข้อมูล Amazon + คำนวณราคาขาย THB อัตโนมัติ (Exchange Rate = TTB / 1.025) + List ลง Shopee")

rate_info = fetch_thb_jpy_rate()

with st.expander("🔑 เครื่องมือแลก/อัปเดต Shopee Access Token (ใช้งานสดทันที)"):
    col_c, col_s = st.columns(2)
    with col_c:
        input_code = st.text_input("วางรหัส Code ล่าสุดจาก URL:", value="")
    with col_s:
        input_shop = st.text_input("Shop ID:", value="1376701906")
    
    if st.button("🔑 กดเพื่อแลก Token สดสำหรับ session นี้"):
        if input_code and input_shop:
            token_res = exchange_code_for_token(input_code, input_shop)
            if token_res.get("error") == "":
                access_t = token_res.get("access_token", "")
                st.session_state["live_access_token"] = access_t
                st.success(f"✅ เปิดใช้ Access Token สดสำเร็จ! (Token ถูกล็อกไว้ในระบบแล้ว)")
                st.code(access_t, language="text")
            else:
                st.error(f"❌ ล้มเหลว: {token_res.get('message', token_res.get('error'))}")
        else:
            st.warning("กรุณากรอก Code และ Shop ID ให้ครบถ้วน")

st.markdown("---")

col_input, col_config = st.columns([6, 4])

with col_input:
    asin_code = st.text_input("กรอกรหัส Amazon ASIN (Parent ASIN):", value="B0C3VWDT3M")
    category_id = st.text_input("Shopee Category ID Target:", value="102551")

with col_config:
    weight_input = st.number_input("น้ำหนักสินค้า (กรัม):", value=160.0, step=10.0)
    profit_input = st.number_input("อัตรากำไร Profit Margin (%):", value=50.0, step=5.0)
    
    col_l, col_w, col_h = st.columns(3)
    with col_l: len_input = st.number_input("ยาว (CM):", value=17.6)
    with col_w: width_input = st.number_input("กว้าง (CM):", value=16.2)
    with col_h: height_input = st.number_input("สูง (CM):", value=5.2)

st.info(f"💱 **อัตราแลกเปลี่ยน THB/JPY:** TTB = {rate_info['ttb']} | **Exchange Rate (1 THB to JPY) = {rate_info['exchange_rate']}**")

if st.button("🔍 1. ดึงข้อมูลสินค้าจาก Amazon", type="primary"):
    with st.spinner("กำลังดึงข้อมูลภาษาอังกฤษ, รหัส JAN, และขนาดสินค้าจาก Amazon..."):
        data, err_msg = fetch_amazon_data_advanced(asin_code)
        if data:
            st.session_state["fetched_product"] = data
            st.success("ดึงข้อมูลสินค้าสำเร็จ!")
        else:
            st.error(err_msg)

if "fetched_product" in st.session_state:
    pdata = st.session_state["fetched_product"]
    st.markdown("---")
    
    st.subheader(f"📌 {pdata['title']}")
    
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.write(f"🏷️ **JAN Code:** `{pdata['jan_code'] if pdata['jan_code'] else 'N/A'}`")
        st.write(f"📂 **Amazon Category:** {pdata['category_path']}")
    with col_info2:
        st.write(f"📦 **ขนาดกล่อง (กxยxส):** {pdata['length_cm']} x {pdata['width_cm']} x {pdata['height_cm']} cm")
        st.write(f"⚖️ **น้ำหนัก:** {pdata['weight_g']} กรัม")

    st.text_area("รายละเอียดสินค้าภาษาอังกฤษ (English Description)", pdata['description'], height=150)
    
    preview_rows = []
    for v in pdata['variations']:
        sell_p_thb = calculate_net_price_thb(v['cost_jpy'], weight_input, profit_input, rate_info['exchange_rate'])
        preview_rows.append({
            "Child ASIN": v['asin'],
            "JAN/EAN": v['jan'],
            "Option 1": v['opt1'],
            "ทุน (JPY)": v['cost_jpy'],
            "ราคาขาย Shopee (THB)": sell_p_thb
        })
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
    
    st.markdown("---")
    if st.button("🚀 2. อัปโหลดสินค้าเข้าร้าน Shopee ทันที (Direct Publish)", type="secondary"):
        with st.spinner("กำลังแปลงรูปภาพและส่งข้อมูลเข้า Shopee API..."):
            success, msg = publish_to_shopee(
                pdata, category_id, weight_input, profit_input, rate_info['exchange_rate'],
                len_input, width_input, height_input
            )
            if success:
                st.balloons()
                st.success(msg)
            else:
                st.error(msg)
