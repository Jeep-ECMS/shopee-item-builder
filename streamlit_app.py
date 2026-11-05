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
    """ ดาวน์โหลดรูปจาก Amazon พร้อม Header แล้วอัปโหลดเข้า Shopee """
    path = "/api/v2/product/upload_image"
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        img_res = requests.get(image_url, headers=headers, timeout=15)
        
        if img_res.status_code != 200:
            return None, f"HTTP Error {img_res.status_code} ตอนดาวน์โหลดรูปภาพจาก Amazon"
            
        img_bytes = img_res.content
    except Exception as e:
        return None, f"ดาวน์โหลดรูปไม่สำเร็จ: {e}"

    timestamp = int(time.time())
    sign = generate_shopee_sign(path, timestamp, SHOPEE_PARTNER_ID, SHOPEE_PARTNER_KEY, SHOPEE_ACCESS_TOKEN, SHOPEE_SHOP_ID)
    
    url = f"https://partner.shopeemobile.com{path}?partner_id={SHOPEE_PARTNER_ID}&timestamp={timestamp}&access_token={SHOPEE_ACCESS_TOKEN}&shop_id={SHOPEE_SHOP_ID}&sign={sign}"
    
    files = {'image': ('image.jpg', img_bytes, 'image/jpeg')}
    
    try:
        res = requests.post(url, files=files, timeout=30).json()
        if res.get("error") == "":
            return res.get("response", {}).get("image_info", {}).get("image_id"), None
        else:
            return None, f"Shopee Error: {res.get('message', res.get('error'))}"
    except Exception as e:
        return None, f"API Connection Fail: {e}"

# --- 3. PRICE & EXCHANGE RATE CALCULATION FORMULA ---
@st.cache_data(ttl=3600)
def fetch_base_rates():
    rates_out = {"THB": 4.73, "PHP": 2.65}
    try:
        req = urllib.request.Request("https://open.er-api.com/v6/latest/PHP", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            rates_out["PHP"] = round(data
