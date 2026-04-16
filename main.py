import os
import json
import re
import math
import time
import random
import logging
import traceback
import base64
import requests
import pandas as pd
import gspread
from datetime import datetime, timedelta, timezone
from io import BytesIO

from fastapi import FastAPI, Request, BackgroundTasks
from pydantic import BaseModel
from oauth2client.service_account import ServiceAccountCredentials
from pypdf import PdfReader

# --- Selenium / Scraping Imports ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# --- Amazon SP-API Imports ---
try:
    from sp_api.base import SellingApiRequestThrottledException
except ImportError:
    SellingApiRequestThrottledException = Exception

from sp_api.api import CatalogItems, Products, ProductFees
from sp_api.base import Marketplaces

# --- ログ設定 ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- 環境変数の取得とクリーニング ---
def get_clean_env(key, default=None):
    val = os.environ.get(key, default)
    if val:
        return val.strip().strip('"').strip("'").replace('\n', '').replace('\r', '')
    return val

LINE_ACCESS_TOKEN = get_clean_env("LINE_ACCESS_TOKEN")
CHATWORK_TOKEN = get_clean_env("CHATWORK_TOKEN")
OPENAI_API_KEY = get_clean_env("OPENAI_API_KEY")
SHEET_KEY = get_clean_env("SHEET_KEY")
CREDENTIALS_FILE = "credentials.json"

# Yahoo App ID (新規追加)
YAHOO_APP_ID = get_clean_env("YAHOO_APP_ID")

RAKUTEN_APP_ID = get_clean_env("RAKUTEN_APP_ID")
RAKUTEN_API_URL = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601"

LWA_APP_ID = get_clean_env("LWA_APP_ID")
LWA_CLIENT_SECRET = get_clean_env("LWA_CLIENT_SECRET")
REFRESH_TOKEN = get_clean_env("REFRESH_TOKEN")
AWS_ACCESS_KEY = get_clean_env("AWS_ACCESS_KEY")
AWS_SECRET_KEY = get_clean_env("AWS_SECRET_KEY")
ROLE_ARN = get_clean_env("ROLE_ARN", "")
KEEPA_API_KEY = get_clean_env("KEEPA_API_KEY")

app = FastAPI()

# 起動時にキーの状態をログに出力
@app.on_event("startup")
async def startup_event():
    logger.info("=== API Key Check ===")
    logger.info(f"LWA_APP_ID: {LWA_APP_ID[:4]}... (Len: {len(LWA_APP_ID) if LWA_APP_ID else 0})")
    if YAHOO_APP_ID:
        logger.info(f"YAHOO_APP_ID: {YAHOO_APP_ID[:4]}... OK")
    else:
        logger.warning("⚠️ YAHOO_APP_ID is missing!")
        
    if not all([LWA_APP_ID, LWA_CLIENT_SECRET, REFRESH_TOKEN, AWS_ACCESS_KEY, AWS_SECRET_KEY]):
        logger.error("⚠️ CRITICAL: Some Amazon keys are missing!")

def get_jst_time():
    JST = timezone(timedelta(hours=+9), 'JST')
    return datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S')

# Chatwork名前取得
def get_chatwork_name(room_id, account_id):
    if not CHATWORK_TOKEN: return f"ID:{account_id}"
    url = f"https://api.chatwork.com/v2/rooms/{room_id}/members"
    headers = {"X-ChatWorkToken": CHATWORK_TOKEN}
    try:
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            members = res.json()
            for m in members:
                if m.get('account_id') == account_id:
                    return m.get('name')
    except Exception as e:
        logger.error(f"Chatwork Name Error: {e}")
    return f"CWユーザー(ID:{account_id})"

# LINE名前取得
def get_line_user_name(source):
    if not LINE_ACCESS_TOKEN: return "LINEUser"
    user_id = source.get('userId')
    if not user_id: return "LINEUser"
    headers = {"Authorization": f"Bearer {LINE_ACCESS_TOKEN}"}
    source_type = source.get('type')
    url = ""
    if source_type == 'group':
        group_id = source.get('groupId')
        url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
    elif source_type == 'room':
        room_id = source.get('roomId')
        url = f"https://api.line.me/v2/bot/room/{room_id}/member/{user_id}"
    else:
        url = f"https://api.line.me/v2/bot/profile/{user_id}"
    try:
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            profile = res.json()
            return profile.get("displayName", "LINEUser")
    except: pass
    return "LINEUser"

# Chatworkファイルダウンロード
def download_chatwork_file(room_id, file_id):
    if not CHATWORK_TOKEN: return None, None
    url_info = f"https://api.chatwork.com/v2/rooms/{room_id}/files/{file_id}?create_download_url=1"
    headers = {"X-ChatWorkToken": CHATWORK_TOKEN}
    try:
        res_info = requests.get(url_info, headers=headers)
        if res_info.status_code != 200:
            return None, None
        
        data = res_info.json()
        download_url = data.get("download_url")
        filename = data.get("filename")
        
        if download_url:
            res_file = requests.get(download_url)
            if res_file.status_code == 200:
                return res_file.content, filename
    except: pass
    return None, None

def get_spreadsheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials_b64 = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_B64")
    if credentials_b64:
        # Cloud Run: 環境変数から base64 デコードして読み込む
        creds_dict = json.loads(base64.b64decode(credentials_b64).decode('utf-8'))
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        # ローカル開発: ファイルから読み込む
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_KEY)

# ==============================================================================
#  MODULE 1: ChatGPT
# ==============================================================================

def extract_text_from_file(file_bytes, filename):
    text = f"【添付ファイル解析開始: {filename}】\n"
    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
             xls = pd.ExcelFile(BytesIO(file_bytes))
             for sheet_name in xls.sheet_names:
                 df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                 df = df.dropna(how='all').dropna(axis=1, how='all')
                 text += f"\n--- Sheet: {sheet_name} ---\n"
                 text += df.to_csv(index=False, header=False) + "\n"
        elif filename.endswith(".pdf"):
             reader = PdfReader(BytesIO(file_bytes))
             for page in reader.pages:
                 text += page.extract_text() + "\n"
    except Exception as e:
        logger.error(f"File Parse Error: {e}")
    return text

def extract_order_info_gpt(text_content):
    if not OPENAI_API_KEY:
        logger.error("OpenAI API Key is missing.")
        return []
    
    url = "https://api.openai.com/v1/chat/completions"
    
    prompt = f"""
    入力されたテキストから商品情報を抽出しJSON形式で出力してください。
    【重要：出力フォーマット】
    必ずルートキーを "items" とし、その中にリストを作成してください。
    例: {{ "items": [ {{ "jan_code": "...", ... }} ] }}
    
    【データ抽出ルール】
    - jan_code: JANコード (数字のみ)
    - asin: ASINコード
    - product_name: 商品名
    - model_number: 型番
    - cost: 卸金額または単価 (数値)
    - quantity: 数量 (数値)
    - remarks: 備考 (条件、色、サイズなど)
    
    【対象テキスト】
    {text_content[:15000]}
    """
    
    payload = {
        "model": "gpt-4o", 
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0, 
        "response_format": { "type": "json_object" }
    }
    headers = {"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}
    
    try:
        res = requests.post(url, json=payload, headers=headers)
        if res.status_code != 200:
            logger.error(f"OpenAI API Error: {res.text}")
            return []
        content = res.json()['choices'][0]['message']['content']
        data = json.loads(content)
        return data.get("items", [])
    except Exception as e:
        logger.error(f"GPT Logic Error: {e}")
        return []

# ==============================================================================
#  MODULE 2: Amazon Logic (with Keepa)
# ==============================================================================

class SellerNameResolver:
    def __init__(self, keepa_key=None):
        self.keepa_key = keepa_key
        self.file_path = 'sellers.json'
        self.seller_map = self._load_map()
    def _load_map(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f: return json.load(f)
            except: return {}
        return {}
    def _save_map(self):
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f: json.dump(self.seller_map, f, ensure_ascii=False, indent=2)
        except: pass
    def get_name(self, seller_id):
        if not seller_id: return "-"
        if seller_id == 'AN1VRQENFRJN5': return 'Amazon.co.jp'
        if seller_id in self.seller_map: return self.seller_map[seller_id]
        if self.keepa_key:
            try:
                url = f"https://api.keepa.com/seller?key={self.keepa_key}&domain=5&seller={seller_id}"
                res = requests.get(url, timeout=5)
                if res.status_code == 200:
                    data = res.json()
                    if 'sellers' in data and data['sellers']:
                        seller_data = data['sellers'].get(seller_id, {})
                        seller_name = seller_data.get('sellerName')
                        if seller_name:
                            self.seller_map[seller_id] = seller_name
                            self._save_map()
                            return seller_name
            except Exception as e: logger.error(f"Keepa API Error: {e}")
        return seller_id 

class AmazonSearcher:
    def __init__(self):
        self.credentials = {
            'refresh_token': REFRESH_TOKEN, 'lwa_app_id': LWA_APP_ID,
            'lwa_client_secret': LWA_CLIENT_SECRET, 'aws_access_key': AWS_ACCESS_KEY,
            'aws_secret_key': AWS_SECRET_KEY, 'role_arn': ROLE_ARN
        }
        self.marketplace = Marketplaces.JP
        self.mp_id = 'A1VC38T7YXB528'
        self.resolver = SellerNameResolver(keepa_key=KEEPA_API_KEY)
    def log(self, message): logger.info(f"[Amazon] {message}")
    def _call_api_safely(self, func, **kwargs):
        retries = 3
        base_delay = 2.0
        for i in range(retries):
            try: return func(**kwargs)
            except Exception as e:
                error_str = str(e)
                if "429" in error_str or "Throttled" in error_str:
                    wait_time = base_delay * (i + 1) + random.uniform(0.5, 1.5)
                    time.sleep(wait_time)
                else:
                    self.log(f"❌ API Error: {error_str}")
                    if i == retries - 1: return None
        return None
    def calculate_shipping_fee(self, h, l, w):
        try:
            total_size = float(h) + float(l) + float(w)
            if total_size < 60: return 580
            elif total_size <= 80: return 670
            elif total_size <= 100: return 780
            elif total_size <= 120: return 900
            elif total_size <= 140: return 1050
            elif total_size <= 160: return 1300
            else: return 2000
        except: return 0
    def search_by_jan(self, jan_code):
        catalog = CatalogItems(credentials=self.credentials, marketplace=self.marketplace)
        res = self._call_api_safely(catalog.search_catalog_items, keywords=[jan_code], marketplaceIds=[self.mp_id])
        if res and res.payload and 'items' in res.payload:
            items = res.payload['items']
            if items: return items[0].get('asin')
        return None
    def get_product_details_accurate(self, asin):
        result = { "mall": "Amazon", "price": 0, "points_pct": 0, "fee_rate": 0, "shipping": 0, "url": "", "seller": "-", "rank": "-", "category": "-", "order_info": "-", "calc_shipping": 0, "dimensions": "-" }
        catalog = CatalogItems(credentials=self.credentials, marketplace=self.marketplace)
        res_cat = self._call_api_safely(catalog.get_catalog_item, asin=asin, marketplaceIds=[self.mp_id], includedData=['attributes', 'salesRanks', 'summaries'])
        if res_cat and res_cat.payload:
            data = res_cat.payload
            result['url'] = f"https://www.amazon.co.jp/dp/{asin}"
            if 'attributes' in data:
                attrs = data['attributes']
                if 'item_package_dimensions' in attrs:
                    dim = attrs['item_package_dimensions'][0]
                    h = (dim.get('height') or {}).get('value', 0)
                    l = (dim.get('length') or {}).get('value', 0)
                    w = (dim.get('width') or {}).get('value', 0)
                    result['dimensions'] = int(h + l + w)
                    result['calc_shipping'] = self.calculate_shipping_fee(h, l, w)
            if 'salesRanks' in data and data['salesRanks']:
                ranks = data['salesRanks'][0].get('ranks', [])
                if ranks:
                    result['category'] = ranks[0].get('title', '')
                    result['rank'] = f"{ranks[0].get('rank', '-')}位"
        products_api = Products(credentials=self.credentials, marketplace=self.marketplace)
        res_offers = self._call_api_safely(products_api.get_item_offers, asin=asin, MarketplaceId=self.mp_id, item_condition='New')
        if res_offers and res_offers.payload and 'Offers' in res_offers.payload:
            target_offer = None
            for offer in res_offers.payload['Offers']:
                if offer.get('IsBuyBoxWinner', False): target_offer = offer; break
            if not target_offer:
                best_p = float('inf')
                for offer in res_offers.payload['Offers']:
                    p = (offer.get('ListingPrice') or {}).get('Amount', 0)
                    s = (offer.get('Shipping') or {}).get('Amount', 0)
                    if p+s > 0 and p+s < best_p: best_p = p+s; target_offer = offer
            if target_offer:
                p = (target_offer.get('ListingPrice') or {}).get('Amount', 0)
                s = (target_offer.get('Shipping') or {}).get('Amount', 0)
                result['price'] = p + s
                result['shipping'] = s
                pt_data = target_offer.get('Points', {})
                if result['price'] > 0: result['points_pct'] = pt_data.get('PointsNumber', 0) / result['price']
                sid = target_offer.get('SellerId', '')
                result['seller'] = self.resolver.get_name(sid)
        if result['price'] > 0:
            fees_api = ProductFees(credentials=self.credentials, marketplace=self.marketplace)
            res_fee = self._call_api_safely(fees_api.get_product_fees_estimate_for_asin, asin=asin, price=result['price'], is_fba=True, identifier=f'fee-{asin}', currency='JPY', marketplace_id=self.mp_id)
            if res_fee and res_fee.payload:
                fees = res_fee.payload.get('FeesEstimateResult', {}).get('FeesEstimate', {}).get('FeeDetailList', [])
                for fee in fees:
                    if fee.get('FeeType') == 'ReferralFee':
                        amt = (fee.get('FinalFee') or {}).get('Amount', 0)
                        if amt > 0: result['fee_rate'] = amt / result['price']
        return result

# ==============================================================================
#  MODULE 3: Rakuten Logic
# ==============================================================================

def get_rakuten_info(jan_code, cost_price=0):
    result = { "mall": "楽天", "price": 0, "points_pct": 0, "fee_rate": 0.12, "shipping": 0, "shipping_text": "-", "url": "", "seller": "-", "rank": "-", "category": "-", "order_info": "-", "calc_shipping": 0, "dimensions": "-" }
    params = { "applicationId": RAKUTEN_APP_ID, "keyword": jan_code, "sort": "+itemPrice", "hits": 30, "formatVersion": 2, "itemCondition": 1, "ngKeyword": "中古" }
    try:
        response = requests.get(RAKUTEN_API_URL, params=params)
        data = response.json()
        if "Items" in data and len(data["Items"]) > 0:
            for item in data["Items"]:
                item_name = item["itemName"]
                if "中古" in item_name or "USED" in item_name.upper(): continue 
                item_price = item["itemPrice"]
                if cost_price > 0 and item_price < (cost_price * 0.5): continue
                is_free_shipping = (item["postageFlag"] == 0)
                if not is_free_shipping:
                    if "送料無料" in item_name: is_free_shipping = True
                    elif "catchcopy" in item and "送料無料" in item["catchcopy"]: is_free_shipping = True
                result["price"] = item_price
                if is_free_shipping:
                    result["shipping"] = 0
                    result["shipping_text"] = "0"
                else:
                    result["shipping"] = 0 
                    result["shipping_text"] = "送料別"
                result["points_pct"] = item["pointRate"] / 100
                result["seller"] = item["shopName"]
                result["url"] = item["itemUrl"]
                result["order_info"] = f"レビュー: {item['reviewCount']}件"
                break
    except Exception as e: logger.error(f"Rakuten Error: {e}")
    return result

# ==============================================================================
#  MODULE 4: Yahoo Logic (API + Scraping Hybrid)
# ==============================================================================

def get_yahoo_info(jan):
    # 結果格納用辞書の初期化
    result = { 
        "mall": "Yahoo", 
        "price": 0, 
        "points_pct": 0, 
        "fee_rate": 0.10, 
        "shipping": 0, 
        "url": "", 
        "seller": "-", 
        "rank": "-", 
        "category": "-", 
        "order_info": "-", 
        "calc_shipping": 0, 
        "dimensions": "-" 
    }

    # 1. APIによる価格・商品情報の取得
    if not YAHOO_APP_ID:
        logger.error("Yahoo App ID is missing.")
        return result

    api_url = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"
    params = {
        "appid": YAHOO_APP_ID,
        "jan_code": jan,
        "sort": "+price",      # 安い順
        "condition": "new",    # 新品のみ
        "results": 1           # 最安の1件だけ取得
    }

    try:
        res = requests.get(api_url, params=params)
        data = res.json()
        
        hits = data.get("hits", [])
        if not hits:
            return result

        # 最安値商品の情報を抽出
        best_item = hits[0]
        result["price"] = best_item.get("price", 0)
        result["seller"] = best_item.get("seller", {}).get("name", "-")
        result["url"] = best_item.get("url", "")
        
        # ポイント倍率の取得 (APIは倍率を返すため、%になおすために /100 する)
        point_times = best_item.get("point", {}).get("times", 0)
        result["points_pct"] = point_times / 100

        # 送料情報の簡易判定 (APIのshipping.codeやnameから判定)
        shipping_info = best_item.get("shipping", {})
        shipping_code = shipping_info.get("code")
        shipping_name = shipping_info.get("name", "")
        
        if shipping_code == 2 or "無料" in shipping_name or "0円" in shipping_name:
            result["shipping"] = 0
        else:
            result["shipping"] = 0 

    except Exception as e:
        logger.error(f"Yahoo API Error: {e}")
        return result

    # 2. Seleniumによる「注文状況」のスクレイピング (商品が見つかった場合のみ実行)
    if result["url"]:
        driver = None
        try:
            # ブラウザ設定
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36")
            
            service = Service("/usr/bin/chromedriver")
            if not os.path.exists("/usr/bin/chromedriver"): 
                service = Service(ChromeDriverManager().install())
            
            driver = webdriver.Chrome(service=service, options=options)
            
            # APIで取得した商品URLへ直接アクセス
            driver.get(result["url"])
            time.sleep(2) # ページ読み込み待機

            # ページ内のテキストを取得して判定
            try:
                btext = driver.find_element(By.TAG_NAME, "body").text
                found_on_page = ""
                
                if "24時間以内に注文" in btext: 
                    found_on_page = "24時間以内に注文した方がいます"
                elif "3日以内に注文" in btext: 
                    found_on_page = "3日以内に注文した方がいます"
                elif "7日以内に注文" in btext: 
                    found_on_page = "7日以内に注文した方がいます"
                elif "人がカート" in btext or "人が検討" in btext:
                    match = re.search(r"(\d+人が(カート|検討).*?います)", btext)
                    if match: found_on_page = match.group(1)
                
                if found_on_page:
                    result["order_info"] = found_on_page
            except:
                pass # 要素が見つからない場合等は無視

        except Exception as e:
            logger.error(f"Yahoo Scraping Error (Order Info): {e}")
        finally:
            if driver:
                driver.quit()

    return result

# ==============================================================================
#  CORE Logic
# ==============================================================================

def calculate_profit(mall_data, cost):
    price = float(mall_data["price"])
    if price == 0: return 0, 0
    cost = float(cost)
    tax_amt = price - (price / 1.1)
    points_val = price * mall_data["points_pct"]
    fee_val = price * mall_data["fee_rate"]
    
    # ▼▼▼ 修正：全モール共通で、計算された送料(calc_shipping)を使用する ▼▼▼
    shipping_val = mall_data["calc_shipping"]
    
    profit = price - tax_amt - points_val - fee_val - shipping_val - cost
    ex_tax_price = price / 1.1
    margin = profit / ex_tax_price if ex_tax_price > 0 else 0
    return int(profit), margin

def _run_analysis_for_item(amz_searcher, sheet2, item):
    """1商品分の分析を実行してAnalysisシートに書き込む共通処理"""
    jan = item.get("jan_code")
    if not jan:
        return
    try:
        cost = float(item.get("cost", 0))
    except:
        cost = 0
    name = item.get("product_name", "Unknown")
    sender = item.get("sender_name", "-")

    amz_data = {"mall": "Amazon", "price": 0, "points_pct": 0, "fee_rate": 0, "shipping": 0, "url": "", "seller": "-", "rank": "-", "category": "-", "order_info": "-", "calc_shipping": 0, "dimensions": "-"}
    asin = amz_searcher.search_by_jan(jan)
    if asin:
        amz_data = amz_searcher.get_product_details_accurate(asin)

    rak_data = get_rakuten_info(jan, cost)
    yah_data = get_yahoo_info(jan)

    estimated_shipping_fee = amz_data.get('calc_shipping', 0)
    item_dimensions = amz_data.get('dimensions', '-')
    rak_data['calc_shipping'] = estimated_shipping_fee
    rak_data['dimensions'] = item_dimensions
    yah_data['calc_shipping'] = estimated_shipping_fee
    yah_data['dimensions'] = item_dimensions

    for data in [amz_data, rak_data, yah_data]:
        profit, margin = calculate_profit(data, cost)
        shipping_display = data.get("shipping_text", data["shipping"])
        if data['mall'] == 'Amazon' and data['shipping'] == 0:
            shipping_display = 0
        elif data['mall'] == 'Yahoo':
            shipping_display = data['shipping']
        row = [
            sender, jan, name, data['mall'], data['price'], cost, data['seller'],
            shipping_display,
            data['category'] if data['mall'] == 'Amazon' else '-',
            data['rank'] if data['mall'] == 'Amazon' else '-',
            f"{data['fee_rate']:.0%}", f"{data['points_pct']:.1%}",
            data['dimensions'], data['calc_shipping'], profit, f"{margin:.1%}",
            data['order_info'], data['url']
        ]
        sheet2.append_row(row)


def process_analysis(input_data):
    logger.info("Starting Analysis Task...")
    try:
        client = get_spreadsheet()
        try:
            sheet2 = client.worksheet("Analysis")
        except:
            sheet2 = client.add_worksheet("Analysis", 1000, 20)
            sheet2.append_row(["送信者名", "JAN", "商品名", "モール", "価格", "仕入原価", "店舗名", "送料", "amz分類", "amzランク", "手数料%", "ポイント率", "３辺合計", "送料目安", "利益", "利益率", "備考（yh注文状況、Rレビュー件数）", "URL"])

        amz_searcher = AmazonSearcher()
        for item in input_data:
            _run_analysis_for_item(amz_searcher, sheet2, item)
        logger.info("Analysis Task Completed.")
    except Exception as e:
        logger.error(f"Analysis Error: {e}")
        logger.error(traceback.format_exc())


# ==============================================================================
#  MODULE 5: Spreadsheet Direct Input
# ==============================================================================

INPUT_SHEET_NAME = "入力"
INPUT_SHEET_HEADERS = ["JANコード", "商品名", "数量", "仕入れ価格", "ステータス", "処理日時"]

# 入力シートの列インデックス（1始まり）
COL_JAN = 1
COL_NAME = 2
COL_QTY = 3
COL_COST = 4
COL_STATUS = 5
COL_PROCESSED_AT = 6


def ensure_input_sheet(client):
    """入力シートがなければ作成してヘッダーを書き込む"""
    try:
        sheet = client.worksheet(INPUT_SHEET_NAME)
    except:
        sheet = client.add_worksheet(INPUT_SHEET_NAME, 1000, 6)
        sheet.append_row(INPUT_SHEET_HEADERS)
    return sheet


def process_spreadsheet_input(pending_items):
    """入力シートから読み込んだ商品を分析し、ステータスを更新する"""
    logger.info(f"Spreadsheet Input Task: {len(pending_items)} items")
    try:
        client = get_spreadsheet()
        input_sheet = client.worksheet(INPUT_SHEET_NAME)

        try:
            analysis_sheet = client.worksheet("Analysis")
        except:
            analysis_sheet = client.add_worksheet("Analysis", 1000, 20)
            analysis_sheet.append_row(["送信者名", "JAN", "商品名", "モール", "価格", "仕入原価", "店舗名", "送料", "amz分類", "amzランク", "手数料%", "ポイント率", "３辺合計", "送料目安", "利益", "利益率", "備考（yh注文状況、Rレビュー件数）", "URL"])

        amz_searcher = AmazonSearcher()
        now = get_jst_time()

        for entry in pending_items:
            row_index = entry["row_index"]
            item = entry["item"]

            # ステータスを「処理中」に更新
            input_sheet.update_cell(row_index, COL_STATUS, "処理中")

            try:
                _run_analysis_for_item(amz_searcher, analysis_sheet, item)
                input_sheet.update_cell(row_index, COL_STATUS, "完了")
                input_sheet.update_cell(row_index, COL_PROCESSED_AT, now)
            except Exception as e:
                logger.error(f"Item analysis failed (row {row_index}): {e}")
                input_sheet.update_cell(row_index, COL_STATUS, "エラー")
                input_sheet.update_cell(row_index, COL_PROCESSED_AT, now)

        logger.info("Spreadsheet Input Task Completed.")
    except Exception as e:
        logger.error(f"Spreadsheet Input Error: {e}")
        logger.error(traceback.format_exc())

@app.get("/")
def health():
    return {"status": "running"}


class SpreadsheetItem(BaseModel):
    jan_code: str
    product_name: str = "Unknown"
    quantity: str = ""
    cost: float = 0
    sender_name: str = "スプレッドシート入力"


class SpreadsheetPayload(BaseModel):
    items: list[SpreadsheetItem]


def process_direct_items(items_data: list[dict]):
    """GASから直接送信されたアイテムリストを分析する"""
    logger.info(f"Direct Items Task: {len(items_data)} items")
    try:
        client = get_spreadsheet()
        try:
            analysis_sheet = client.worksheet("Analysis")
        except:
            analysis_sheet = client.add_worksheet("Analysis", 1000, 20)
            analysis_sheet.append_row(["送信者名", "JAN", "商品名", "モール", "価格", "仕入原価", "店舗名", "送料", "amz分類", "amzランク", "手数料%", "ポイント率", "３辺合計", "送料目安", "利益", "利益率", "備考（yh注文状況、Rレビュー件数）", "URL"])

        amz_searcher = AmazonSearcher()
        for item in items_data:
            _run_analysis_for_item(amz_searcher, analysis_sheet, item)
        logger.info("Direct Items Task Completed.")
    except Exception as e:
        logger.error(f"Direct Items Error: {e}")
        logger.error(traceback.format_exc())


@app.post("/trigger/spreadsheet")
async def trigger_spreadsheet(payload: SpreadsheetPayload, background_tasks: BackgroundTasks):
    """
    GASが現在のシートのデータをJSONで送信 → バックグラウンドで分析を実行。
    列構成: A=JANコード, B=商品名, C=数量/在庫, D=仕入れ価格/下代
    """
    if not payload.items:
        return {"status": "ok", "queued": 0, "message": "処理対象の行がありません"}

    items_data = [item.model_dump() for item in payload.items]
    background_tasks.add_task(process_direct_items, items_data)
    logger.info(f"/trigger/spreadsheet: {len(items_data)} items queued.")
    return {"status": "ok", "queued": len(items_data), "message": f"{len(items_data)}件の分析を開始しました"}

@app.post("/webhook/line")
async def line_webhook(request: Request, background_tasks: BackgroundTasks):
    try:
        body = await request.body()
        data = json.loads(body.decode('utf-8'))
        events = data.get('events', [])
        for event in events:
            if event.get('type') != 'message': continue
            msg = event.get('message', {})
            source = event.get('source', {})
            sender_name = get_line_user_name(source)
            if msg.get('type') == 'text':
                items = extract_order_info_gpt(msg.get('text', ''))
                if items:
                    client = get_spreadsheet()
                    sheet1 = client.sheet1
                    for it in items:
                        it['sender_name'] = sender_name
                        sheet1.append_row(["LINE", sender_name, it.get("jan_code"), it.get("asin"), it.get("product_name"), "", it.get("cost"), it.get("quantity"), it.get("remarks"), get_jst_time()])
                    background_tasks.add_task(process_analysis, items)
    except: pass
    return "OK"

@app.post("/webhook/chatwork")
async def chatwork_webhook(request: Request, background_tasks: BackgroundTasks):
    try:
        body = await request.body()
        data = json.loads(body.decode('utf-8'))
        
        webhook_event = data.get("webhook_event", {})
        msg_body = webhook_event.get("body", "")
        room_id = webhook_event.get("room_id")
        
        file_matches = re.findall(r'\[download:(\d+)\]', msg_body)
        if file_matches:
            for file_id in file_matches:
                file_data, filename = download_chatwork_file(room_id, file_id)
                if file_data:
                    extracted_text = extract_text_from_file(file_data, filename)
                    msg_body += f"\n\n【添付ファイル内容: {filename}】\n{extracted_text}"

        account_id = webhook_event.get("account_id")
        sender_name = get_chatwork_name(room_id, account_id)
        items = extract_order_info_gpt(msg_body)
        
        if items:
            client = get_spreadsheet()
            sheet1 = client.sheet1
            for it in items:
                it['sender_name'] = sender_name
                sheet1.append_row(["Chatwork", sender_name, it.get("jan_code"), it.get("asin"), it.get("product_name"), "", it.get("cost"), it.get("quantity"), it.get("remarks"), get_jst_time()])
            background_tasks.add_task(process_analysis, items)
            
    except Exception as e:
        logger.error(f"CW Error: {e}")
    return "OK"

@app.post("/webhook/email")
async def email_webhook(request: Request, background_tasks: BackgroundTasks):
    try:
        body = await request.body()
        data = json.loads(body.decode('utf-8'))
        sender = data.get("sender", "EmailUser")
        subject = data.get("subject", "")
        email_body = data.get("body", "")
        full_text = f"件名: {subject}\n\n{email_body}"
        if email_body:
            items = extract_order_info_gpt(full_text)
            if items:
                client = get_spreadsheet()
                sheet1 = client.sheet1
                for it in items:
                    it['sender_name'] = sender
                    sheet1.append_row(["Email", sender, it.get("jan_code", ""), it.get("asin", ""), it.get("product_name", ""), "", it.get("cost", ""), it.get("quantity", ""), it.get("remarks", ""), get_jst_time()])
                background_tasks.add_task(process_analysis, items)
    except Exception as e:
        logger.error(f"Email Webhook Error: {e}")
    return "OK"