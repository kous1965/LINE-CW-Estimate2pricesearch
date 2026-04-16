import time
import re
import math
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- ログ設定 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36")
    
    # Cloud Shell / Cloud Run 環境用のドライバ設定
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

def test_yahoo_scraping(jan):
    driver = init_driver()
    wait = WebDriverWait(driver, 30)
    logger.info(f"🚀 テスト開始 JAN: {jan}")

    try:
        # 1. 検索
        url = f"https://shopping.yahoo.co.jp/search?first=1&tab_ex=commerce&fr=shp-prop&p={jan}"
        logger.info(f"URLへ移動: {url}")
        driver.get(url)
        time.sleep(3)

        # 2. 製品ページ遷移 (元コードのロジック)
        try:
            logger.info("製品ページ(カタログ)へのリンクを探しています...")
            product_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/products/') and not(contains(@href, 'search'))]")))
            target_url = product_link.get_attribute("href")
            logger.info(f"製品ページ発見: {target_url}")
            driver.get(target_url)
            time.sleep(5)
        except Exception as e:
            logger.warning(f"⚠️ 製品ページが見つかりませんでした。検索結果ページで続行します。エラー: {e}")

        # 3. リスト表示へ切り替え
        try:
            list_view_btn = driver.find_elements(By.XPATH, "//li[contains(@class, 'ChangeView__item--list')]//a | //a[contains(text(), 'リスト')]")
            if list_view_btn:
                logger.info("リスト表示に切り替えます")
                driver.execute_script("arguments[0].click();", list_view_btn[0])
                time.sleep(3)
        except: pass

        # 4. 「送料込み」へ切り替え（元コードのJSロジック完全再現）
        logger.info("「送料込み」フィルタを適用します...")
        switched = False
        for attempt in range(3):
            try:
                body_text = driver.find_element(By.TAG_NAME, "body").text
                if "表示価格：送料込みの価格" in body_text or "条件指定：送料込み" in body_text:
                    switched = True
                    break

                # JSでボタンクリック
                driver.execute_script("""
                    let buttons = document.querySelectorAll('button, div[role="button"], span');
                    for(let b of buttons){
                        if(b.innerText.includes('表示価格') || b.innerText.includes('実質価格') || b.innerText.includes('本体価格')) { b.click(); }
                    }
                """)
                time.sleep(1)
                driver.execute_script("""
                    let options = document.querySelectorAll('li, a, button, label');
                    for (let o of options) {
                        if (o.innerText.includes("送料込みの価格")) { o.click(); }
                    }
                """)
                
                time.sleep(3)
                
                # 確認
                curr_text = driver.find_element(By.TAG_NAME, "body").text
                if "送料込みの価格" in curr_text:
                    switched = True
                    logger.info("✅ 送料込みへの切り替え成功")
                    break
            except:
                driver.refresh()
                time.sleep(3)
        
        if not switched:
            logger.warning("⚠️ 送料込みへの切り替えに失敗した可能性があります")

        # 5. データ取得
        items = driver.find_elements(By.XPATH, "//li[contains(@class, 'elItem')]")
        if not items: items = driver.find_elements(By.XPATH, "//div[contains(@class, 'LoopList__item')]")
        if not items: items = driver.find_elements(By.XPATH, "//div[contains(@class, 'SearchResultItem')]")
        
        logger.info(f"取得した商品要素数: {len(items)}")
        
        for i, item in enumerate(items[:3]): # 上位3件だけテスト
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", item)
                raw_text = item.text
                clean_text = re.sub(r'\s+', ' ', raw_text)
                
                # 価格
                price = "不明"
                try:
                    pe = item.find_element(By.XPATH, ".//span[contains(@class, 'elPriceValue')]")
                    price = pe.text
                except:
                    pm = re.search(r'([0-9,]+)\s*円', clean_text)
                    if pm: price = pm.group(1)

                # 店名
                shop = "不明"
                try:
                    se = item.find_element(By.XPATH, ".//*[contains(@class, 'Store') or contains(@class, 'store')]//a")
                    shop = se.text
                except: pass

                # 送料
                postage = "不明"
                if "送料無料" in clean_text or "送料0円" in clean_text:
                    postage = "無料"
                else:
                    pm = re.search(r'送料([0-9,]+)円', clean_text)
                    if pm: postage = f"{pm.group(1)}円"

                print("-" * 30)
                print(f"ランク: {i+1}位")
                print(f"店舗名: {shop}")
                print(f"価格　: {price}")
                print(f"送料　: {postage}")
                
            except Exception as e:
                print(f"解析エラー: {e}")

    except Exception as e:
        logger.error(f"全体エラー: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    # 指定のJANコードでテスト
    test_yahoo_scraping("4906128524236")