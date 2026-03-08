"""미쏘 페이지 HTML 구조 진단"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

options = Options()
options.add_argument('--headless=new')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--window-size=1920,1080')
options.page_load_strategy = 'eager'

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.set_page_load_timeout(30)

try:
    driver.get("https://mixxo.com/product/list.html?cate_no=45")
    time.sleep(5)
    
    # 스크롤
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    
    # 상품 리스트 HTML 추출
    items = driver.find_elements(By.CSS_SELECTOR, "ul.prdList > li")
    print(f"Found {len(items)} items")
    
    if items:
        # 첫 3개 아이템의 HTML 출력
        for i, item in enumerate(items[:3]):
            html = item.get_attribute('outerHTML')
            print(f"\n{'='*60}")
            print(f"ITEM {i+1}:")
            print(f"{'='*60}")
            print(html[:2000])
            
        # 상품명 추출 테스트
        print(f"\n{'='*60}")
        print("NAME EXTRACTION TEST:")
        print(f"{'='*60}")
        for i, item in enumerate(items[:5]):
            text = item.text.strip()
            print(f"\nItem {i+1} full text:")
            print(text[:300])
            print("---")
            
            # 다양한 셀렉터 테스트
            for sel in [".description .name a", ".name a", "p.name a", 
                        ".prd-name a", ".description a", ".name span",
                        ".name", ".description .name", "a[href*='/product/']"]:
                try:
                    elem = item.find_element(By.CSS_SELECTOR, sel)
                    t = elem.text.strip()
                    if t:
                        print(f"  [{sel}] = {t[:60]}")
                except:
                    pass

finally:
    driver.quit()
