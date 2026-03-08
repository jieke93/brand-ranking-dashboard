"""
SPAO 베스트 상품 크롤러
- 여성 베스트: https://spao.com/product/list.html?cate_no=6833
- 남성 베스트: https://spao.com/product/list.html?cate_no=6834
- Selenium 스크린샷 캡처 방식 (다른 크롤러와 동일)
- 결과: spao_history.json + product_images_hd/ + image_archive/ 저장
"""

import json
import os
import re
import io
import time
import hashlib
import base64
from datetime import datetime
from PIL import Image as PILImage

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(WORK_DIR, 'spao_history.json')
LOG_FILE = os.path.join(WORK_DIR, 'spao_crawler_log.txt')

# 이미지 설정
IMG_WIDTH = 80
IMG_HEIGHT = 107
HD_IMG_WIDTH = 400   # 대시보드용 고해상도 이미지
HD_IMG_HEIGHT = 534  # 3:4 비율

CATEGORIES = {
    '여성': {'cate_no': 6833, 'pages': 2},
    '남성': {'cate_no': 6834, 'pages': 2},
}


def log(msg):
    print(msg)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except:
        pass


def clean_product_name(raw_name):
    """상품명 정리: 코드 제거"""
    name = raw_name.strip()
    name = re.sub(r'_\(?[WM]?\)?SP\w+$', '', name).strip()
    name = re.sub(r'\s*\(SP\w+(?:\s+RE)?\)\s*', ' ', name).strip()
    name = name.rstrip('_').strip()
    return name


def setup_driver():
    """Chrome 드라이버 초기화"""
    log("[1/4] Chrome 드라이버 초기화")
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--disable-extensions')
    options.add_argument('--remote-debugging-port=0')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36')
    options.page_load_strategy = 'eager'

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
    })
    driver.set_page_load_timeout(60)
    log("  Chrome 초기화 완료!")
    return driver


def capture_image_from_element(element):
    """요소를 스크린샷 캡처하여 이미지로 반환
    반환: (excel_img_bytes, hd_img_bytes) 튜플 또는 None"""
    try:
        png_data = element.screenshot_as_png
        if not png_data:
            return None
        img = PILImage.open(io.BytesIO(png_data))
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')

        hd_img = img.resize((HD_IMG_WIDTH, HD_IMG_HEIGHT), PILImage.Resampling.LANCZOS)
        hd_bytes = io.BytesIO()
        hd_img.save(hd_bytes, format='JPEG', quality=92)
        hd_bytes.seek(0)

        xl_img = img.resize((IMG_WIDTH, IMG_HEIGHT), PILImage.Resampling.LANCZOS)
        xl_bytes = io.BytesIO()
        xl_img.save(xl_bytes, format='JPEG', quality=85)
        xl_bytes.seek(0)

        return (xl_bytes, hd_bytes)
    except Exception as e:
        return None


def scrape_spao_category(driver, category_name, cate_no, max_pages=2):
    """Selenium으로 SPAO 카테고리 크롤링 + 이미지 스크린샷 캡처"""
    products = []
    seen_names = set()
    img_success = 0

    for page in range(1, max_pages + 1):
        url = f'https://spao.com/product/list.html?cate_no={cate_no}&page={page}'
        log(f"  [{category_name}] 페이지 {page} 로딩: {url}")

        try:
            driver.get(url)
            time.sleep(3)  # 페이지 로딩 대기
        except Exception as e:
            log(f"  [오류] 페이지 로딩 실패: {e}")
            continue

        # 스크롤하여 이미지 lazy-load 트리거
        for i in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.8)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)

        # 상품 리스트 파싱
        items = driver.find_elements(By.CSS_SELECTOR, 'ul.prdList.grid3list > li')
        if not items:
            log(f"  [경고] {category_name} 페이지 {page}: 상품 없음")
            continue

        page_count = 0
        for item in items:
            try:
                # 상품명
                name_el = item.find_elements(By.CSS_SELECTOR, '.description .name a span')
                if not name_el:
                    continue
                raw_name = name_el[0].text.strip()
                if not raw_name or len(raw_name) < 3:
                    continue
                clean_name = clean_product_name(raw_name)
                if clean_name in seen_names:
                    continue
                seen_names.add(clean_name)

                # 판매가
                sale_price = 0
                price_els = item.find_elements(By.CSS_SELECTOR, '.price_box .price')
                if price_els:
                    price_text = price_els[0].text.strip()
                    digits = re.sub(r'[^\d]', '', price_text)
                    sale_price = int(digits) if digits else 0

                # 정가
                original_price = sale_price
                custom_els = item.find_elements(By.CSS_SELECTOR, '.price_box .custom')
                if custom_els:
                    orig_text = custom_els[0].text.strip()
                    digits = re.sub(r'[^\d]', '', orig_text)
                    if digits:
                        original_price = int(digits)

                # 이미지 스크린샷 캡처
                rank = len(products) + 1
                hd_image_data = None
                xl_image_data = None

                try:
                    img_elem = item.find_element(By.CSS_SELECTOR, '.prdImg img')
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", img_elem)
                    time.sleep(0.2)
                    img_data = capture_image_from_element(img_elem)
                    if img_data:
                        xl_image_data = img_data[0]
                        hd_image_data = img_data[1]
                        img_success += 1
                except:
                    pass

                products.append({
                    'rank': rank,
                    'name': clean_name,
                    'price': sale_price,
                    'original_price': original_price,
                    'image_data': xl_image_data,
                    'hd_image_data': hd_image_data,
                    'review_count': 0,
                })
                page_count += 1

                if rank % 5 == 0:
                    log(f"      -> {rank}번째 상품까지 처리 (이미지: {img_success}개)")

            except Exception as e:
                log(f"  [오류] 상품 파싱: {str(e)[:60]}")
                continue

        log(f"  {category_name} 페이지 {page}: {page_count}개 추출 (누적: {len(products)}개, 이미지: {img_success}개)")

    return products


def _save_hd_images(all_data, brand_name):
    """대시보드용 고해상도 이미지 저장 (product_images_hd/ + image_archive/)"""
    hd_dir = os.path.join(WORK_DIR, 'product_images_hd')
    archive_dir = os.path.join(WORK_DIR, 'image_archive')
    os.makedirs(hd_dir, exist_ok=True)
    os.makedirs(archive_dir, exist_ok=True)

    file_hash = hashlib.md5(
        f"spao_{datetime.now().strftime('%Y%m%d')}".encode()
    ).hexdigest()[:8]

    saved_hd = 0
    saved_archive = 0

    for sheet_name, products in all_data.items():
        for p in products:
            hd_data = p.get('hd_image_data')
            if hd_data:
                name = p.get('name', '')[:20]
                safe = re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')

                # product_images_hd/ 저장 (dashboard에서 key 매칭용)
                fname = f"{file_hash}_{brand_name}_{sheet_name}_{p['rank']}_{safe}.jpg"
                fpath = os.path.join(hd_dir, fname)
                if not os.path.exists(fpath):
                    try:
                        hd_data.seek(0)
                        with open(fpath, 'wb') as f:
                            f.write(hd_data.read())
                        saved_hd += 1
                    except:
                        pass

                # image_archive/ 저장 (상품명 기반 영구 보관)
                safe_archive = re.sub(r'[\\/:*?"<>|]', '_', f"{brand_name}_{p['name']}")
                archive_path = os.path.join(archive_dir, f"{safe_archive}.jpg")
                if not os.path.exists(archive_path):
                    try:
                        hd_data.seek(0)
                        with open(archive_path, 'wb') as f:
                            f.write(hd_data.read())
                        saved_archive += 1
                    except:
                        pass

    log(f"  [HD] product_images_hd/에 {saved_hd}개 저장")
    log(f"  [Archive] image_archive/에 {saved_archive}개 저장")


def scrape_all():
    """전체 SPAO 베스트 크롤링 (Selenium 스크린샷 캡처)"""
    # 로그 초기화
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write("")
    except:
        pass

    log("=" * 60)
    log("  SPAO 베스트 상품 크롤러 (스크린샷 캡처)")
    log("=" * 60)

    today = datetime.now().strftime('%Y%m%d')
    results = {}
    all_data_for_images = {}

    driver = setup_driver()

    try:
        log("\n[2/4] 데이터 수집 시작")
        for cat_name, config in CATEGORIES.items():
            log(f"\n  === {cat_name} ===")
            products = scrape_spao_category(driver, cat_name, config['cate_no'], config['pages'])

            if products:
                cat_key = f"스파오_{cat_name}"
                items_dict = {}
                for p in products:
                    items_dict[p['name']] = {
                        'rank': p['rank'],
                        'price': p['price'],
                        'original_price': p['original_price'],
                        'image_url': '',  # 캡처 방식으로 변경, URL 불필요
                        'review_count': p.get('review_count', 0),
                    }
                results[cat_key] = {today: items_dict}
                all_data_for_images[cat_name] = products
                log(f"  -> {cat_name}: {len(products)}개 수집 완료")
            else:
                log(f"  -> {cat_name}: 수집 실패")
    finally:
        driver.quit()
        log("  Chrome 종료")

    # [3/4] 이미지 저장
    log("\n[3/4] 이미지 저장")
    if all_data_for_images:
        _save_hd_images(all_data_for_images, '스파오')

    # [4/4] JSON 저장
    log("\n[4/4] JSON 저장")
    history = {}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except (json.JSONDecodeError, Exception):
            history = {}

    for cat_key, dates_data in results.items():
        if cat_key not in history:
            history[cat_key] = {}
        for date_key, items in dates_data.items():
            history[cat_key][date_key] = items

    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    total = sum(len(v) for dates in results.values() for v in dates.values())
    log(f"\n{'=' * 60}")
    log(f"완료! 총 {total}개 상품 수집 -> spao_history.json 저장")
    log(f"{'=' * 60}")

    return results


def load_spao_data():
    """저장된 SPAO 데이터 로드 (최신 날짜만)"""
    if not os.path.exists(HISTORY_FILE):
        return {}

    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            history = json.load(f)
    except (json.JSONDecodeError, Exception):
        return {}

    latest_data = {}
    for cat_key, dates_data in history.items():
        if not dates_data:
            continue
        latest_date = max(dates_data.keys())
        latest_data[cat_key] = dates_data[latest_date]

    return latest_data


if __name__ == '__main__':
    scrape_all()
