"""
SPAO 베스트 상품 크롤러
- 여성 베스트: https://spao.com/product/list.html?cate_no=6833
- 남성 베스트: https://spao.com/product/list.html?cate_no=6834
- requests + BeautifulSoup 기반 (Selenium 불필요)
- 결과: spao_history.json 저장
"""

import requests
from bs4 import BeautifulSoup
import json
import os
import re
from datetime import datetime

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(WORK_DIR, 'spao_history.json')

CATEGORIES = {
    '여성': {'cate_no': 6833, 'pages': 2},
    '남성': {'cate_no': 6834, 'pages': 2},
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Referer': 'https://spao.com/',
}


def clean_product_name(raw_name):
    """상품명 정리: 코드 제거
    '[리사이클] 베이직 퍼플리스 집업 (SPFZF4TC01 RE)_(W)SPFZG11C01'
    -> '[리사이클] 베이직 퍼플리스 집업'
    """
    name = raw_name.strip()
    # _SP코드 또는 _(W)SP코드 제거
    name = re.sub(r'_\(?[WM]?\)?SP\w+$', '', name).strip()
    # (SP코드 RE) 제거
    name = re.sub(r'\s*\(SP\w+(?:\s+RE)?\)\s*', ' ', name).strip()
    # 끝 _ 제거
    name = name.rstrip('_').strip()
    return name


def scrape_spao_category(category_name, cate_no, max_pages=2):
    """특정 카테고리의 베스트 상품 크롤링"""
    products = []
    seen_names = set()

    for page in range(1, max_pages + 1):
        url = f'https://spao.com/product/list.html?cate_no={cate_no}&page={page}'

        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            resp.encoding = 'utf-8'
        except requests.RequestException as e:
            print(f"  [오류] {category_name} 페이지 {page} 요청 실패: {e}")
            continue

        soup = BeautifulSoup(resp.text, 'html.parser')

        # 메인 상품 리스트: ul.prdList.grid3list (Cafe24 쇼핑몰)
        main_list = soup.select_one('ul.prdList.grid3list')
        if not main_list:
            print(f"  [경고] {category_name} 페이지 {page}: 상품 목록 없음")
            continue

        items = main_list.select(':scope > li')
        page_count = 0

        for item in items:
            try:
                # 상품명: .description .name a span
                name_el = item.select_one('.description .name a span')
                if not name_el:
                    continue

                raw_name = name_el.get_text(strip=True)
                if not raw_name or len(raw_name) < 3:
                    continue

                clean_name = clean_product_name(raw_name)

                # 중복 체크
                if clean_name in seen_names:
                    continue
                seen_names.add(clean_name)

                # 판매가: span.price
                price_el = item.select_one('.price_box .price')
                sale_price = 0
                if price_el:
                    price_text = price_el.get_text(strip=True)
                    digits = re.sub(r'[^\d]', '', price_text)
                    sale_price = int(digits) if digits else 0

                # 정가: span.custom
                original_price = sale_price
                custom_el = item.select_one('.price_box .custom')
                if custom_el:
                    orig_text = custom_el.get_text(strip=True)
                    digits = re.sub(r'[^\d]', '', orig_text)
                    if digits:
                        original_price = int(digits)

                # 이미지
                img_el = item.select_one('.prdImg img') or item.select_one('img[src*="cafe24img"]')
                img_url = ''
                if img_el:
                    img_url = img_el.get('src') or img_el.get('data-original') or \
                              img_el.get('ec-data-src') or ''

                # 리뷰 수
                review_count = 0
                review_el = item.select_one('.review_count') or item.select_one('[class*="review"]')
                if review_el:
                    rv_text = review_el.get_text(strip=True)
                    rv_match = re.search(r'[\d,]+', rv_text)
                    if rv_match:
                        review_count = int(rv_match.group().replace(',', ''))

                rank = len(products) + 1
                products.append({
                    'rank': rank,
                    'name': clean_name,
                    'price': sale_price,
                    'original_price': original_price,
                    'image_url': img_url,
                    'review_count': review_count,
                })
                page_count += 1

            except Exception as e:
                print(f"  [오류] 상품 파싱: {e}")
                continue

        print(f"  {category_name} 페이지 {page}: {page_count}개 추출 (누적: {len(products)}개)")

    return products


def scrape_all():
    """전체 SPAO 베스트 크롤링"""
    print("=" * 50)
    print("SPAO 베스트 상품 크롤링 시작")
    print("=" * 50)

    today = datetime.now().strftime('%Y%m%d')
    results = {}

    for cat_name, config in CATEGORIES.items():
        print(f"\n[{cat_name}] 크롤링 중...")
        products = scrape_spao_category(cat_name, config['cate_no'], config['pages'])

        if products:
            cat_key = f"스파오_{cat_name}"
            items_dict = {}
            for p in products:
                items_dict[p['name']] = {
                    'rank': p['rank'],
                    'price': p['price'],
                    'original_price': p['original_price'],
                    'image_url': p.get('image_url', ''),
                    'review_count': p.get('review_count', 0),
                }
            results[cat_key] = {today: items_dict}
            print(f"  -> {cat_name}: {len(products)}개 수집 완료")
        else:
            print(f"  -> {cat_name}: 수집 실패")

    # JSON 저장 (히스토리 누적)
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
    print(f"\n{'=' * 50}")
    print(f"완료! 총 {total}개 상품 수집 -> spao_history.json 저장")
    print(f"{'=' * 50}")

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
