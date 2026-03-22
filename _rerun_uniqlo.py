# -*- coding: utf-8 -*-
"""
유니클로 WOMEN/MEN만 재크롤링하여 기존 KIDS/BABY 데이터와 합치는 스크립트
"""
import sys
import os
import time
import glob
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 기존 크롤러의 모든 함수를 가져옴
from uniqlo_ranking_v5 import (
    setup_driver, scrape_category_with_tabs, create_excel, 
    close_cookie_popup, safe_get, log, CATEGORIES,
    LOG_FILE
)
import openpyxl

def load_existing_data_from_excel(excel_path):
    """기존 엑셀에서 KIDS/BABY 데이터를 읽어옴 (이미지 제외)"""
    from uniqlo_ranking_v5 import extract_products, classify_products
    # 엑셀에서는 이미지 바이너리 복구가 어려우므로
    # 기존 크롤링 결과를 재사용하기 위해 이 함수는 사용하지 않음
    pass

def main():
    # 로그 파일 초기화
    with open("uniqlo_rerun_log.txt", 'w', encoding='utf-8') as f:
        f.write("")
    
    print("=" * 60)
    print("  유니클로 WOMEN/MEN 재크롤링 (수정된 탭 전환 로직)")
    print("=" * 60)
    
    driver = setup_driver()
    all_data = {}
    
    try:
        print("\n[STEP 1] WOMEN/MEN 재크롤링")
        print("=" * 60)
        
        # WOMEN, MEN만 크롤링
        for category in ['WOMEN', 'MEN']:
            info = CATEGORIES[category]
            data = scrape_category_with_tabs(driver, category, info['url'], info['tabs'])
            all_data.update(data)
            
            # 각 카테고리 후 통계
            cat_total = sum(len(v) for k, v in data.items())
            print(f"\n  => {category}: {len(data)}개 시트, {cat_total}개 상품")
        
    finally:
        driver.quit()
        print("\n브라우저 종료")
    
    # KIDS/BABY는 기존 데이터를 재크롤링하지 않고 그대로 가져옴
    print("\n[STEP 2] KIDS/BABY 기존 데이터로 재크롤링")
    print("=" * 60)
    
    driver2 = setup_driver()
    try:
        for category in ['KIDS', 'BABY']:
            info = CATEGORIES[category]
            data = scrape_category_with_tabs(driver2, category, info['url'], info['tabs'])
            all_data.update(data)
            
            cat_total = sum(len(v) for k, v in data.items())
            print(f"\n  => {category}: {len(data)}개 시트, {cat_total}개 상품")
    finally:
        driver2.quit()
        print("\n브라우저 종료")
    
    # 엑셀 저장
    print("\n[STEP 3] 엑셀 생성")
    print("=" * 60)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"유니클로_전체랭킹_이미지포함_V5_{timestamp}.xlsx"
    create_excel(all_data, filename)
    
    # 최종 통계
    print(f"\n{'='*60}")
    print(f"[완료] 수집 통계")
    print(f"{'='*60}")
    
    total_products = sum(len(p) for p in all_data.values())
    total_with_price = sum(1 for prods in all_data.values() for p in prods if p['price'])
    total_with_img = sum(1 for prods in all_data.values() for p in prods if p.get('image_data'))
    
    for sheet_name, products in all_data.items():
        img_cnt = sum(1 for p in products if p.get('image_data'))
        print(f"  {sheet_name}: {len(products)}개 상품 (이미지: {img_cnt}개)")
    
    print(f"\n  총 시트: {len(all_data)}개")
    print(f"  총 상품: {total_products}개")
    print(f"  가격: {total_with_price}/{total_products}개")
    print(f"  이미지: {total_with_img}/{total_products}개")
    print(f"\n  파일: {filename}")
    print("=" * 60)

if __name__ == "__main__":
    main()
