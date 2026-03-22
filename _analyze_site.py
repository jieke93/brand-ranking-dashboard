# -*- coding: utf-8 -*-
"""유니클로 사이트 구조 분석 스크립트"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--window-size=1920,1080')
options.page_load_strategy = 'eager'

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.set_page_load_timeout(90)
driver.set_script_timeout(60)

# WOMEN 랭킹 페이지 접속
try:
    driver.get('https://www.uniqlo.com/kr/ko/spl/ranking/women')
except Exception as e:
    print(f'Page load timeout (expected with eager): {str(e)[:50]}')
time.sleep(8)

# 1) 쿠키 팝업 요소 확인
print('=== COOKIE POPUP ELEMENTS ===')
selectors = [
    '[class*="cookie"]', '[class*="consent"]', '[id*="onetrust"]',
    '[class*="onetrust"]', '[class*="banner"]', '[class*="popup"]',
    '[class*="modal"]', '[class*="overlay"]'
]
for sel in selectors:
    elems = driver.find_elements(By.CSS_SELECTOR, sel)
    for e in elems:
        if e.is_displayed():
            tag = e.tag_name
            cls = (e.get_attribute('class') or '')[:100]
            eid = e.get_attribute('id') or ''
            rect = e.rect
            print(f'  VISIBLE: tag={tag} id={eid} class={cls} rect={rect}')

# 쿠키 팝업 버튼 확인
print('\n=== COOKIE BUTTONS ===')
btns = driver.find_elements(By.CSS_SELECTOR, 'button')
for b in btns:
    txt = b.text.strip()
    if any(kw in txt for kw in ['동의', '수락', 'Accept', 'Cookie', 'accept', '닫기']):
        cls = (b.get_attribute('class') or '')[:80]
        eid = b.get_attribute('id') or ''
        vis = b.is_displayed()
        print(f'  btn: text=[{txt[:30]}] id={eid} class={cls} visible={vis}')

# 2) 탭 요소 확인
print('\n=== TAB ELEMENTS (a.fr-ec-tab) ===')
tabs_a = driver.find_elements(By.CSS_SELECTOR, 'a.fr-ec-tab')
print(f'count: {len(tabs_a)}')
for t in tabs_a[:15]:
    txt = t.text.strip()[:30]
    vis = t.is_displayed()
    href = (t.get_attribute('href') or '')[:80]
    cls = (t.get_attribute('class') or '')[:80]
    print(f'  text=[{txt}] visible={vis} href={href} class={cls}')

print('\n=== TAB ELEMENTS ([role=tab]) ===')
tabs_role = driver.find_elements(By.CSS_SELECTOR, '[role="tab"]')
print(f'count: {len(tabs_role)}')
for t in tabs_role[:15]:
    txt = t.text.strip()[:30]
    vis = t.is_displayed()
    cls = (t.get_attribute('class') or '')[:80]
    aria = t.get_attribute('aria-selected') or ''
    print(f'  text=[{txt}] visible={vis} aria-selected={aria} class={cls}')

# 3) 탭 바 / 네비게이션 분석
print('\n=== TAB BAR / NAV ===')
for sel in ['[role="tablist"]', '.fr-ec-tab-bar', 'nav[class*="tab"]', '[class*="tab-group"]', '[class*="TabGroup"]']:
    elems = driver.find_elements(By.CSS_SELECTOR, sel)
    for e in elems:
        cls = (e.get_attribute('class') or '')[:100]
        print(f'  selector={sel} class={cls}')

# 4) 모두보기 상태에서 상품 타일
print('\n=== PRODUCT TILES (모두보기) ===')
tiles = driver.find_elements(By.CSS_SELECTOR, '.product-tile')
print(f'.product-tile: {len(tiles)}')

# 5) 쿠키 배너 제거 후 클릭 시도
print('\n=== REMOVING COOKIE BANNERS ===')
driver.execute_script("""
    // OneTrust
    var ot = document.getElementById('onetrust-banner-sdk');
    if (ot) { ot.remove(); console.log('onetrust removed'); }
    // 기타 배너
    document.querySelectorAll('[class*="cookie"], [class*="consent"], [id*="onetrust"]').forEach(function(b) {
        b.remove();
    });
    // 오버레이 제거
    document.querySelectorAll('.onetrust-pc-dark-filter, [class*="overlay"]').forEach(function(b) {
        if (b.style.position === 'fixed' || getComputedStyle(b).position === 'fixed') {
            b.remove();
        }
    });
""")
time.sleep(1)

# 6) 상의 탭 클릭 시도
print('\n=== CLICKING 상의 TAB ===')
clicked = False

# 방법A: a.fr-ec-tab
tabs_a = driver.find_elements(By.CSS_SELECTOR, 'a.fr-ec-tab')
for t in tabs_a:
    if '상의' in t.text:
        driver.execute_script('arguments[0].scrollIntoView({block:"center"})', t)
        time.sleep(0.3)
        driver.execute_script('arguments[0].click()', t)
        print(f'  방법A 클릭! text=[{t.text.strip()}]')
        clicked = True
        break

if not clicked:
    # 방법B: [role=tab]
    tabs_role = driver.find_elements(By.CSS_SELECTOR, '[role="tab"]')
    for t in tabs_role:
        if '상의' in t.text:
            driver.execute_script('arguments[0].scrollIntoView({block:"center"})', t)
            time.sleep(0.3)
            driver.execute_script('arguments[0].click()', t)
            print(f'  방법B 클릭! text=[{t.text.strip()}]')
            clicked = True
            break

if not clicked:
    # 방법C: XPath
    xpaths = [
        "//a[contains(text(), '상의')]",
        "//button[contains(text(), '상의')]",
        "//span[contains(text(), '상의')]/parent::a",
        "//span[contains(text(), '상의')]/parent::button",
        "//*[text()='상의']",
    ]
    for xp in xpaths:
        elems = driver.find_elements(By.XPATH, xp)
        for el in elems:
            if el.is_displayed():
                driver.execute_script('arguments[0].scrollIntoView({block:"center"})', el)
                time.sleep(0.3)
                driver.execute_script('arguments[0].click()', el)
                print(f'  방법C 클릭! xpath={xp} text=[{el.text.strip()[:30]}]')
                clicked = True
                break
        if clicked:
            break

if not clicked:
    print('  모든 방법 실패!')
    # 페이지 내 '상의' 텍스트가 포함된 모든 요소 덤프
    print('\n=== ALL ELEMENTS WITH 상의 TEXT ===')
    all_elems = driver.find_elements(By.XPATH, "//*[contains(text(), '상의')]")
    for el in all_elems[:20]:
        tag = el.tag_name
        cls = (el.get_attribute('class') or '')[:60]
        vis = el.is_displayed()
        txt = el.text.strip()[:40]
        print(f'  tag={tag} class={cls} visible={vis} text=[{txt}]')

time.sleep(3)

# 클릭 후 상품 확인
tiles_after = driver.find_elements(By.CSS_SELECTOR, '.product-tile')
print(f'\n=== AFTER CLICK ===')
print(f'.product-tile: {len(tiles_after)}')
print(f'URL: {driver.current_url}')

# 탭 영역 HTML 저장
tab_html = driver.execute_script("""
    var tablist = document.querySelector('[role="tablist"]');
    if (tablist) return tablist.outerHTML.substring(0, 3000);
    var tabs = document.querySelector('.fr-ec-tab-bar');
    if (tabs) return tabs.outerHTML.substring(0, 3000);
    return 'NOT FOUND';
""")
print(f'\nTab HTML:\n{tab_html[:2000]}')

driver.quit()
print('\nDone!')
