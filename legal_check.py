# -*- coding: utf-8 -*-
"""
웹 크롤링 법적 점검 도구
- robots.txt 확인
- 크롤링 정책 분석
"""
import requests
import time
from urllib.robotparser import RobotFileParser

def check_robots_txt(base_url, user_agent="*"):
    """robots.txt 파일 확인"""
    try:
        robots_url = f"{base_url}/robots.txt"
        response = requests.get(robots_url, timeout=10)
        
        if response.status_code == 200:
            print(f"✅ robots.txt 발견: {robots_url}")
            print("=" * 50)
            print(response.text[:1000])
            print("=" * 50)
            
            # robots.txt 파싱
            rp = RobotFileParser()
            rp.set_url(robots_url)
            rp.read()
            
            # 주요 경로 확인
            test_paths = [
                "/women/most-popular.html",
                "/men/most-popular.html", 
                "/display/ranking",
                "/women/",
                "/men/",
                "/"
            ]
            
            print("\n📋 크롤링 허용 여부:")
            for path in test_paths:
                can_crawl = rp.can_fetch(user_agent, base_url + path)
                status = "✅ 허용" if can_crawl else "❌ 금지"
                print(f"  {path}: {status}")
                
            return True, response.text
        else:
            print(f"⚠️  robots.txt 없음: {robots_url} (상태: {response.status_code})")
            return False, None
            
    except Exception as e:
        print(f"❌ robots.txt 확인 오류: {e}")
        return False, None

def analyze_crawling_legality():
    """웹사이트별 크롤링 법적 분석"""
    
    sites = {
        "유니클로": "https://www.uniqlo.com",
        "탑텐": "https://topten10.goodwearmall.com", 
        "아르켓": "https://www.arket.com"
    }
    
    print("🔍 웹 크롤링 법적 점검 시작...")
    print("=" * 60)
    
    results = {}
    
    for name, url in sites.items():
        print(f"\n📋 [{name}] 점검 중...")
        print(f"🌐 URL: {url}")
        
        # robots.txt 확인
        has_robots, robots_content = check_robots_txt(url)
        results[name] = {
            'url': url,
            'has_robots': has_robots,
            'robots_content': robots_content
        }
        
        time.sleep(2)  # 요청 간격 조절
    
    # 종합 분석 리포트
    print("\n" + "=" * 60)
    print("📊 종합 법적 분석 리포트")
    print("=" * 60)
    
    for name, data in results.items():
        print(f"\n🏢 [{name}]")
        print("-" * 30)
        
        if data['has_robots']:
            robots_text = data['robots_content'].lower()
            
            # 일반적인 크롤링 제한 패턴 확인
            restrictions = []
            if 'disallow: /' in robots_text and 'allow:' not in robots_text:
                restrictions.append("전체 크롤링 금지")
            if 'crawl-delay:' in robots_text:
                restrictions.append("크롤링 지연 요구")
            if 'user-agent: *' in robots_text and 'disallow:' in robots_text:
                restrictions.append("일반 봇 제한")
                
            if restrictions:
                print("❌ 제한사항:")
                for r in restrictions:
                    print(f"  - {r}")
            else:
                print("✅ 특별한 제한 없음")
        else:
            print("📝 robots.txt 없음 - 별도 확인 필요")
        
        # 권장사항
        print("\n📋 권장사항:")
        print("  - 이용약관 직접 확인 필요")
        print("  - 요청 간격 준수 (1-2초)")
        print("  - 개인정보 수집 금지")
        print("  - 상업적 사용 시 별도 허가")

if __name__ == "__main__":
    try:
        analyze_crawling_legality()
        
        print("\n" + "=" * 60)
        print("⚖️  법적 준수 가이드라인")
        print("=" * 60)
        print("1. 🤖 User-Agent 정확히 식별")
        print("2. ⏰ 적절한 요청 간격 유지")
        print("3. 📊 공개 데이터만 수집")
        print("4. 🚫 개인정보 수집 금지")
        print("5. 💼 상업적 사용 시 사전 협의")
        print("6. 📋 이용약관 정기 검토")
        print("7. 🛑 robots.txt 준수")
        print("\n✨ 크롤링은 공정이용 원칙 하에 진행하세요!")
        
    except Exception as e:
        print(f"❌ 분석 중 오류: {e}")