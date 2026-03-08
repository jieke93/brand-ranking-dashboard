#!/usr/bin/env python3
"""
3사 브랜드 통합 크롤링 + 분석 원클릭 실행기
──────────────────────────────────────────
유니클로 · 아르켓 · 탑텐 크롤러 → 통합 분석 → 엑셀 1개 출력

사용법:
  python run_all.py              # 크롤링 + 분석 모두 실행
  python run_all.py --analyze    # 분석만 실행 (기존 크롤링 데이터 사용)
"""

import subprocess
import sys
import os
import time

WORK_DIR = os.path.dirname(os.path.abspath(__file__))

# ── 실행할 크롤러 목록 (순서대로) ──
CRAWLERS = [
    {'name': '유니클로', 'script': 'uniqlo_ranking_v5.py', 'desc': '유니클로 전체 랭킹 크롤링'},
    {'name': '탑텐',     'script': 'topten_ranking_v3.py',  'desc': '탑텐 주간베스트 크롤링'},
    {'name': '아르켓',   'script': 'arket_ranking_v5.py',   'desc': '아르켓 인기상품 크롤링'},
]

# ── 분석 스크립트 ──
ANALYZER = 'analyze_all_brands.py'
BRAND_ANALYSIS = 'brand_analysis.py'


def run_script(script_path, desc):
    """스크립트 실행 (서브프로세스)"""
    full_path = os.path.join(WORK_DIR, script_path)
    if not os.path.exists(full_path):
        print(f"  [SKIP] {script_path} 파일이 없습니다")
        return False

    print(f"\n{'─' * 50}")
    print(f"  {desc}")
    print(f"  파일: {script_path}")
    print(f"{'─' * 50}")

    try:
        result = subprocess.run(
            [sys.executable, full_path],
            cwd=WORK_DIR,
            timeout=600,  # 10분 타임아웃
        )
        if result.returncode == 0:
            print(f"  [OK] {desc} 완료")
            return True
        else:
            print(f"  [WARN] {desc} 종료코드={result.returncode}")
            return True  # 경고지만 계속 진행
    except subprocess.TimeoutExpired:
        print(f"  [TIMEOUT] {desc} - 10분 초과")
        return False
    except Exception as e:
        print(f"  [ERROR] {desc}: {e}")
        return False


def main():
    analyze_only = '--analyze' in sys.argv or '-a' in sys.argv

    brand_names = ' \u00b7 '.join(c['name'] for c in CRAWLERS)
    n = len(CRAWLERS)

    print("=" * 60)
    print(f"  {n}사 브랜드 통합 랭킹 시스템")
    print(f"  {brand_names}")
    print("=" * 60)
    print(f"  모드: {'분석만' if analyze_only else '크롤링 + 분석'}")
    print(f"  작업 디렉토리: {WORK_DIR}")
    start = time.time()

    # 1) 크롤링
    if not analyze_only:
        print("\n" + "=" * 60)
        print(f"  [STEP 1] {n}사 크롤링 시작")
        print("=" * 60)
        for crawler in CRAWLERS:
            run_script(crawler['script'], crawler['desc'])
            time.sleep(2)  # 크롤러 간 간격

    # 2) 통합 분석
    print("\n" + "=" * 60)
    step = '2' if not analyze_only else '1'
    print(f"  [STEP {step}] 통합 분석 실행")
    print("=" * 60)
    run_script(ANALYZER, f'{n}사 통합 랭킹 분석')

    # 3) AI 크로스브랜드 인사이트 분석
    print("\n" + "=" * 60)
    step_ai = '3' if not analyze_only else '2'
    print(f"  [STEP {step_ai}] AI 크로스브랜드 인사이트 분석")
    print("=" * 60)
    run_script(BRAND_ANALYSIS, 'AI 브랜드 랭킹 인사이트 분석 → analysis_history.json')

    elapsed = time.time() - start
    print("\n" + "=" * 60)
    print(f"  전체 완료! (소요: {elapsed:.0f}초)")
    print("=" * 60)


if __name__ == '__main__':
    main()
