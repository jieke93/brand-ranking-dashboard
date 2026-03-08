#!/usr/bin/env python3
"""
유니클로 랭킹 수집 + 분석 통합 실행기
─────────────────────────────────────
"랭킹 내려줘" 한 마디로 전체 프로세스 실행:
  1단계: 크롤링 (uniqlo_ranking_v5.py)
  2단계: 분석 + 히스토리 누적 (analyze_ranking_v2.py)

히스토리 누적 구조:
  - ranking_history.json : 날짜별 랭킹 데이터 (자동 누적)
  - ranking_history_backup.json : 백업 (매 실행시 자동 생성)
  - 이전 크롤링 엑셀 파일 : JSON 유실 시 자동 복구 소스

사용법:
  python run_ranking.py           # 크롤링 + 분석 전체 실행
  python run_ranking.py --analyze # 분석만 실행 (이미 크롤링한 데이터로)
  python run_ranking.py --recover # 이전 엑셀에서 히스토리 강제 복구
"""

import os
import sys
import subprocess
import time
from datetime import datetime

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
CRAWLER_SCRIPT = os.path.join(WORK_DIR, 'uniqlo_ranking_v5.py')
ANALYZER_SCRIPT = os.path.join(WORK_DIR, 'analyze_ranking_v2.py')


def log(msg):
    print(msg)


def run_crawler():
    """크롤러 실행"""
    log("=" * 60)
    log("  [STEP 1] 유니클로 랭킹 크롤링 시작")
    log("=" * 60)
    
    if not os.path.exists(CRAWLER_SCRIPT):
        log(f"  [ERROR] 크롤러 파일 없음: {CRAWLER_SCRIPT}")
        return False
    
    start = time.time()
    result = subprocess.run(
        [sys.executable, CRAWLER_SCRIPT],
        cwd=WORK_DIR,
        encoding='utf-8',
        errors='replace',
    )
    elapsed = time.time() - start
    
    if result.returncode != 0:
        log(f"\n  [ERROR] 크롤링 실패 (코드: {result.returncode})")
        return False
    
    log(f"\n  [OK] 크롤링 완료 ({elapsed:.0f}초 소요)")
    return True


def run_analyzer():
    """분석기 실행"""
    log("\n" + "=" * 60)
    log("  [STEP 2] 랭킹 분석 + 히스토리 누적")
    log("=" * 60)
    
    if not os.path.exists(ANALYZER_SCRIPT):
        log(f"  [ERROR] 분석기 파일 없음: {ANALYZER_SCRIPT}")
        return False
    
    start = time.time()
    result = subprocess.run(
        [sys.executable, ANALYZER_SCRIPT],
        cwd=WORK_DIR,
        encoding='utf-8',
        errors='replace',
    )
    elapsed = time.time() - start
    
    if result.returncode != 0:
        log(f"\n  [ERROR] 분석 실패 (코드: {result.returncode})")
        return False
    
    log(f"\n  [OK] 분석 완료 ({elapsed:.0f}초 소요)")
    return True


def run_history_recovery():
    """이전 엑셀에서 히스토리 강제 복구"""
    log("=" * 60)
    log("  [복구] 이전 크롤링 엑셀에서 히스토리 복구")
    log("=" * 60)
    
    # analyze_ranking_v2의 복구 함수를 직접 임포트해서 실행
    sys.path.insert(0, WORK_DIR)
    from analyze_ranking_v2 import (
        recover_history_from_excel_files, load_history, 
        merge_history, save_history
    )
    
    existing = load_history()
    recovered = recover_history_from_excel_files()
    
    if recovered:
        merged = merge_history(existing, recovered)
        save_history(merged)
        
        all_dates = set()
        for cat in merged.values():
            all_dates.update(cat.keys())
        log(f"\n  [OK] 히스토리 복구 완료: {len(all_dates)}일치 데이터")
    else:
        log("\n  [INFO] 복구할 크롤링 파일이 없습니다.")


def show_history_status():
    """현재 히스토리 상태 표시"""
    history_file = os.path.join(WORK_DIR, 'ranking_history.json')
    
    if not os.path.exists(history_file):
        log("  [히스토리] 없음 (첫 실행)")
        return
    
    try:
        import json
        with open(history_file, 'r', encoding='utf-8') as f:
            history = json.load(f)
        
        all_dates = set()
        for cat in history.values():
            all_dates.update(cat.keys())
        sorted_dates = sorted(all_dates)
        
        log(f"  [히스토리] {len(sorted_dates)}일치 데이터 보유")
        if sorted_dates:
            log(f"  [히스토리] 기간: {sorted_dates[0]} ~ {sorted_dates[-1]}")
            log(f"  [히스토리] 날짜목록: {', '.join(sorted_dates)}")
        
        for cat in sorted(history.keys()):
            cat_dates = sorted(history[cat].keys())
            products_latest = len(history[cat][cat_dates[-1]]) if cat_dates else 0
            log(f"    {cat}: {len(cat_dates)}회 수집, 최근 {products_latest}개 상품")
    except Exception as e:
        log(f"  [히스토리] 읽기 오류: {e}")


def main():
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log("")
    log("╔" + "═" * 58 + "╗")
    log("║  유니클로 랭킹 수집 + 분석 통합 실행기              ║")
    log(f"║  실행시간: {now}                    ║")
    log("╚" + "═" * 58 + "╝")
    
    # 현재 히스토리 상태
    log("")
    show_history_status()
    log("")
    
    # 명령줄 인자 처리
    args = sys.argv[1:]
    
    if '--recover' in args:
        run_history_recovery()
        return
    
    if '--analyze' in args:
        # 분석만 실행
        success = run_analyzer()
        if success:
            log("\n" + "=" * 60)
            log("  완료! 분석 결과 엑셀이 생성되었습니다.")
            log("=" * 60)
        return
    
    # 전체 실행 (크롤링 + 분석)
    crawler_ok = run_crawler()
    
    if not crawler_ok:
        log("\n  [WARN] 크롤링이 실패했습니다.")
        log("  기존 데이터로 분석을 진행하시려면: python run_ranking.py --analyze")
        return
    
    # 크롤링 성공 시 분석 자동 실행
    analyzer_ok = run_analyzer()
    
    # 최종 상태
    log("\n")
    log("╔" + "═" * 58 + "╗")
    log("║  실행 완료                                         ║")
    log("╠" + "═" * 58 + "╣")
    log(f"║  크롤링: {'성공 ✓' if crawler_ok else '실패 ✗':50s}║")
    log(f"║  분석:   {'성공 ✓' if analyzer_ok else '실패 ✗':50s}║")
    log("╠" + "═" * 58 + "╣")
    
    show_history_status()
    log("╚" + "═" * 58 + "╝")


if __name__ == '__main__':
    main()
