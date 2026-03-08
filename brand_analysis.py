"""
패션 브랜드 랭킹 인사이트 엔진
===============================
3축 분석: 상품 단위 / 유형 단위 / 브랜드 단위
출력: 명확한 문장형 인사이트 (analysis_history.json 누적)
"""

import json
import os
import re
from collections import defaultdict
from datetime import datetime

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
ANALYSIS_HISTORY_FILE = os.path.join(WORK_DIR, 'analysis_history.json')


# ═══════════════════════════════════════════════════════
#  데이터 로드
# ═══════════════════════════════════════════════════════

def load_history():
    path = os.path.join(WORK_DIR, 'all_brands_history.json')
    with open(path, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    records = []
    for cat_key, date_data in raw.items():
        parts = cat_key.split('_', 1)
        brand = parts[0]
        sub = parts[1] if len(parts) > 1 else ''
        gender = _norm_gender(sub)
        subcategory = _extract_sub(sub)

        for date_str, products in date_data.items():
            for name, info in products.items():
                records.append({
                    'brand': brand, 'gender': gender, 'subcategory': subcategory,
                    'category_key': cat_key, 'date': date_str,
                    'product': name,
                    'rank': info.get('rank', 999),
                    'item_type': info.get('item_type', '미분류'),
                    'price': _parse_price(info.get('price', '0')),
                })
    return records


def _norm_gender(sub):
    s = sub.upper()
    if 'WOMEN' in s or '여성' in s: return '여성'
    if 'MEN' in s or '남성' in s: return '남성'
    if 'KIDS' in s or '키즈' in s: return '키즈'
    if 'BABY' in s or '베이비' in s: return '베이비'
    if '전체' in sub: return '전체'
    return '기타'


def _extract_sub(sub):
    parts = sub.split('_', 1)
    return parts[1] if len(parts) >= 2 else '모두보기'


def _parse_price(p):
    if not p: return 0
    nums = re.findall(r'[\d,]+', str(p))
    return int(nums[0].replace(',', '')) if nums else 0


def _fmt(d):
    return f"{d[:4]}.{d[4:6]}.{d[6:]}"


def _price_label(p):
    if p < 20000: return '2만원 미만'
    if p < 50000: return '2~5만원'
    if p < 100000: return '5~10만원'
    return '10만원 이상'


# ═══════════════════════════════════════════════════════
#  상품 단위 분석
# ═══════════════════════════════════════════════════════

def analyze_products(records, dates, brands):
    """상품 단위 인사이트: 급상승 경향성 + 지속 상위권 경향성 + 신규/이탈"""
    insights = []

    if len(dates) < 2:
        return insights

    prev, cur = dates[-2], dates[-1]

    prev_all = [r for r in records if r['date'] == prev and r['subcategory'] == '모두보기'
                and r['gender'] in ('여성', '남성')]
    cur_all = [r for r in records if r['date'] == cur and r['subcategory'] == '모두보기'
               and r['gender'] in ('여성', '남성')]

    prev_map = {}
    for r in prev_all:
        prev_map[(r['brand'], r['product'])] = r
    cur_map = {}
    for r in cur_all:
        cur_map[(r['brand'], r['product'])] = r

    # ── 1) 급상승 상품 ──
    risers = []
    for key in set(prev_map) & set(cur_map):
        change = prev_map[key]['rank'] - cur_map[key]['rank']
        if change >= 5:
            risers.append({
                'brand': key[0], 'product': key[1],
                'prev_rank': prev_map[key]['rank'], 'cur_rank': cur_map[key]['rank'],
                'change': change,
                'item_type': cur_map[key]['item_type'],
                'price': cur_map[key]['price'],
                'gender': cur_map[key]['gender'],
            })
    risers.sort(key=lambda x: x['change'], reverse=True)

    if risers:
        type_counts = defaultdict(int)
        price_sum, brand_counts = 0, defaultdict(int)
        for r in risers:
            type_counts[r['item_type']] += 1
            price_sum += r['price']
            brand_counts[r['brand']] += 1

        top_type = max(type_counts, key=type_counts.get)
        top_brand = max(brand_counts, key=brand_counts.get)
        avg_price = price_sum / len(risers)

        top5_lines = []
        for r in risers[:5]:
            top5_lines.append(
                f"{r['brand']} [{r['product'][:25]}] "
                f"{r['prev_rank']}위→{r['cur_rank']}위 (+{r['change']}), "
                f"{r['item_type']}, {r['price']:,}원"
            )

        price_bands = defaultdict(int)
        for r in risers:
            price_bands[_price_label(r['price'])] += 1

        sub = []
        sub.append(f"급상승 유형: {', '.join(f'{t} {c}개' for t, c in sorted(type_counts.items(), key=lambda x: -x[1])[:5])}")
        sub.append(f"급상승 브랜드: {', '.join(f'{b} {c}개' for b, c in sorted(brand_counts.items(), key=lambda x: -x[1]))}")
        sub.append(f"급상승 가격대: {', '.join(f'{b} {c}개' for b, c in sorted(price_bands.items(), key=lambda x: -x[1]))}")

        insights.append({
            'category': '상품',
            'title': '급상승 상품 경향성',
            'summary': (
                f"5계단 이상 상승한 상품이 총 {len(risers)}개입니다. "
                f"이 중 가장 많은 유형은 '{top_type}'({type_counts[top_type]}개)이며, "
                f"'{top_brand}'에서 {brand_counts[top_brand]}개로 가장 많이 상승했습니다. "
                f"급상승 상품의 평균 가격대는 {avg_price:,.0f}원({_price_label(avg_price)})입니다."
            ),
            'details': top5_lines,
            'sub_insights': sub,
        })

    # ── 2) 지속 상위권 상품 ──
    steady = []
    for key in set(prev_map) & set(cur_map):
        if prev_map[key]['rank'] <= 10 and cur_map[key]['rank'] <= 10:
            steady.append({
                'brand': key[0], 'product': key[1],
                'prev_rank': prev_map[key]['rank'], 'cur_rank': cur_map[key]['rank'],
                'item_type': cur_map[key]['item_type'],
                'price': cur_map[key]['price'],
                'gender': cur_map[key]['gender'],
            })

    if steady:
        type_counts = defaultdict(int)
        price_sum, brand_counts = 0, defaultdict(int)
        for s in steady:
            type_counts[s['item_type']] += 1
            price_sum += s['price']
            brand_counts[s['brand']] += 1

        top_type = max(type_counts, key=type_counts.get)
        avg_price = price_sum / len(steady)
        brand_share = ', '.join(f"{b} {c}개" for b, c in sorted(brand_counts.items(), key=lambda x: -x[1]))

        detail_lines = []
        for s in sorted(steady, key=lambda x: x['cur_rank'])[:5]:
            detail_lines.append(
                f"{s['brand']} [{s['product'][:25]}] "
                f"{s['prev_rank']}위→{s['cur_rank']}위, "
                f"{s['item_type']}, {s['price']:,}원"
            )

        sub = []
        sub.append(f"상위권 유형: {', '.join(f'{t} {c}개' for t, c in sorted(type_counts.items(), key=lambda x: -x[1])[:5])}")
        price_bands = defaultdict(int)
        for s in steady:
            price_bands[_price_label(s['price'])] += 1
        sub.append(f"상위권 가격대: {', '.join(f'{b} {c}개' for b, c in sorted(price_bands.items(), key=lambda x: -x[1]))}")

        insights.append({
            'category': '상품',
            'title': '지속 상위권 상품 경향성',
            'summary': (
                f"두 차수 연속 Top10을 유지한 상품은 {len(steady)}개입니다. "
                f"브랜드별로는 {brand_share}. "
                f"가장 많은 유형은 '{top_type}'({type_counts[top_type]}개)이고, "
                f"평균 가격은 {avg_price:,.0f}원({_price_label(avg_price)})입니다."
            ),
            'details': detail_lines,
            'sub_insights': sub,
        })

    # ── 3) 신규 진입 & 이탈 ──
    new_entries = set(cur_map) - set(prev_map)
    dropped = set(prev_map) - set(cur_map)

    if new_entries:
        new_recs = [cur_map[k] for k in new_entries]
        new_type = defaultdict(int)
        new_brand = defaultdict(int)
        for r in new_recs:
            new_type[r['item_type']] += 1
            new_brand[r['brand']] += 1

        top_new_type = max(new_type, key=new_type.get)
        top_new_brand = max(new_brand, key=new_brand.get)
        new_top10 = [r for r in new_recs if r['rank'] <= 10]

        detail = []
        for r in sorted(new_recs, key=lambda x: x['rank'])[:5]:
            detail.append(f"{r['brand']} [{r['product'][:25]}] {r['rank']}위 신규, {r['item_type']}, {r['price']:,}원")

        summary = (
            f"신규 진입 {len(new_entries)}개, 이탈 {len(dropped)}개. "
            f"신규 중 '{top_new_type}'이 {new_type[top_new_type]}개로 가장 많고, "
            f"'{top_new_brand}'에서 {new_brand[top_new_brand]}개가 새로 등장했습니다."
        )
        if new_top10:
            summary += f" Top10에 바로 진입한 상품이 {len(new_top10)}개 있습니다."

        insights.append({
            'category': '상품',
            'title': '신규 진입 & 이탈 상품',
            'summary': summary,
            'details': detail,
            'sub_insights': [],
        })

    return insights


# ═══════════════════════════════════════════════════════
#  유형 단위 분석
# ═══════════════════════════════════════════════════════

def analyze_types(records, dates, brands):
    """유형(아이템타입) 단위 인사이트: 브랜드별 상위 유형"""
    insights = []
    cur = dates[-1]

    for gender in ['여성', '남성']:
        # ── 브랜드별 Top10 내 유형 분포 ──
        brand_type_data = {}
        for brand in brands:
            recs = [r for r in records if r['brand'] == brand and r['date'] == cur
                    and r['gender'] == gender and r['subcategory'] == '모두보기']
            top10 = [r for r in recs if r['rank'] <= 10]

            if not top10:
                continue

            type_top10 = defaultdict(list)
            for r in top10:
                type_top10[r['item_type']].append(r['rank'])

            brand_type_data[brand] = {
                'top10_types': dict(type_top10),
                'total': len(recs),
            }

        if not brand_type_data:
            continue

        brand_lines = []
        detail_lines = []
        for brand in brands:
            if brand not in brand_type_data:
                continue
            data = brand_type_data[brand]
            sorted_types = sorted(data['top10_types'].items(),
                                  key=lambda x: (-len(x[1]), sum(x[1]) / len(x[1])))
            top3_str = ', '.join(
                f"'{t}'({len(ranks)}개, 평균{sum(ranks)/len(ranks):.0f}위)"
                for t, ranks in sorted_types[:3]
            )
            brand_lines.append(f"{brand}의 Top10 주력은 {top3_str}")

            for t, ranks in sorted_types[:3]:
                detail_lines.append(
                    f"{brand} [{gender}] {t}: Top10 내 {len(ranks)}개 (평균 {sum(ranks)/len(ranks):.1f}위)"
                )

        summary = f"[{gender}] " + '. '.join(brand_lines) + '.'

        # 크로스 브랜드 경쟁
        all_types = set()
        for data in brand_type_data.values():
            all_types.update(data['top10_types'].keys())

        sub = []
        for item_type in sorted(all_types):
            bw = []
            for brand, data in brand_type_data.items():
                if item_type in data['top10_types']:
                    ranks = data['top10_types'][item_type]
                    bw.append((brand, len(ranks), sum(ranks) / len(ranks)))
            if len(bw) >= 2:
                bw.sort(key=lambda x: (-x[1], x[2]))
                sub.append(f"'{item_type}' 경쟁: {', '.join(f'{b}({c}개, 평균{a:.0f}위)' for b, c, a in bw)}")

        insights.append({
            'category': '유형',
            'title': f'{gender} 브랜드별 상위 유형',
            'summary': summary,
            'details': detail_lines,
            'sub_insights': sub[:5],
        })

        # ── 유형별 랭킹 추이 ──
        if len(dates) >= 2:
            prev = dates[-2]
            trend = _type_trend(records, prev, cur, gender)
            if trend:
                insights.append(trend)

    return insights


def _type_trend(records, prev, cur, gender):
    prev_types = defaultdict(list)
    cur_types = defaultdict(list)

    for r in records:
        if r['subcategory'] != '모두보기' or r['gender'] != gender:
            continue
        if r['date'] == prev:
            prev_types[r['item_type']].append(r['rank'])
        elif r['date'] == cur:
            cur_types[r['item_type']].append(r['rank'])

    changes = []
    for t in set(prev_types) & set(cur_types):
        pa = sum(prev_types[t]) / len(prev_types[t])
        ca = sum(cur_types[t]) / len(cur_types[t])
        changes.append({'type': t, 'prev': pa, 'cur': ca, 'change': pa - ca, 'count': len(cur_types[t])})

    if not changes:
        return None

    changes.sort(key=lambda x: x['change'], reverse=True)
    rising = [c for c in changes if c['change'] > 1]
    falling = [c for c in changes if c['change'] < -1]

    rise_str = ', '.join(f"'{c['type']}'(+{c['change']:.1f})" for c in rising[:3]) if rising else '없음'
    fall_str = ', '.join(f"'{c['type']}'({c['change']:.1f})" for c in falling[:3]) if falling else '없음'

    detail = []
    for c in (rising[:3] + falling[:3]):
        arrow = '▲' if c['change'] > 0 else '▼'
        detail.append(f"{arrow} {c['type']}: {c['prev']:.1f}위→{c['cur']:.1f}위 ({c['change']:+.1f}), {c['count']}개")

    return {
        'category': '유형',
        'title': f'{gender} 유형별 랭킹 변화',
        'summary': f"[{gender}] 상승: {rise_str}. 하락: {fall_str}. "
                   f"{len(changes)}개 유형 중 {len(rising)}개 상승, {len(falling)}개 하락.",
        'details': detail,
        'sub_insights': [],
    }


# ═══════════════════════════════════════════════════════
#  브랜드 단위 분석
# ═══════════════════════════════════════════════════════

def analyze_brands(records, dates, brands):
    """브랜드 단위 인사이트: 브랜드 추이 비교"""
    insights = []
    cur = dates[-1]

    # ── 1) 성별 브랜드 종합 비교 ──
    for gender in ['여성', '남성']:
        brand_stats = {}
        for brand in brands:
            recs = [r for r in records if r['brand'] == brand and r['date'] == cur
                    and r['gender'] == gender and r['subcategory'] == '모두보기']
            if not recs:
                continue
            ranks = [r['rank'] for r in recs]
            prices = [r['price'] for r in recs]
            brand_stats[brand] = {
                'count': len(recs),
                'avg_rank': sum(ranks) / len(ranks),
                'top5': sum(1 for r in ranks if r <= 5),
                'top10': sum(1 for r in ranks if r <= 10),
                'top20': sum(1 for r in ranks if r <= 20),
                'avg_price': sum(prices) / len(prices) if prices else 0,
                'type_count': len(set(r['item_type'] for r in recs)),
            }

        if len(brand_stats) < 2:
            continue

        sorted_b = sorted(brand_stats.items(), key=lambda x: x[1]['avg_rank'])
        leader = sorted_b[0]
        details = []
        for b, s in sorted_b:
            details.append(
                f"{b}: 평균 {s['avg_rank']:.1f}위, Top10 {s['top10']}개, "
                f"Top20 {s['top20']}개, 평균가격 {s['avg_price']:,.0f}원, "
                f"유형 {s['type_count']}종, 총 {s['count']}개"
            )

        summary = (
            f"[{gender}] 평균 랭킹 기준 '{leader[0]}'이 {leader[1]['avg_rank']:.1f}위로 선두입니다. "
            + ' / '.join(f"{b} {s['avg_rank']:.1f}위" for b, s in sorted_b) + '.'
        )

        sub = []
        if len(dates) >= 2:
            prev = dates[-2]
            for brand in brands:
                pr = [r for r in records if r['brand'] == brand and r['date'] == prev
                      and r['gender'] == gender and r['subcategory'] == '모두보기']
                if brand in brand_stats and pr:
                    pa = sum(r['rank'] for r in pr) / len(pr)
                    ca = brand_stats[brand]['avg_rank']
                    diff = pa - ca
                    d = '상승' if diff > 0 else '하락'
                    sub.append(f"{brand}: {pa:.1f}위→{ca:.1f}위 ({d} {abs(diff):.1f}계단)")

        insights.append({
            'category': '브랜드',
            'title': f'{gender} 브랜드 종합 비교',
            'summary': summary,
            'details': details,
            'sub_insights': sub,
        })

    # ── 2) 안정성 비교 ──
    if len(dates) >= 2:
        prev, cur_d = dates[-2], dates[-1]
        brand_vol = {}

        for brand in brands:
            pm = {r['product']: r['rank'] for r in records
                  if r['brand'] == brand and r['date'] == prev and r['subcategory'] == '모두보기'
                  and r['gender'] in ('여성', '남성')}
            cm = {r['product']: r['rank'] for r in records
                  if r['brand'] == brand and r['date'] == cur_d and r['subcategory'] == '모두보기'
                  and r['gender'] in ('여성', '남성')}

            common = set(pm) & set(cm)
            if len(common) < 3:
                continue

            changes = [abs(pm[p] - cm[p]) for p in common]
            stable = sum(1 for c in changes if c <= 2)
            volatile = sum(1 for c in changes if c > 5)

            max_up_p = max(common, key=lambda p: pm[p] - cm[p])
            max_down_p = min(common, key=lambda p: pm[p] - cm[p])

            brand_vol[brand] = {
                'avg_vol': sum(changes) / len(changes),
                'stable_pct': stable / len(common) * 100,
                'volatile': volatile,
                'common': len(common),
                'max_up': (max_up_p, pm[max_up_p], cm[max_up_p]),
                'max_down': (max_down_p, pm[max_down_p], cm[max_down_p]),
            }

        if len(brand_vol) >= 2:
            sb = sorted(brand_vol.items(), key=lambda x: x[1]['avg_vol'])
            most_stable = sb[0]
            most_volatile = sb[-1]

            details = []
            for b, v in sb:
                details.append(
                    f"{b}: 평균 변동 {v['avg_vol']:.1f}계단, "
                    f"안정 {v['stable_pct']:.0f}%, "
                    f"최대상승 [{v['max_up'][0][:20]}] {v['max_up'][1]}→{v['max_up'][2]}위, "
                    f"최대하락 [{v['max_down'][0][:20]}] {v['max_down'][1]}→{v['max_down'][2]}위"
                )

            insights.append({
                'category': '브랜드',
                'title': '브랜드 랭킹 안정성 비교',
                'summary': (
                    f"가장 안정적인 브랜드는 '{most_stable[0]}'"
                    f"(평균 {most_stable[1]['avg_vol']:.1f}계단, 안정 {most_stable[1]['stable_pct']:.0f}%)이고, "
                    f"가장 변동이 큰 브랜드는 '{most_volatile[0]}'"
                    f"(평균 {most_volatile[1]['avg_vol']:.1f}계단)입니다."
                ),
                'details': details,
                'sub_insights': [],
            })

    # ── 3) 가격 포지셔닝 ──
    bp = {}
    for brand in brands:
        recs = [r for r in records if r['brand'] == brand and r['date'] == cur
                and r['subcategory'] == '모두보기' and r['gender'] in ('여성', '남성')]
        prices = [r['price'] for r in recs if r['price'] > 0]
        t10p = [r['price'] for r in recs if r['rank'] <= 10 and r['price'] > 0]
        if prices:
            bp[brand] = {
                'avg': sum(prices) / len(prices),
                'top10_avg': sum(t10p) / len(t10p) if t10p else 0,
                'min': min(prices), 'max': max(prices),
            }

    if len(bp) >= 2:
        sb = sorted(bp.items(), key=lambda x: x[1]['avg'])
        details = [
            f"{b}: 전체 평균 {p['avg']:,.0f}원, Top10 평균 {p['top10_avg']:,.0f}원, "
            f"범위 {p['min']:,.0f}~{p['max']:,.0f}원"
            for b, p in sb
        ]
        insights.append({
            'category': '브랜드',
            'title': '브랜드 가격 포지셔닝',
            'summary': (
                f"평균 가격 최저 '{sb[0][0]}'({sb[0][1]['avg']:,.0f}원) ~ "
                f"최고 '{sb[-1][0]}'({sb[-1][1]['avg']:,.0f}원). "
                f"가격 차이 약 {sb[-1][1]['avg'] - sb[0][1]['avg']:,.0f}원."
            ),
            'details': details,
            'sub_insights': [],
        })

    return insights


# ═══════════════════════════════════════════════════════
#  전체 분석 + 저장
# ═══════════════════════════════════════════════════════

def analyze_all(records):
    dates = sorted(set(r['date'] for r in records))
    brands = sorted(set(r['brand'] for r in records))
    ins = []
    ins += analyze_products(records, dates, brands)
    ins += analyze_types(records, dates, brands)
    ins += analyze_brands(records, dates, brands)
    return ins, dates, brands


def save_analysis(date_str, insights, dates, brands, records):
    history = {}
    if os.path.exists(ANALYSIS_HISTORY_FILE):
        with open(ANALYSIS_HISTORY_FILE, 'r', encoding='utf-8') as f:
            history = json.load(f)

    history[date_str] = {
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'data_dates': dates,
        'brands': brands,
        'total_records': len(records),
        'total_insights': len(insights),
        'product_insights': [i for i in insights if i['category'] == '상품'],
        'type_insights': [i for i in insights if i['category'] == '유형'],
        'brand_insights': [i for i in insights if i['category'] == '브랜드'],
    }

    with open(ANALYSIS_HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    print(f"  [저장] {ANALYSIS_HISTORY_FILE}")
    print(f"  누적 분석: {len(history)}회 ({', '.join(sorted(history.keys()))})")
    return history


def run_analysis():
    """외부 호출 API (run_all.py 등에서 호출)"""
    records = load_history()
    insights, dates, brands = analyze_all(records)
    date_str = dates[-1] if dates else datetime.now().strftime('%Y%m%d')
    save_analysis(date_str, insights, dates, brands, records)
    return date_str, insights


def main():
    records = load_history()
    insights, dates, brands = analyze_all(records)

    print(f"\n{'='*70}")
    print(f"  패션 브랜드 랭킹 인사이트")
    print(f"  데이터: {len(dates)}회 ({', '.join(_fmt(d) for d in dates)})")
    print(f"  브랜드: {', '.join(brands)} | 레코드: {len(records):,}건")
    print(f"{'='*70}")

    for cat in ['상품', '유형', '브랜드']:
        cat_ins = [i for i in insights if i['category'] == cat]
        if not cat_ins:
            continue
        emoji = {'상품': '📦', '유형': '👕', '브랜드': '🏢'}[cat]
        print(f"\n{emoji} [{cat} 단위] ({len(cat_ins)}개)")
        print(f"{'─'*70}")
        for ins in cat_ins:
            print(f"\n  ■ {ins['title']}")
            print(f"    {ins['summary']}")
            if ins.get('details'):
                for d in ins['details']:
                    print(f"      · {d}")
            if ins.get('sub_insights'):
                for s in ins['sub_insights']:
                    print(f"      ▸ {s}")

    date_str = dates[-1] if dates else datetime.now().strftime('%Y%m%d')
    save_analysis(date_str, insights, dates, brands, records)

    print(f"\n{'='*70}")
    print(f"  총 {len(insights)}개 인사이트 생성 완료")
    print(f"{'='*70}")


if __name__ == '__main__':
    main()
