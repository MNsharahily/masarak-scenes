import itertools, json
from collections import defaultdict

# Starting stats (confirmed from source doc paragraph)
START = dict(K=20, D=20, C=20, B=20, R=20, M=20, P=10)

def clamp(v):
    return max(0, min(100, v))

# Each chapter: list of (code, delta_dict, flag)
CHAPTERS = {
    'C1': [
        ('C1-A', dict(K=8, D=6, B=2, P=4), None),
        ('C1-B', dict(M=8, C=6, D=3, P=5), 'اول_عمل'),
    ],
    'C2': [
        ('C2-A', dict(K=10, M=-5, P=4), 'studyPath=جامعة'),
        ('C2-B', dict(K=7, D=8, C=2, M=-2), 'studyPath=دبلوم'),
        ('C2-C', dict(M=10, C=7, D=4, K=-2, P=6), 'studyPath=عمل'),
    ],
    'C3': [
        ('C3-A', dict(M=-8, K=8, D=3), 'استثمار_مهاري'),
        ('C3-B', dict(D=6, B=7), 'صندوق_طوارئ'),
        ('C3-C', dict(M=-7, C=8, D=2, P=4), 'مشروع_جانبي'),
    ],
    'C4': [
        ('C4-A', dict(K=6, B=6, R=4, P=-4), 'مرشد'),
        ('C4-B', dict(D=7, C=3, P=8), 'حل_منفرد'),
        ('C4-C', dict(C=9, B=3, M=-3, R=-2), 'تحول'),
    ],
    'C5': [
        ('C5-A', dict(K=10, D=5, M=-4, P=4), 'متخصص'),
        ('C5-B', dict(M=11, C=6, R=2, P=6), 'مسار_سريع'),
        ('C5-C', dict(K=5, D=5, M=5, P=10), 'مسار_هجين'),
    ],
    'C6': [
        ('C6-A', dict(M=-10, K=11, D=4, R=3), 'معتمد'),  # updated: added reputation source (professional certification builds credibility)
        ('C6-B', dict(D=6, B=9, P=-3), 'احتياطي_قوي'),
        ('C6-C', dict(M=-9, C=10, D=3, P=6), 'بذرة_مشروع'),
    ],
    'C7': [
        ('C7-A', dict(R=10, D=4, B=4, M=-2), 'شفافية'),
        ('C7-B', dict(M=5, C=3, R=-8, P=9), 'خطأ_مخفي'),
        ('C7-C', dict(C=6, R=-3, D=-2, P=6), 'صدام_علني'),
    ],
    'C8': [
        ('C8-A', dict(M=13, C=8, D=4, P=13, B=-6), 'ترقية'),
        ('C8-B', dict(K=13, D=7, M=-7, P=5), 'دراسة_متقدمة'),
        ('C8-C', dict(B=14, D=4, P=-12, M=-3), 'إعادة_ضبط'),
    ],
    'C9': [
        ('C9-A', dict(C=11, M=7, R=3, P=12), 'شراكة_سريعة'),
        ('C9-B', dict(K=5, D=9, C=5, M=-5, R=5, P=5), 'تجربة_محدودة'),
        ('C9-C', dict(B=10, M=7, D=3, C=-2, P=-5), 'استقرار'),
    ],
    'C10': [
        ('C10-A', dict(C=9, M=12, D=-3, P=15), 'دين'),
        ('C10-B', dict(D=7, B=9, R=5, M=-5, P=-6), 'إعادة_هيكلة'),
        ('C10-C', dict(B=6, C=4, M=-3, R=-4, P=-10), 'خروج'),
    ],
    'C11': [
        ('C11-A', dict(R=12, B=7, D=4, M=-3, P=2), 'مسؤولية'),
        ('C11-B', dict(M=5, C=4, R=-10, P=7), 'لوم'),
        ('C11-C', dict(M=7, R=-15, P=11, D=-3), 'إخفاء_فريق'),
    ],
    'C12': [
        ('C12-A', dict(C=13, M=16, R=3, P=17, B=-7), 'توسع_سريع'),
        ('C12-B', dict(D=9, K=7, R=10, M=7, B=5, P=3), 'استدامة'),
        ('C12-C', dict(K=14, B=11, D=5, M=-5, P=-10), 'تمكّن'),
    ],
}

ORDER = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12']

MITIGATION_FLAGS = {'صندوق_طوارئ', 'احتياطي_قوي'}  # C10 mitigation per table 20


def resolve_ending(stats, flags):
    K, D, C, B, R, M, P = stats['K'], stats['D'], stats['C'], stats['B'], stats['R'], stats['M'], stats['P']
    e1 = (R >= 60 and B >= 55 and D >= 55 and 'مسؤولية' in flags and
          ({'استدامة', 'تجربة_محدودة', 'إعادة_هيكلة'} & flags))
    if e1:
        return 'E1'
    e2 = (C >= 65 and M >= 55 and ({'توسع_سريع', 'شراكة_سريعة', 'دين'} & flags))
    if e2:
        return 'E2'
    e3 = (K >= 60 and D >= 55)
    if e3:
        return 'E3'
    return 'E4'


def simulate(choice_indices):
    stats = dict(START)
    flags = set()
    path = []
    for chapter, idx in zip(ORDER, choice_indices):
        code, delta, flag = CHAPTERS[chapter][idx]
        d = dict(delta)
        if chapter == 'C10' and (MITIGATION_FLAGS & flags):
            d['M'] = d.get('M', 0) + 4
            d['P'] = d.get('P', 0) - 5
        for k, v in d.items():
            stats[k] = clamp(stats[k] + v)
        if flag:
            flags.add(flag)
        path.append(code)
    ending = resolve_ending(stats, flags)
    return path, stats, flags, ending


def main():
    option_counts = [len(CHAPTERS[c]) for c in ORDER]
    total = 1
    for n in option_counts:
        total *= n

    ending_counts = defaultdict(int)
    ending_samples = {}
    overlap_e2_would_be_e4_special = 0  # R<40 & P>=65 but actually resolves to E2/E3
    e4_special_hits = 0
    e4_pure_default = 0
    max_seen = defaultdict(lambda: -1)
    min_seen = defaultdict(lambda: 999)

    for combo in itertools.product(*[range(n) for n in option_counts]):
        path, stats, flags, ending = simulate(combo)
        ending_counts[ending] += 1
        if ending not in ending_samples:
            ending_samples[ending] = (path, dict(stats), sorted(flags))
        if stats['R'] < 40 and stats['P'] >= 65:
            e4_special_hits += 1
            if ending != 'E4':
                overlap_e2_would_be_e4_special += 1
        if ending == 'E4' and not (stats['R'] < 40 and stats['P'] >= 65):
            e4_pure_default += 1
        for k in stats:
            max_seen[k] = max(max_seen[k], stats[k])
            min_seen[k] = min(min_seen[k], stats[k])

    print('TOTAL PATHS:', total)
    print()
    print('Ending distribution:')
    for e in ['E1', 'E2', 'E3', 'E4']:
        cnt = ending_counts.get(e, 0)
        print(f'  {e}: {cnt} paths ({cnt/total*100:.2f}%)')
    print()
    print('Reachability check: all endings reachable?', all(ending_counts.get(e, 0) > 0 for e in ['E1', 'E2', 'E3', 'E4']))
    print()
    print('Stat ranges achievable across all paths:')
    for k in ['K', 'D', 'C', 'B', 'R', 'M', 'P']:
        print(f'  {k}: min={min_seen[k]} max={max_seen[k]}')
    print()
    print('R<40 & P>=65 cases (the E4 "special" narrative condition):', e4_special_hits)
    print('  of those, how many actually resolved to something other than E4 (i.e. overridden by higher-priority ending):', overlap_e2_would_be_e4_special)
    print('E4 paths reached via pure default (no special condition met):', e4_pure_default, 'out of', ending_counts.get('E4', 0))
    print()
    print('Sample path per ending:')
    for e in ['E1', 'E2', 'E3', 'E4']:
        if e in ending_samples:
            path, stats, flags = ending_samples[e]
            print(f'  {e}: {" -> ".join(path)}')
            print(f'      stats: {stats}')
            print(f'      flags: {flags}')
        else:
            print(f'  {e}: NOT REACHABLE')
    print()

    # Check: is E1 reachable only via a very narrow set of choices? Count how many distinct "recipes" (last few key choices)
    # Which C-choices are necessary/sufficient contributors -> just report which chapters' choices appear in first E1 sample
    with open('/home/user/workspace/ending_simulation_report.json', 'w', encoding='utf-8') as f:
        json.dump({
            'total_paths': total,
            'ending_counts': dict(ending_counts),
            'ending_percentages': {e: round(ending_counts.get(e, 0) / total * 100, 3) for e in ['E1', 'E2', 'E3', 'E4']},
            'stat_ranges': {k: [min_seen[k], max_seen[k]] for k in ['K', 'D', 'C', 'B', 'R', 'M', 'P']},
            'e4_special_condition_hits': e4_special_hits,
            'e4_special_overridden_by_higher_priority': overlap_e2_would_be_e4_special,
            'e4_pure_default_count': e4_pure_default,
            'samples': {e: {'path': ending_samples[e][0], 'stats': ending_samples[e][1], 'flags': ending_samples[e][2]} for e in ending_samples},
        }, f, ensure_ascii=False, indent=2)

if __name__ == '__main__':
    main()
