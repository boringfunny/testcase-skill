"""V3 design intent; no source inspection or acceptance execution."""
TARGETS = {'服务端', '客户端简查'}


def validate_focus(plan, strict=False):
    if strict and plan.get('schema_version') != 2:
        raise ValueError('V3新设计须使用schema_version=2；旧图保真转换不使用--require-v3')
    for c in plan.get('cases', []):
        if 'acceptance_path' in c and not (isinstance(c['acceptance_path'],str) and c['acceptance_path'].strip()):
            raise ValueError(f'{c["id"]}: acceptance_path须为非空字符串')
        if strict and 'acceptance_path' not in c:
            raise ValueError(f'{c["id"]}: V3新设计缺少最短验收路径acceptance_path')
        focus = c.get('review_focus')
        if 'review_focus' in c:
            if not (isinstance(focus, dict) and isinstance(focus.get('target'),str) and focus['target'] in TARGETS and
                    isinstance(focus.get('checks'), list) and focus['checks'] and
                    all(isinstance(s, str) and s.strip() for s in focus['checks'])):
                raise ValueError(f'{c["id"]}: review_focus须含服务端/客户端简查target与非空checks字符串数组')
        elif strict:
            raise ValueError(f'{c["id"]}: V3新设计缺少review_focus')
    for q in plan.get('questions', []):
        if 'impact' in q and not (isinstance(q['impact'], str) and q['impact'].strip()):
            raise ValueError(f'{q["id"]}: impact须为非空的通俗影响说明')
        if strict and 'impact' not in q:
            raise ValueError(f'{q["id"]}: V3新问题缺少impact')


def focus_text(c):
    f = c.get('review_focus')
    if not f: return ''
    return '检查目标：' + f['target'] + '\n核对项：\n' + '\n'.join(f'{n}. {s}' for n,s in enumerate(f['checks'],1))


def detail_action(c):
    return c['action'] + ('\n\n' + focus_text(c) if c.get('review_focus') else '')


def output_path(c):
    if 'display_path' in c: return c['display_path']
    path = c['path']
    if c.get('review_focus',{}).get('target') == '客户端简查' and (len(path)<2 or path[1]!='客户端简查'):
        return [path[0], '客户端简查', *path[1:]]
    return path


def ordered_cases(cases):
    # Presentation only. Preserve IDs and input path; do not classify old cases.
    if not any('review_focus' in c for c in cases): return list(cases)
    modules = list(dict.fromkeys(c['path'][0] for c in cases))
    return sorted(cases, key=lambda c:(modules.index(c['path'][0]),c.get('review_focus',{}).get('target')=='客户端简查'))


def review_report(plan):
    return {
        'targets': {t:[c['id'] for c in plan['cases'] if c.get('review_focus',{}).get('target')==t] for t in sorted(TARGETS)},
        'cases_without_acceptance_path':[c['id'] for c in plan['cases'] if 'acceptance_path' not in c],
        'unclassified_legacy_cases':[c['id'] for c in plan['cases'] if 'review_focus' not in c],
        'questions_without_impact':[q['id'] for q in plan.get('questions',[]) if 'impact' not in q],
        'scope':'设计目标与引用检查；不证明实现正确或语义覆盖完整。',
    }
