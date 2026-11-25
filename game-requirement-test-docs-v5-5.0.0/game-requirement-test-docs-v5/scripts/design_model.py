"""Validate design v2, render execution instructions, and compare design baselines.

Structural coverage is not proof of semantic completeness or executed coverage.
"""
import copy
import hashlib
import json
from review_design import validate_focus, ordered_cases, focus_text
from presentation_model import validate_display, enabled, ordered_display_cases

KINDS = {'需求验证', '边界派生', '补充建议', '待确认'}
METHODS = {'直接验证', '等价类', '边界值', '决策表', '状态迁移', '风险经验'}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def text(value):
    return isinstance(value, str) and bool(value.strip())


def strings(value, nonempty=True):
    return isinstance(value, list) and (bool(value) or not nonempty) and all(text(x) for x in value)


def index(items, prefix):
    require(isinstance(items, list), f'{prefix}: 必须是数组')
    result = {}
    for item in items:
        require(isinstance(item, dict) and text(item.get('id')), f'{prefix}: 缺少ID')
        require(item['id'] not in result, f'{prefix}: 重复ID {item["id"]}')
        result[item['id']] = item
    return result


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def prepare(plan, evidence, allow_legacy=False, baseline=None):
    """Return a separate rendered plan; never modify the input or reuse execution results."""
    p = copy.deepcopy(plan)
    validate_display(p)
    if p.get('schema_version') != 2:
        require(allow_legacy and 'schema_version' not in p,
                '新设计必须使用schema_version=2；旧导图转表须显式--allow-legacy，不自动补造规则')
        require(baseline is None, '旧格式不支持基线比较；先人工补齐v2设计')
        p.setdefault('scope_notes', []).append('旧格式保真转换：未建立v2规则覆盖模型，不报告设计覆盖率。')
        return p
    require(text(p.get('requirement_version')), '缺少需求版本/文件指纹描述')
    require(text(p.get('source_review')), '缺少正文、图片、修订读取范围和未读部分说明')
    validate_focus(p)
    rules = index(p.get('rules'), 'rules')
    items = index(p.get('coverage_items'), 'coverage_items')
    cases = index(p.get('cases'), 'cases')
    qs = index(p.get('questions', []), 'questions')
    records = index(evidence.get('records'), 'evidence')
    require(rules and items and cases, '规则、覆盖项和用例均不能为空')
    for rule in rules.values():
        rid = rule['id']
        require(text(rule.get('statement')) and rule.get('kind') in KINDS, f'{rid}: 缺规则或性质')
        require(strings(rule.get('evidence')), f'{rid}: 缺少原文证据')
        require(all(e in records and records[e]['category'] == 'requirement' for e in rule['evidence']), f'{rid}: 无效需求证据')
        require(strings(rule.get('questions', []), False) and all(q in qs for q in rule.get('questions', [])), f'{rid}: 无效Q')
        require((rule['kind'] == '待确认') == bool(rule.get('questions')), f'{rid}: 待确认性质与Q不一致')
        require(text(rule.get('rationale')), f'{rid}: 缺少原文支持/推导/建议理由')
    for item in items.values():
        cid = item['id']
        require(item.get('rule_id') in rules, f'{cid}: 无效规则引用')
        require(item.get('method') in METHODS, f'{cid}: 未知设计方法')
        require(text(item.get('target')) and text(item.get('rationale')), f'{cid}: 缺覆盖目标或选取理由')
        require(item.get('disposition') in {'planned', 'deferred', 'excluded'}, f'{cid}: 无效设计处置')
        if item['disposition'] != 'planned':
            require(text(item.get('reason')), f'{cid}: 暂缓/排除必须说明原因')
        if item['method'] in {'等价类', '边界值', '决策表', '状态迁移'}:
            model = item.get('model', {})
            fields = {'等价类': ['partition', 'representative'],
                      '边界值': ['threshold', 'unit', 'step', 'selected_values'],
                      '决策表': ['conditions', 'outcome'],
                      '状态迁移': ['from_state', 'event', 'to_state']}[item['method']]
            require(isinstance(model, dict) and all(text(model.get(k)) for k in fields), f'{cid}: 缺方法模型 {fields}')
    links = {i: [] for i in items}
    for c in cases.values():
        cid = c['id']
        if 'xmind_points' in c:
            require(strings(c['xmind_points']), f'{cid}: xmind_points须为非空字符串数组；省略该字段可保留旧展示')
        require(strings(c.get('coverage_ids')) and len(c['coverage_ids']) == len(set(c['coverage_ids'])), f'{cid}: 覆盖项引用为空或重复')
        require(all(i in items and items[i]['disposition'] == 'planned' for i in c['coverage_ids']), f'{cid}: 无效覆盖项，或引用了排除/暂缓项')
        rule_ids = sorted({items[i]['rule_id'] for i in c['coverage_ids']})
        c['rule_ids'] = rule_ids
        required_q = {q for rid in rule_ids for q in rules[rid].get('questions', [])}
        require(required_q <= set(c.get('questions', [])), f'{cid}: 遗漏关联规则的待确认问题')
        required_e = {e for rid in rule_ids for e in rules[rid]['evidence']}
        require(required_e <= set(c.get('evidence', [])), f'{cid}: 遗漏关联规则的证据')
        require(all(e in records and records[e]['category']=='requirement' for e in c.get('evidence', [])), f'{cid}: 无效用例证据')
        require(all(q in qs for q in c.get('questions', [])), f'{cid}: 无效用例Q')
        require(strings(c.get('preconditions')), f'{cid}: 缺前置条件；无特殊前置也须明示')
        require(isinstance(c.get('test_data'), list), f'{cid}: test_data须为数组；无需数据用空数组')
        for d in c['test_data']:
            require(isinstance(d, dict) and all(text(d.get(k)) for k in ['name', 'value', 'basis']), f'{cid}: 数据缺名称、值或来源/构造说明')
        require(isinstance(c.get('steps'), list) and c['steps'], f'{cid}: 缺步骤')
        for s in c['steps']:
            require(isinstance(s, dict) and all(text(s.get(k)) for k in ['action', 'expected', 'observe']), f'{cid}: 步骤缺操作、预期或观察位置/时点')
        require(strings(c.get('postconditions')), f'{cid}: 缺结束状态/清理要求')
        require(text(c.get('priority_reason')), f'{cid}: 缺优先级理由')
        v = c.get('verification', {})
        require(isinstance(v, dict) and text(v.get('method')) and strings(v.get('requirements')) and strings(v.get('limitations'), False), f'{cid}: 缺建议验证方式、能力需求或限制说明')
        require(c.get('kind') in KINDS, f'{cid}: 无效性质')
        if c['kind']=='需求验证':
            require(all(rules[r]['kind']=='需求验证' for r in rule_ids), f'{cid}: 推导/建议/未知规则不能标原文明示')
        require((c['kind'] == '待确认') == bool(c.get('questions')), f'{cid}: 待确认与Q不一致')
        if not c.get('ideas'):
            require(text(c.get('ideas_reason')), f'{cid}: 无适用思路须逐条说明')
        for key in ['actual', 'result', 'status', 'current_code', 'acceptance_status']:
            require(not c.get(key), f'{cid}: 新设计不得预填执行结果 {key}')
        for i in c['coverage_ids']:
            links[i].append(cid)
        c['action'] = '\n'.join(['前置：' + '；'.join(c['preconditions']),
            '数据：' + ('；'.join(d['name'] + '=' + d['value'] + '（' + d['basis'] + '）' for d in c['test_data']) or '无需额外数据')]
            + [f'{n}. {s["action"]}' for n, s in enumerate(c['steps'], 1)]
            + ['结束/清理：' + '；'.join(c['postconditions'])])
        c['expected'] = '\n'.join(f'{n}. {s["expected"]}；观察：{s["observe"]}' for n, s in enumerate(c['steps'], 1))
    for i, item in items.items():
        require(item['disposition'] != 'planned' or links[i], f'{i}: 计划覆盖项没有用例')
    for rid in rules:
        require(any(i['rule_id'] == rid for i in items.values()), f'{rid}: 无覆盖项；不适用也须写排除原因')
    counts = {s: sum(i['disposition'] == s for i in items.values()) for s in ['planned', 'deferred', 'excluded']}
    counts['pending_planned'] = sum(i['disposition'] == 'planned' and (rules[i['rule_id']]['kind'] == '待确认' or any(cases[c]['kind']=='待确认' for c in links[i['id']])) for i in items.values())
    p['_coverage'] = {'rules': len(rules), 'items': len(items), **counts, 'links': links,
        'scope': '仅已登记规则/覆盖项的设计映射；不证明规则提取完整、语义正确或已执行。'}
    # Source text (rather than extraction sequence IDs) participates in semantic fingerprints.
    rule_hashes = {rid: digest({'statement': r['statement'], 'kind': r['kind'], 'rationale': r['rationale'],
        'questions': [qs[q] for q in r.get('questions', [])],
        'source_text': [records[e]['text'] for e in r['evidence']]}) for rid, r in rules.items()}
    case_hashes = {}
    for cid, c in cases.items():
        # Display summaries do not change the executable business design.
        design = {k: v for k, v in c.items() if k not in {'evidence', 'ideas', 'action', 'expected', 'xmind_points', 'acceptance_path', 'display_group', 'display_path'}}
        case_hashes[cid] = digest({'case': design, 'rules': {r: rule_hashes[r] for r in c['rule_ids']},
            'coverage': [items[i] for i in c['coverage_ids']], 'questions': [qs[q] for q in c.get('questions', [])],
            'source_text': [records[e]['text'] for e in c.get('evidence', [])]})
    p['_snapshot'] = {'snapshot_version': 1, 'title': p['title'], 'requirement_version': p['requirement_version'],
        'rules': rule_hashes, 'cases': case_hashes, 'source_hashes': evidence.get('sources', [])}
    old = {}
    if baseline is not None:
        require(baseline.get('snapshot_version') == 1 and isinstance(baseline.get('cases'), dict) and isinstance(baseline.get('rules'), dict), '不支持的设计基线')
        require(baseline.get('title') == p['title'], '基线标题不匹配；请先核实是否属于同一需求')
        old = baseline['cases']
    p['_changes'] = []
    for cid in sorted(set(old) | set(case_hashes)):
        state = ('首次设计' if baseline is None else '新增') if cid not in old else '移除' if cid not in case_hashes else '未变化' if old[cid] == case_hashes[cid] else '需重新验证'
        p['_changes'].append({'id': cid, 'change': state,
            'action': '保留历史记录，不在本次验收表执行' if state == '移除' else '本次结果留空；是否复用旧证据须另行核实环境及有效性' if state == '未变化' else '本次结果留空，按当前设计执行'})
    p['cases'] = ordered_display_cases(p['cases']) if enabled(p) else ordered_cases(p['cases'])
    return p


def tables(plan):
    """Structured read-only sheets: no duplicate execution result columns."""
    if plan.get('schema_version') != 2:
        return []
    coverage = plan['_coverage']
    rules = {r['id']: r for r in plan['rules']}
    rows = [[i['id'], i['rule_id'], rules[i['rule_id']]['statement'], rules[i['rule_id']]['kind'],
        i['method'], i['target'], i['rationale'], json.dumps(i.get('model', {}), ensure_ascii=False),
        i['disposition'], i.get('reason', ''), '、'.join(coverage['links'][i['id']]),
        '、'.join(rules[i['rule_id']]['evidence']), '、'.join(rules[i['rule_id']].get('questions', [])), rules[i['rule_id']]['rationale']] for i in plan['coverage_items']]
    prep = [[c['id'], '、'.join(c['rule_ids']), '、'.join(c['coverage_ids']), c['priority_reason'],
        '\n'.join(c['preconditions']), '\n'.join(d['name'] + '=' + d['value'] + '；' + d['basis'] for d in c['test_data']),
        c['verification']['method'], '\n'.join(c['verification']['requirements']) + ('\n\n'+focus_text(c) if c.get('review_focus') else ''),
        '\n'.join(c['verification']['limitations']) or '本次未列额外限制；不代表已具备执行条件', '\n'.join(c['postconditions'])] for c in plan['cases']]
    changes = [[c['id'], c['change'], c['action']] for c in plan['_changes']]
    return [
        ('规则与设计覆盖', ['覆盖项ID', '规则ID', '规则内容', '规则性质', '设计方法', '覆盖目标', '选择理由', '方法模型', '设计处置', '暂缓或排除原因', '用例ID', '需求证据ID', '问题ID', '规则依据或推导说明'], [110,100,420,130,110,340,360,400,120,350,240,180,120,420], rows),
        ('执行准备', ['用例ID', '规则ID', '覆盖项ID', '优先级理由', '前置条件', '测试数据与构造依据', '建议验证方式', '所需能力与环境', '限制或待准备', '结束与清理'], [100,160,180,320,350,420,180,400,400,350], prep),
        ('变更影响', ['用例ID', '设计变化', '本次处理'], [110,160,650], changes)]
