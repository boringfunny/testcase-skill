"""Synthetic behavioral checks for v2; not actual game acceptance."""
import copy
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path
import openpyxl
from design_model import prepare
from build_documents import check_plan, build_xmind, kids, walk, notes, case_leaves, presentation_report


def exercise_v2(work):
    work.mkdir(parents=True)
    source = work/'requirement.txt'
    source.write_text('合成需求，非真实游戏：积分>=100可领一次奖励；VIP且活动开启可见入口；结束后补发规则未定。', encoding='utf-8')
    import hashlib
    evidence = {'sources': [{'path': str(source), 'category': 'requirement', 'sha256': hashlib.sha256(source.read_bytes()).hexdigest()}],
        'records': [{'id': 'R1', 'category': 'requirement', 'file': str(source), 'locator': '第1行', 'text': source.read_text(encoding='utf-8')}], 'images': []}
    methods = [('直接验证', {}), ('边界值', {'threshold': '100', 'unit': '整数积分', 'step': '1', 'selected_values': '100'}),
        ('决策表', {'conditions': 'VIP=true, 活动开启=true', 'outcome': '入口可见'}),
        ('状态迁移', {'from_state': '已领取', 'event': '重复领取', 'to_state': '已领取，无额外奖励'}),
        ('等价类', {'partition': '整数积分0至99', 'representative': '50'})]
    plan = {'schema_version': 2, 'title': 'v2合成验收', 'requirement_version': '合成v1', 'source_review': '单段合成文本，无图片修订',
        'rules': [], 'coverage_items': [], 'cases': [], 'questions': [], 'scope_notes': ['合成自测，没有适用通用思路表条目'], 'unknowns': [], 'failed_attempts': []}
    definitions = [('积分达到100可领取', '积分100且未领取', '领取奖励', '获得一次奖励', '奖励资产和领取标志'),
        ('积分阈值100', '积分100且未领取', '打开领奖界面', '达到领取条件', '按钮状态'),
        ('VIP且活动开启可见入口', 'VIP角色，活动开启', '打开活动列表', '入口可见', '活动列表'),
        ('奖励仅一次', '同一角色已领本档奖励', '再次请求领取', '不再增加奖励', '前后资产与领取标志'),
        ('积分未达100不可领取', '积分50且未领取', '尝试领取奖励', '不可领取，资产不变', '按钮与资产')]
    for n, ((method, model), (statement, pre, action, expected, observe)) in enumerate(zip(methods, definitions), 1):
        rid, vid, cid = f'RULE-{n:03}', f'COV-{n:03}', f'TC-{n:03}'
        plan['rules'].append({'id': rid, 'statement': statement, 'kind': '需求验证', 'evidence': ['R1'], 'rationale': '合成原文约束', 'questions': []})
        plan['coverage_items'].append({'id': vid, 'rule_id': rid, 'method': method, 'target': statement, 'rationale': '测试方法模型的合成验证', 'model': model, 'disposition': 'planned'})
        plan['cases'].append({'id': cid, 'path': ['合成业务'], 'title': statement, 'coverage_ids': [vid], 'preconditions': [pre],
            'test_data': [{'name': '角色状态', 'value': pre, 'basis': '合成夹具准备；没有执行'}],
            'steps': [{'action': action, 'expected': expected, 'observe': observe}], 'postconditions': ['恢复合成角色状态'],
            'impact': '检查资产、领取标志和界面一致性', 'priority': 'P0', 'priority_reason': '合成领取/资格规则',
            'verification': {'method': '运行实测', 'requirements': ['可控积分、角色资格和活动状态；未确认工具'], 'limitations': ['只检查设计导出，不执行游戏']},
            'kind': '需求验证', 'evidence': ['R1'], 'ideas': [], 'ideas_reason': '合成结构测试不使用真实思路', 'questions': []})
    q = {'id': 'Q01', 'title': '结束后补发', 'gap': '原文未定', 'ask': '结束后是否补发未领奖励', 'evidence': ['R1']}
    plan['questions'].append(q)
    plan['rules'].append({'id': 'RULE-006', 'statement': '结束后补发口径未定', 'kind': '待确认', 'evidence': ['R1'], 'rationale': '原文明确未定', 'questions': ['Q01']})
    plan['coverage_items'].append({'id': 'COV-006', 'rule_id': 'RULE-006', 'method': '风险经验', 'target': '结束后未领取', 'rationale': '待明确预期', 'disposition': 'planned'})
    c = copy.deepcopy(plan['cases'][0]); c.update(id='TC-006', title='结束后补发待确认', coverage_ids=['COV-006'], kind='待确认', questions=['Q01'],
        preconditions=['活动已结束且本档未领取'], test_data=[], steps=[{'action': '待口径确认后观察结束处理', 'expected': '待Q01批准后确定，不预设补发结果', 'observe': '实际可观察发放行为及资产'}])
    plan['cases'].append(c)
    for key, state, reason in [('007', 'deferred', '跨服环境本轮不具备，需后续设计确认'), ('008', 'excluded', '合成需求没有排行榜，不适用该方向')]:
        plan['coverage_items'].append({'id': 'COV-'+key, 'rule_id': 'RULE-001', 'method': '风险经验', 'target': reason, 'rationale': reason, 'disposition': state, 'reason': reason})
    prepared = prepare(plan, evidence); check_plan(prepared, evidence)
    assert prepared['_coverage']['planned'] == 6 and prepared['_coverage']['pending_planned'] == 1
    assert prepared['_coverage']['deferred'] == 1 and prepared['_coverage']['excluded'] == 1
    assert 'action' not in plan['cases'][0], 'prepare不得修改输入'
    rejections = 0
    def reject(edit):
        nonlocal rejections
        bad = copy.deepcopy(plan); edit(bad)
        try:
            check_plan(prepare(bad, evidence), evidence)
        except (ValueError, AssertionError):
            rejections += 1
        else:
            raise AssertionError('错误计划未被拒绝')
    reject(lambda p: p['cases'][0].update(coverage_ids=['COV-missing']))
    reject(lambda p: p['coverage_items'][0].update(rule_id='RULE-missing'))
    reject(lambda p: p['cases'].pop(0))
    reject(lambda p: p['coverage_items'][6].pop('reason'))
    reject(lambda p: p['coverage_items'][1]['model'].pop('step'))
    reject(lambda p: p['cases'][5].update(questions=[]))
    reject(lambda p: p['cases'][0]['steps'][0].pop('observe'))
    reject(lambda p: p['cases'][0].update(actual='未经执行的通过'))
    reject(lambda p: p['cases'][0].update(coverage_ids=['COV-007']))
    reject(lambda p: p['rules'][0].update(evidence=['R-missing']))
    reject(lambda p: p['rules'].append(copy.deepcopy(p['rules'][0])))
    reject(lambda p: p['cases'][0].pop('ideas_reason'))
    reject(lambda p: p['rules'][0].update(kind='边界派生'))
    reject(lambda p: p['cases'][0]['evidence'].append('R-missing'))
    baseline = prepared['_snapshot']
    # Keep the no-summary baseline to exercise backward-compatible fingerprints.
    legacy_plan = copy.deepcopy(plan)
    plan['cases'][0]['xmind_points'] = ['积分100且未领取时领取，获得一次奖励']
    plan['cases'][1]['xmind_points'] = ['积分恰好100时达到领取条件']
    plan['cases'][2]['xmind_points'] = ['VIP且活动开启时，活动列表显示入口']
    plan['cases'][3]['steps'].append({'action': '刷新后再次查看资产', 'expected': '仍不增加奖励', 'observe': '刷新后的资产与领取标志'})
    # Multi-step detail is part of this fixture's baseline, unlike its summaries.
    legacy_plan['cases'][3]['steps'] = copy.deepcopy(plan['cases'][3]['steps'])
    baseline = prepare(legacy_plan, evidence)['_snapshot']
    plan['cases'][3]['xmind_points'] = ['已领取后再次请求，奖励不增加', '刷新后仍不增加奖励']
    plan['cases'][4]['xmind_points'] = ['积分50时不可领取，资产不变']
    plan['cases'][5]['xmind_points'] = ['确认活动结束后是否补发未领奖励']
    for value in [[], [''], ['  '], '不是数组', [1], None]:
        reject(lambda p, v=value: p['cases'][0].update(xmind_points=v))
    same = prepare(plan, evidence, baseline=baseline)
    assert all(c['change'] == '未变化' for c in same['_changes'])
    rewritten = copy.deepcopy(plan); rewritten['cases'][0]['xmind_points'] = ['未领奖励且积分100时，可领取一次']
    assert prepare(rewritten, evidence)['_snapshot']['cases'] == baseline['cases']
    legacy_render = prepare(legacy_plan, evidence)
    refs, ideas, qs = check_plan(legacy_render, evidence)
    old_nodes, _ = build_xmind(legacy_render, refs, ideas, qs, work/'legacy-v2.xmind')
    assert len(presentation_report(legacy_render)['legacy_detail_cases']) == 6
    assert presentation_report(legacy_render)['warnings']
    assert kids(old_nodes['TC-006'])[-1]['title'] == '待确认：Q01'
    mixed = copy.deepcopy(plan); mixed['cases'][1].pop('xmind_points')
    mixed = prepare(mixed, evidence)
    mixed_nodes, _ = build_xmind(mixed, refs, ideas, qs, work/'mixed.xmind')
    assert kids(mixed_nodes['TC-002'])[0]['title'].startswith('条件/操作：')
    assert [t['title'] for t in kids(mixed_nodes['TC-001'])] == plan['cases'][0]['xmind_points']
    assert presentation_report(mixed)['legacy_detail_cases'] == ['TC-002']
    changed = copy.deepcopy(plan); changed['rules'][0]['statement'] = '合成修改：积分阈值改为200'
    compared = prepare(changed, evidence, baseline=baseline)
    assert next(c for c in compared['_changes'] if c['id'] == 'TC-001')['change'] == '需重新验证'
    assert next(c for c in compared['_changes'] if c['id'] == 'TC-002')['change'] == '未变化'
    removed = copy.deepcopy(plan); removed['cases'].pop(0); removed['coverage_items'][0].update(disposition='deferred', reason='本轮暂缓')
    assert next(c for c in prepare(removed, evidence, baseline=baseline)['_changes'] if c['id'] == 'TC-001')['change'] == '移除'
    legacy = {'title': '旧格式', 'cases': []}
    try:
        prepare(legacy, evidence)
    except ValueError:
        rejections += 1
    else:
        raise AssertionError('旧格式必须显式选择')
    assert '_coverage' not in prepare(legacy, evidence, allow_legacy=True)
    pp, ep, bp = work/'plan.json', work/'evidence.json', work/'baseline.json'
    for path, value in [(pp, plan), (ep, evidence), (bp, baseline)]:
        path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    cmd = [sys.executable, str(Path(__file__).with_name('build_documents.py')), '--plan', str(pp), '--evidence', str(ep), '--baseline', str(bp), '--out', str(work/'output')]
    result = subprocess.run(cmd, capture_output=True, encoding='utf-8', env=dict(os.environ, PYTHONIOENCODING='utf-8'))
    assert result.returncode == 0, (result.stdout, result.stderr)
    workbook = openpyxl.load_workbook(work/'output'/'v2合成验收-AI验收表.xlsx', data_only=True)
    from self_test_status import check_saved_workbook
    check_saved_workbook(work/'output'/'v2合成验收-AI验收表.xlsx',6)
    assert len(workbook.sheetnames) == 7
    assert workbook['规则与设计覆盖'].max_row == 13
    main = workbook['AI实现检查']
    assert main.cell(11,6).value == '【待确认】'
    assert all(main.cell(row, col).value is None for row in range(6,12) for col in [7,8])
    saved = json.loads((work/'output'/'.support/design-baseline.json').read_text(encoding='utf-8'))
    assert saved['cases'] == baseline['cases']
    with zipfile.ZipFile(work/'output'/'v2合成验收-测试用例.xmind') as z:
        sheets = json.loads(z.read('content.json'))
    exported = {t['labels'][0]: t for t in walk(sheets[0]['rootTopic']) if t.get('labels')}
    compact = prepare(plan, evidence)
    for row, c in enumerate(compact['cases'], 6):
        t = exported[c['id']]
        assert [k['title'] for k in kids(t)] == c['xmind_points']
        assert all(not kids(k) for k in kids(t)), '简述不得再有固定分类分支'
        assert t['labels'] == [c['id'], c['priority'], c['kind']]
        full = t['notes']['plain']['content']
        for value in [c['action'], c['expected'], c['impact'], *c['preconditions'], *c['postconditions']]:
            assert value in full
        assert full == workbook['检查详情'].cell(row,11).value
        assert workbook['检查详情'].cell(row,3).value == c['action'] and workbook['检查详情'].cell(row,4).value == c['expected']
    assert exported['TC-006']['title'].endswith('【待确认 Q01】')
    qnode = kids(sheets[1]['rootTopic'])[0]
    assert [k['title'].split('：')[0] for k in kids(qnode)] == ['原文/缺口','需确认','影响用例']
    assert 'TC-006' in kids(qnode)[2]['title']
    info_titles = [t['title'] for t in walk(sheets[2]['rootTopic'])]
    assert not any(t in ['规则与设计覆盖','执行准备','变更影响'] for t in info_titles)
    return {'cases': 6, 'rules': 6, 'coverage_items': 8, 'rejected_invalid_designs': rejections,
        'checked': ['all five structured design methods', 'pending propagation', 'deferred/excluded reasons', 'immutable input', 'rule-only change invalidates case', 'unchanged/removed cases', 'explicit legacy mode', 'seven sheets', 'all execution columns empty', 'baseline roundtrip', 'ordinary/boundary/multistep/pending compact cases', 'complete notes and Excel detail', 'Q sheet retained', 'summary-only fingerprint unchanged', 'old V2 and mixed fallback', 'invalid summaries rejected', 'JSON/XML hierarchy and labels']}
