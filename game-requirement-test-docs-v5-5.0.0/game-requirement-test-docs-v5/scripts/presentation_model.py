"""V5 display-only classification; never infer rules or overwrite business paths."""
import hashlib
import json
import re
from pathlib import Path

GROUPS = ('入口', '红点', '主功能', '特殊点验证', '异常状态', '配置检查', 'UI检查')
PROFILES = ('活动', '玩法', '养成', '英雄', '通用')
PLACEHOLDER = re.compile(r'^(?:主功能|次级模块|功能点|分支主题|细分主题|中心主题)[\d一二三四五六七八九十]*$')


def enabled(plan):
    return plan.get('schema_version') == 2 and plan.get('template_profile') in PROFILES


def validate_display(plan, strict=False):
    from review_design import validate_focus
    present = 'template_profile' in plan or any('display_group' in c or 'display_path' in c for c in plan.get('cases', []))
    if not present and not strict:
        return
    validate_focus(plan, strict)
    if plan.get('schema_version') != 2 or plan.get('template_profile') not in PROFILES:
        raise ValueError('V5展示须使用schema_version=2和有效template_profile；旧导图不套用模板')
    for c in plan.get('cases', []):
        group, path = c.get('display_group'), c.get('display_path')
        if not isinstance(group, str) or group not in GROUPS or not isinstance(path, list) or not path or not all(isinstance(s, str) and s.strip() == s and s for s in path):
            raise ValueError(f'{c["id"]}: 须同时填写有效display_group及非空display_path；部分配置不能自动猜测')
        if any(PLACEHOLDER.fullmatch(s) for s in path):
            raise ValueError(f'{c["id"]}: 展示路径不能使用模板占位名称')
        if group == '主功能':
            if path[0] in GROUPS:
                raise ValueError(f'{c["id"]}: 主功能路径须以具体业务名称开头')
        elif path[0] != group:
            raise ValueError(f'{c["id"]}: 展示路径须以{group}开头')
        if c.get('review_focus', {}).get('target') == '客户端简查' and group != 'UI检查':
            raise ValueError(f'{c["id"]}: 纯客户端简查须归入UI检查')
        if strict and not c.get('xmind_points'):
            raise ValueError(f'{c["id"]}: V5新设计须人工编写xmind_points')
    decisions = plan.get('template_decisions', {})
    if not isinstance(decisions, dict) or any(k not in GROUPS or not isinstance(v, str) or not v.strip() for k, v in decisions.items()):
        raise ValueError('template_decisions须为分类名称到非空取舍理由的字典')


def display_module(c):
    return c.get('display_path', c['path'])[0]


def ordered_display_cases(cases):
    """Traverse the same ordered path tree as XMind, including interleaved subpaths."""
    ranked = sorted(cases, key=lambda c: GROUPS.index(c['display_group']))
    tree = {}
    for c in ranked:
        branch = tree
        for seg in c['display_path']:
            branch = branch.setdefault(('group', seg), {})
        branch[('case', c['id'])] = c
    def leaves(branch):
        for key, value in branch.items():
            if key[0] == 'case':
                yield value
            else:
                yield from leaves(value)
    return list(leaves(tree))


def template_report(plan):
    if not enabled(plan):
        return {'mode': 'legacy', 'warnings': ['旧计划缺少V5展示配置，保留原组织方式；可人工补充后再启用模板。']}
    root = Path(__file__).resolve().parents[1]
    manifest = json.loads((root/'assets/format-templates/manifest.json').read_text(encoding='utf-8'))
    for item in manifest['templates']:
        path = root/item['path']
        if hashlib.sha256(path.read_bytes()).hexdigest() != item['sha256']:
            raise ValueError(f'模板快照指纹不符：{item["path"]}')
    used = {c['display_group'] for c in plan['cases']}
    decisions = plan.get('template_decisions', {})
    return {'mode': 'v5', 'profile': plan['template_profile'], 'template_snapshots': manifest['templates'],
            'modules': list(dict.fromkeys(display_module(c) for c in plan['cases'])),
            'category_decisions': {g: {'shown': g in used, 'cases': [c['id'] for c in plan['cases'] if c['display_group'] == g],
                'reason': decisions.get(g, '已分配实际用例' if g in used else '本计划未分配用例；适用性理由未提供，需设计者复核')} for g in GROUPS},
            'warnings': [f'{g}未记录设计取舍理由' for g in GROUPS if g not in decisions],
            'boundary': '模板只提供组织与样式；模板示例不构成当前业务规则。'}
