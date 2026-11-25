"""Synthetic server-review design and compatibility fixtures, never real acceptance."""
import copy
import json
import os
from pathlib import Path
import subprocess
import sys
import zipfile
import openpyxl
from design_model import prepare
from review_design import validate_focus, detail_action, output_path
from build_documents import check_plan, walk, kids

HERE=Path(__file__).resolve().parent


def run(script,*args,ok=True):
    proc=subprocess.run([sys.executable,str(HERE/script),*map(str,args)],capture_output=True,encoding='utf-8',env=dict(os.environ,PYTHONDONTWRITEBYTECODE='1',PYTHONIOENCODING='utf-8'))
    assert (proc.returncode==0)==ok,(proc.stdout,proc.stderr)
    return proc


def fixture(work):
    work.mkdir(parents=True)
    source=work/'合成需求.txt'
    source.write_text('''合成需求：仅用于技能对比，不代表真实游戏。
活动开启且VIP时可以领奖；不满足任一条件不可领取。
积分按整数计，阈值配置为100；积分达到100且本角色本期未领取时，发放1张奖励券并保存领取状态。
积分不足100时不发奖、不改变领取状态；积分超过100仍只发1张奖励券。同一角色本期重复领取不再发奖。
积分是否跨天保留未明确，需确认；未获批准不假定清零或累计。
客户端说明文字为“100积分可领奖”。领取成功播放奖励动画。
''',encoding='utf-8')
    ep=work/'evidence.json'
    run('read_sources.py','--workspace',work,'--requirement',source,'--ideas',HERE.parent/'assets/通用游戏测试思路.xlsx','--out',ep)
    ev=json.loads(ep.read_text(encoding='utf-8'))
    refs=[r['id'] for r in ev['records'] if r['category']=='requirement']
    def rule(n,statement,q=False):
        return dict(id=f'RULE-{n:03}',statement=statement,kind='待确认' if q else '需求验证',evidence=refs,rationale='合成需求原文明示',questions=['Q01'] if q else [])
    rules=[rule(1,'积分阈值为100，符合资格且未领取时发1张奖励券并保存领取状态；不足不发，超出不加量'),
           rule(2,'活动开启且VIP才可领取'),rule(3,'同一角色本期重复领取不再发奖'),
           rule(4,'积分是否跨天保留未明确',True),rule(5,'客户端说明为“100积分可领奖”'),rule(6,'领取成功播放奖励动画')]
    # id,rule,title,pre,action,expected,ideas,method,model
    defs=[
      (1,1,'积分阈值与奖励配置','本期活动配置可读取','核对阈值、奖励类型与数量','阈值100，奖励为1张奖励券',['COM-021'],'直接验证',{}),
      (2,1,'积分99时领取','VIP、活动开启、未领取、积分99','尝试领取','不发奖励，领取状态不变',['COM-010'],'边界值',dict(threshold='100',unit='积分',step='1',selected_values='99')),
      (3,1,'积分100时首次领取','VIP、活动开启、未领取、积分100','领取奖励','发1张奖励券并保存领取状态',['COM-010','COM-030'],'边界值',dict(threshold='100',unit='积分',step='1',selected_values='100')),
      (4,1,'积分101时领取','VIP、活动开启、未领取、积分101','领取奖励','仍只发1张奖励券并保存领取状态',['COM-010'],'边界值',dict(threshold='100',unit='积分',step='1',selected_values='101')),
      (5,2,'非VIP、活动开启','非VIP、活动开启、积分100、未领取','尝试领取','不可领取，不发奖励',['COM-003'],'决策表',dict(conditions='VIP=false，活动开启=true',outcome='拒绝领取')),
      (6,2,'VIP、活动关闭','VIP、活动关闭、积分100、未领取','尝试领取','不可领取，不发奖励',['ACT-001'],'决策表',dict(conditions='VIP=true，活动开启=false',outcome='拒绝领取')),
      (7,2,'非VIP、活动关闭','非VIP、活动关闭、积分100、未领取','尝试领取','不可领取，不发奖励',['COM-003'],'决策表',dict(conditions='VIP=false，活动开启=false',outcome='拒绝领取')),
      (8,3,'同一期重复领取','本角色本期已领取','再次请求领取','不再发奖，领取状态保持已领取',['COM-013','ACT-008'],'状态迁移',dict(from_state='本期已领取',event='再次领取',to_state='已领取且奖励不增加')),
      (9,4,'跨天累计口径','前一天有积分且未领取','跨天后查看积分和领取条件','待Q01明确后判定，不预设清零或保留',['COM-015','COM-016'],'风险经验',{}),
      (10,5,'说明文字','活动页面可访问','查看说明文字','显示“100积分可领奖”',['COM-024'],'直接验证',{}),
      (11,6,'领取动画','本次领取成功','观察奖励动画','播放奖励动画',['COM-024'],'直接验证',{}),
    ]
    base=dict(schema_version=2,title='合成积分领奖',requirement_version='synthetic-v1',source_review='已读取合成文本全部6行，无图片与修订',
              rules=rules,coverage_items=[],cases=[],questions=[dict(id='Q01',title='跨天累计口径',gap='积分跨日规则未明确',ask='确认跨日累计口径',evidence=refs)],
              scope_notes=['合成对比，所有检查均未执行；按随包通用思路选择适用条目，不代表真实需求质量评测。'],unknowns=[],failed_attempts=[])
    for n,rid,title,pre,act,expected,ideas,method,model in defs:
        cov=f'COV-{n:03}';q=rid==4
        base['coverage_items'].append(dict(id=cov,rule_id=f'RULE-{rid:03}',method=method,target=title,rationale='按合成规则及所引用通用思路选取该场景',model=model,disposition='planned'))
        base['cases'].append(dict(id=f'TC-{n:03}',path=['积分领奖'],title=title,coverage_ids=[cov],preconditions=[pre],
            test_data=[dict(name='角色/活动状态',value=pre,basis='依据合成需求准备；未创建实际账号')],
            steps=[dict(action=act,expected=expected,observe='配置原值' if n==1 else '操作后的界面、奖励资产和领取状态' if n<10 else '对应画面')],
            postconditions=['未执行，无需实际清理；后续运行请恢复合成测试数据'],impact='资格、资产与领取状态按需求一致' if n<10 else '保持说明与反馈可理解',
            priority='P1' if n>=9 else 'P0',priority_reason='需求不明或客户端反馈' if n>=9 else '资格与奖励资产',
            verification=dict(method='配置核对' if n==1 else '运行实测',requirements=['本期配置与需求版本' if n==1 else '可控制资格、积分与活动状态的测试环境'],limitations=['本次只设计文档']),
            kind='待确认' if q else '需求验证',evidence=refs,ideas=ideas,questions=['Q01'] if q else [],xmind_points=[expected]))
    modern=copy.deepcopy(base)
    modern['questions'][0].update(title='积分是否跨天保留？',gap='需求说明积分达到100可以领取，没有说明跨天后积分是否保留。',
        ask='假设示例：昨天有60积分，今天再获得40积分，是否可以领取100积分奖励？',impact='影响跨天后的达标判断和奖励领取。')
    for c in modern['cases']:
        n=int(c['id'].split('-')[1]);client=n>=10
        c['review_focus']=dict(target='客户端简查' if client else '服务端',checks=[])
        if client:
            c['review_focus']['checks']=[c['steps'][0]['expected']+'；仅有资源或代码不能证明实际画面效果']
            c['verification']['requirements']=['客户端运行画面或可定位的效果证据；资源配置只支持静态部分']
            continue
        c['verification']=dict(method='配置核对' if n==1 else '静态代码检查' if n==9 else '代码与配置核对',
            requirements=['本次需求对应的源码/配置版本及可定位证据；实际文件、字段、函数由检查AI查证'],limitations=['未取得真实项目材料；本次不执行检查'])
        expected=c['steps'][0]['expected']
        checks=([expected+'，核对配置中的业务值和单位'] if n==1 else
                ['核对积分跨天的需求约定是否已批准','批准后核对跨天积分处理及领取资格判断；不以代码现状代替批准'] if n==9 else
                ['核对入口至领取处理的条件：'+c['preconditions'][0],
                 '追踪对应分支：'+expected,
                 '确认状态与资产更新覆盖本角色本期；记录实际配置、关键条件及结果证据'])
        c['review_focus']['checks']=checks
        c['steps']=[dict(action='核对配置业务值' if n==1 else '批准规则后追踪跨天积分处理' if n==9 else '追踪领取处理并核对条件与对应分支',
            expected=expected,observe='配置值与版本证据' if n==1 else '实际调用关系、条件分支与资产/状态处理证据；定位由检查AI填写')]
        if n==9:c.update(title='积分是否跨天保留？',xmind_points=['确认昨天的积分能否与今天的积分合并达标'])
    paths=[
        '需准备符合资格且100积分的未领奖角色 → 领取奖励 → 检查奖励增加1张；游戏内抽查不能替代配置逐项核对',
        '需准备VIP、活动开启、99积分未领奖角色 → 尝试领取 → 检查未发奖、领取状态不变',
        '需准备VIP、活动开启、100积分未领奖角色 → 领取 → 检查奖励增加1张、状态已领取',
        '需准备VIP、活动开启、101积分未领奖角色 → 领取 → 检查仍只增加1张奖励券',
        '需准备非VIP、活动开启、100积分角色 → 尝试领取 → 检查不能领奖；入口无法操作时需请求工具补验服务端拒绝',
        '需准备VIP、活动关闭、100积分角色 → 尝试领取 → 检查不能领奖；入口关闭时需请求工具补验服务端拒绝',
        '需准备非VIP、活动关闭、100积分角色 → 尝试领取 → 检查不能领奖；无法操作时需请求工具补验',
        '领取1次 → 再次尝试领取并重登 → 检查奖励只增加1张；按钮不可操作时需请求工具验证重复请求',
        '需准备昨天60积分、未领奖角色 → 跨天后增加40积分 → 记录总积分及领奖状态，待Q01确认后判定',
        '进入活动页面 → 查看说明 → 检查文字为“100积分可领奖”',
        '需准备可领奖角色 → 成功领取 → 观察是否播放奖励动画',
    ]
    for c,path in zip(modern['cases'],paths):c['acceptance_path']=path
    # Deliberate input disorder exercises display-only ordering without renumbering.
    modern['cases']=[modern['cases'][9],*modern['cases'][:9],modern['cases'][10]]
    for name,p in [('v2-plan.json',base),('v3-plan.json',modern)]:
        (work/name).write_text(json.dumps(p,ensure_ascii=False,indent=2),encoding='utf-8')
    return base,modern,ev


def exercise_v3(work):
    base,modern,ev=fixture(work)
    p=prepare(modern,ev);validate_focus(p,True);check_plan(p,ev)
    assert [c['id'] for c in modern['cases']][0]=='TC-010','input mutated'
    assert [c['id'] for c in p['cases']]==[f'TC-{i:03}' for i in range(1,12)]
    assert prepare(base,ev)['_snapshot']['cases']==prepare(base,ev,baseline=prepare(base,ev)['_snapshot'])['_snapshot']['cases']
    checks=0
    def reject(edit,strict=True):
        nonlocal checks
        bad=copy.deepcopy(modern);edit(bad)
        try: validate_focus(prepare(bad,ev),strict)
        except ValueError: checks+=1
        else: raise AssertionError('invalid V3 plan accepted')
    for value in [None,{},dict(target=[],checks=['核对']),dict(target='客户端',checks=['核对']),dict(target='服务端',checks=[]),dict(target='服务端',checks=[' ']),dict(target='服务端',checks='字符串')]:
        reject(lambda p,v=value:p['cases'][0].update(review_focus=v))
    reject(lambda p:p['cases'][0].pop('review_focus'))
    reject(lambda p:p['cases'][0].pop('acceptance_path'))
    for value in ['', '  ', None, [], 7]:reject(lambda p,v=value:p['cases'][0].update(acceptance_path=v))
    for value in ['',None,[]]:reject(lambda p,v=value:p['questions'][0].update(impact=v))
    reject(lambda p:p['questions'][0].pop('impact'))
    assert not any('review_focus' in c for c in prepare(base,ev)['cases'])
    for field in ['checks','target']:
        revised=copy.deepcopy(modern)
        f=next(c for c in revised['cases'] if c['id']=='TC-001')['review_focus']
        f[field]=['变更后的合成检查项'] if field=='checks' else '客户端简查'
        changes=prepare(revised,ev,baseline=p['_snapshot'])['_changes']
        assert [c['id'] for c in changes if c['change']=='需重新验证']==['TC-001']
    revised=copy.deepcopy(modern);revised['questions'][0]['impact']='影响跨天后的积分保留。'
    assert [c['id'] for c in prepare(revised,ev,baseline=p['_snapshot'])['_changes'] if c['change']=='需重新验证']==['TC-009']
    revised=copy.deepcopy(modern);revised['cases'].reverse();revised['cases'][0]['xmind_points']=['改写展示短句']
    assert prepare(revised,ev)['_snapshot']['cases']==p['_snapshot']['cases']
    revised=copy.deepcopy(modern);revised['cases'][0]['acceptance_path']='查看活动说明 → 核对文字（改写指引）'
    assert prepare(revised,ev)['_snapshot']['cases']==p['_snapshot']['cases']
    out=work/'output'
    run('build_documents.py','--plan',work/'v3-plan.json','--evidence',work/'evidence.json','--out',out,'--require-v3')
    run('build_documents.py','--plan',work/'v2-plan.json','--evidence',work/'evidence.json','--out',work/'strict-rejected','--require-v3',ok=False)
    run('build_documents.py','--plan',work/'v3-plan.json','--evidence',work/'evidence.json','--out',work/'flags-rejected','--require-v3','--allow-legacy',ok=False)
    assert not (work/'strict-rejected').exists()
    assert not (work/'flags-rejected').exists()
    run('build_documents.py','--plan',work/'v2-plan.json','--evidence',work/'evidence.json','--out',work/'old-plan-output')
    w=openpyxl.load_workbook(out/'合成积分领奖-AI验收表.xlsx',data_only=True)
    assert w['AI实现检查']['H5'].value=='最短验收路径'
    assert len(w.sheetnames)==7 and w['AI实现检查'].max_column==10
    assert w['待确认规则']['E6'].value==modern['questions'][0]['impact']+'\nTC-009'
    with zipfile.ZipFile(out/'合成积分领奖-测试用例.xmind') as z:sheets=json.loads(z.read('content.json'))
    nodes={t['labels'][0]:t for t in walk(sheets[0]['rootTopic']) if t.get('labels')}
    for r,c in enumerate(p['cases'],6):
        assert w['AI实现检查'].cell(r,1).value==c['id']
        assert w['AI实现检查'].cell(r,8).value==c['acceptance_path']
        assert w['AI实现检查'].cell(r,7).value is None
        assert w['检查详情'].cell(r,3).value==detail_action(c)
        for check in c['review_focus']['checks']:
            assert check in nodes[c['id']]['notes']['plain']['content']
            assert check in w['执行准备'].cell(r,8).value
        assert w['检查详情'].cell(r,12).value.startswith(' > '.join([p['title']]+output_path(c)))
    module=kids(sheets[0]['rootTopic'])[0]
    assert kids(module)[-1]['title']=='客户端简查'
    assert [c['labels'][0] for c in kids(kids(module)[-1])]==['TC-010','TC-011']
    assert modern['questions'][0]['impact'] in kids(kids(sheets[1]['rootTopic'])[0])[2]['title']
    return dict(cases=11,server_cases=9,client_cases=2,rejected_invalid_extensions=checks,
        checked=['真实思路快照引用','严格新设计/兼容旧计划','检查目标与清单导出','通俗问题与影响','检查项和影响变更触发复核','展示排序和短句不改指纹','客户端末尾分支','两文件ID与完整核对项一致'],
        limitation='人工编写合成对照，未执行真实游戏或代码验收，不证明真实需求质量提升')
