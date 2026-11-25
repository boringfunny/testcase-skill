"""V5 format regressions with explicitly synthetic, non-game fixtures."""
import copy
import json
import zipfile
from pathlib import Path
import openpyxl
from self_test_v3 import fixture, run
from design_model import prepare
from presentation_model import validate_display, ordered_display_cases, template_report, GROUPS
from build_documents import walk, kids


def classify(modern):
    p=copy.deepcopy(modern);p['template_profile']='活动'
    mapping={1:('配置检查',['配置检查']), 8:('异常状态',['异常状态','重复请求']),
             9:('特殊点验证',['特殊点验证','跨天']),10:('UI检查',['UI检查','积分领奖']),11:('UI检查',['UI检查','积分领奖'])}
    for c in p['cases']:
        n=int(c['id'].split('-')[1])
        group,path=mapping.get(n,('主功能',['积分领奖']) if n<5 else ('入口',['入口','领取资格']))
        c.update(display_group=group,display_path=path)
    p['template_decisions']={g: '合成需求存在对应独立场景' for g in GROUPS}
    p['template_decisions']['红点']='合成需求未规定红点规则，不扩展红点预期。'
    return p


def save(path,value):path.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')


def small_fixture(work, profile, title, rule, client=False):
    # Separate source text supports this fixture; no inherited reward/hero mechanics.
    work.mkdir(parents=True)
    source=work/'合成需求.txt';source.write_text('合成结构样例，不代表真实游戏。\n'+rule,encoding='utf-8')
    evidence=work/'evidence.json'
    run('read_sources.py','--workspace',work,'--requirement',source,'--out',evidence)
    ev=json.loads(evidence.read_text(encoding='utf-8'));refs=[r['id'] for r in ev['records'] if r['category']=='requirement']
    group='UI检查' if client else '主功能';path=['UI检查',title] if client else [title]
    p=dict(schema_version=2,title='合成'+title,requirement_version='synthetic-v1',source_review='已读合成两行文本，无图片或修订',
      template_profile=profile,template_decisions={g:'只测试本样例明示规则，其他方向未提供需求' for g in GROUPS},
      rules=[dict(id='RULE-001',statement=rule,kind='需求验证',evidence=refs,rationale='合成需求明示',questions=[])],
      coverage_items=[dict(id='COV-001',rule_id='RULE-001',method='直接验证',target=rule,rationale='核对明示规则',disposition='planned')],
      questions=[],scope_notes=['合成结构回归；不是项目事实或验收结果。'],cases=[dict(id='TC-001',path=[title],display_group=group,display_path=path,
        title='核对'+title,xmind_points=[rule],coverage_ids=['COV-001'],preconditions=['取得本合成规则对应版本与检查材料'],test_data=[],
        steps=[dict(action='查看对应画面' if client else '追踪条件判断和对应状态处理',expected=rule,observe='实际画面证据' if client else '条件分支及处理结果证据')],
        postconditions=['本次未执行；后续检查结束恢复测试数据'],impact=title+'符合需求',priority='P1',priority_reason='直接业务要求',
        verification=dict(method='运行实测' if client else '静态代码检查',requirements=['需取得相应版本及定位证据'],limitations=['本次未执行']),
        kind='需求验证',evidence=refs,ideas=['COM-024'] if client else ['COM-021'],questions=[],
        review_focus=dict(target='客户端简查' if client else '服务端',checks=[rule]),
        acceptance_path='从已确认的相关入口进入（具体路径待补） → 触发对应场景 → '+rule)])
    save(work/'plan.json',p)
    run('build_documents.py','--plan',work/'plan.json','--evidence',evidence,'--out',work/'output','--require-v5')
    return p


def exercise_v5(work):
    _,modern,ev=fixture(work)
    p=classify(modern);save(work/'v5-plan.json',p)
    rendered=prepare(p,ev);validate_display(rendered,True)
    original=prepare(modern,ev)
    assert rendered['_snapshot']['cases']==original['_snapshot']['cases']
    assert rendered['_snapshot']['rules']==original['_snapshot']['rules']
    assert [c['id'] for c in rendered['cases']]==[f'TC-{i:03}' for i in [5,6,7,2,3,4,9,8,1,10,11]]
    source_by_id={c['id']:c for c in modern['cases']}
    for c in p['cases']:
        assert {k:v for k,v in c.items() if k not in ('display_group','display_path')}==source_by_id[c['id']]
    revised=copy.deepcopy(p);revised['template_profile']='通用'
    c=next(c for c in revised['cases'] if c['id']=='TC-001');c.update(display_group='主功能',display_path=['奖励配置'])
    assert prepare(revised,ev)['_snapshot']==rendered['_snapshot']
    c['steps'][0]['expected']='合成变更：阈值200'
    changes=prepare(revised,ev,baseline=rendered['_snapshot'])['_changes']
    assert [c['id'] for c in changes if c['change']=='需重新验证']==['TC-001']
    invalid=0
    def reject(edit):
        nonlocal invalid
        bad=copy.deepcopy(p);edit(bad)
        try:validate_display(prepare(bad,ev),True)
        except ValueError:invalid+=1
        else:raise AssertionError('Invalid V5 configuration accepted')
    reject(lambda p:p.pop('template_profile'))
    for value in ['未知',None,[],{}]:reject(lambda p,v=value:p.update(template_profile=v))
    for field in ['display_group','display_path','xmind_points','acceptance_path','review_focus']:
        reject(lambda p,f=field:p['cases'][0].pop(f))
    for value in [[],[''],['主功能1'],['次级模块'],['功能点1'],[' UI检查']]:
        reject(lambda p,v=value:p['cases'][0].update(display_path=v))
    reject(lambda p:p['cases'][0].update(display_group='主功能',display_path=['动画']))
    reject(lambda p:p['cases'][0].update(display_group='入口',display_path=['其他']))
    reject(lambda p:p['cases'][0].update(review_focus=None))
    reject(lambda p:p['questions'][0].pop('impact'))
    assert template_report(original)['mode']=='legacy'
    assert original['_snapshot']['cases']==prepare(modern,ev)['_snapshot']['cases']
    # Interleaved functions and deeper groups use tree order, not flat category order.
    order=[dict(id=str(i),display_group=g,display_path=path) for i,g,path in [
      (1,'主功能',['领奖','甲']), (2,'主功能',['商店']), (3,'主功能',['领奖','乙']),
      (4,'主功能',['领奖','甲']), (5,'红点',['红点']), (6,'入口',['入口'])]]
    assert [c['id'] for c in ordered_display_cases(order)]==['6','5','1','4','3','2']
    out=work/'output'
    run('build_documents.py','--plan',work/'v5-plan.json','--evidence',work/'evidence.json','--out',out,'--require-v5')
    for flags in [[],['--allow-legacy']]:
        run('build_documents.py','--plan',work/'v3-plan.json','--evidence',work/'evidence.json','--out',work/'rejected','--require-v5',*flags,ok=False)
    assert not (work/'rejected').exists()
    run('build_documents.py','--plan',work/'v5-plan.json','--evidence',work/'evidence.json','--out',out,'--require-v5',ok=False)
    with zipfile.ZipFile(next(out.glob('*.xmind'))) as z:
        sheets=json.loads(z.read('content.json'));styles=z.read('styles.xml').decode()
    root=sheets[0]['rootTopic']
    assert [b['title'] for b in kids(root)]==['入口','积分领奖','特殊点验证','异常状态','配置检查','UI检查']
    assert root['extensions'][0]['content'][0]['content']=='3'
    assert sheets[0]['theme']['map']['properties']['svg:fill']=='#eee8e6'
    assert '#eee8e6' in styles and '#c1aba5' in styles and 'branchConnection.curve' in styles
    ids=[t['labels'][0] for t in walk(root) if t.get('labels')]
    assert ids==[c['id'] for c in rendered['cases']] and len(set(ids))==11
    assert not any(t['title'] in ['已验收','未通过','主功能1','功能点1'] for t in walk(root))
    w=openpyxl.load_workbook(next(out.glob('*.xlsx')),data_only=True)
    assert len(w.sheetnames)==7 and w['AI实现检查'].max_column==10
    for row,c in enumerate(rendered['cases'],6):
        assert w['AI实现检查'].cell(row,1).value==c['id']
        assert w['AI实现检查'].cell(row,2).value==c['display_path'][0]
        assert w['AI实现检查'].cell(row,8).value==c['acceptance_path']
        assert w['检查详情'].cell(row,29).value==c['display_path'][0]
        assert ' > '.join(c['display_path']) in w['检查详情'].cell(row,12).value
    assert [w['验收说明与汇总'].cell(r,1).value for r in range(25,31)]==[b['title'] for b in kids(root)]
    assert w['验收说明与汇总']['B31'].value==11
    profiles=[]
    for profile,title,rule,client in [
      ('玩法','合成关卡资格','达到10级可进入关卡，低于10级拒绝进入。',False),
      ('养成','合成装备升级','材料足够时消耗2个材料并提升1级，材料不足时不扣除材料。',False),
      ('英雄','合成英雄说明','英雄说明文字显示“合成占位示例”。',True),
      ('通用','合成说明面板','说明面板显示“仅用于结构回归”。',True)]:
        small_fixture(work/profile,profile,title,rule,client);profiles.append(profile)
    return dict(result='passed',cases=11,invalid_rejected=invalid,additional_profiles=profiles,
       checked=['同业务ID/数量/最短路径/指纹不变','分类顺序及树顺序一致','主表/详情/汇总分类一致',
                '旧计划回退及严格入口拒绝','实际设计变更触发复核','双格式暖色样式','客户端仅UI检查','重复导出拒绝'],
       limitations=['合成样例不是业务验收','原生打开与重算另行记录'])
