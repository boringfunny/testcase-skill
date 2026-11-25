"""Exercise document export and installation using synthetic fixtures, without game verification."""
import argparse,copy,importlib.util,json,os,subprocess,sys,tempfile,zipfile
from pathlib import Path
import xlsxwriter

HERE=Path(__file__).resolve().parent

def run(script,*args,ok=True):
    env=dict(os.environ,PYTHONIOENCODING='utf-8')
    if str(script)=='build_documents.py' and '--plan' in args:
        pp=Path(args[args.index('--plan')+1])
        if 'schema_version' not in json.loads(pp.read_text(encoding='utf-8')):
            args=(*args,'--allow-legacy')
    proc=subprocess.run([sys.executable,str(HERE/script),*map(str,args)],capture_output=True,encoding='utf-8',env=env)
    assert (proc.returncode==0)==ok,(script,proc.stdout,proc.stderr)
    return proc

def exercise(work):
    work.mkdir(parents=True,exist_ok=True)
    req=work/'synthetic-requirement.xlsx';ideas=work/'synthetic-ideas.xlsx';style=work/'synthetic-style.xmind'
    for path,rows in [(req,[['合成测试：只用于工具自测'],['点击按钮显示说明。'],['奖励发放时点未明确。']]),(ideas,[['COM-001','按钮交互','核对点击与显示结果']])]:
        w=xlsxwriter.Workbook(path);s=w.add_worksheet('测试资料')
        for r,row in enumerate(rows):s.write_row(r,0,row)
        w.close()
    with zipfile.ZipFile(style,'w') as z:
        z.writestr('content.xml','''<xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0"><sheet><title>合成风格</title><topic><title>模块</title><notes><plain>样本备注需保留</plain></notes></topic></sheet></xmap-content>''')
    efile=work/'evidence.json'
    run('read_sources.py','--workspace',work,'--requirement',req,'--ideas',ideas,'--style',style,'--out',efile)
    ev=json.loads(efile.read_text(encoding='utf-8'))
    rr=[r for r in ev['records'] if r['category']=='requirement']
    assert len(rr)==3 and any(r.get('idea_id')=='COM-001' for r in ev['records'])
    assert any('样本备注需保留' in r['text'] for r in ev['records'] if r['category']=='style')
    docx=work/'revisions.docx'
    with zipfile.ZipFile(docx,'w') as z:
        z.writestr('word/document.xml','''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>正文</w:t></w:r><w:del><w:r><w:delText>旧口径</w:delText></w:r></w:del><w:ins><w:r><w:t>新口径</w:t></w:r></w:ins></w:p></w:body></w:document>''')
    run('read_sources.py','--workspace',work,'--requirement',docx,'--ideas',ideas,'--style',style,'--out',work/'docx-evidence.json')
    drecs=json.loads((work/'docx-evidence.json').read_text(encoding='utf-8'))['records']
    assert any('[del修订：旧口径]' in r['text'] and '[ins修订：新口径]' in r['text'] for r in drecs)
    c=dict(id='TC-001',path=['交互','说明按钮'],title='点击后显示说明',action='点击说明按钮',expected='出现需求说明',impact='核对显示内容',priority='P1',kind='需求验证',evidence=[rr[1]['id']],ideas=['COM-001'],questions=[],original_notes='旧图原始备注：保留多行\n第二行原文')
    cq=dict(c,id='TC-002',title='奖励发放时点口径',kind='待确认',evidence=[rr[2]['id']],questions=['Q01'],expected='记录当前可观察行为，待规则获批后判定')
    plan=dict(title='工具合成样例',cases=[c,cq],questions=[dict(id='Q01',title='发放时点',gap='文档未明确',ask='确认何时发放',evidence=[rr[2]['id']])],scope_notes=['仅工具自测，虚构资料不代表任何实际游戏规则。'])
    pp=work/'plan.json';pp.write_text(json.dumps(plan,ensure_ascii=False),encoding='utf-8')
    out=work/'two-cases';run('build_documents.py','--plan',pp,'--evidence',efile,'--out',out)
    report=json.loads((out/'.support/validation.json').read_text(encoding='utf-8'))
    assert report['cases']==2 and report['questions']==1 and report['initial_statuses']['pending']==1
    with zipfile.ZipFile(out/'工具合成样例-测试用例.xmind') as z:
        exported=json.loads(z.read('content.json'))
    saved_case=exported[0]['rootTopic']['children']['attached'][0]['children']['attached'][0]['children']['attached'][0]
    assert c['original_notes'] in saved_case['notes']['plain']['content']
    assert [t['title'] for t in saved_case['children']['attached']]==['条件/操作：'+c['action'],'预期：'+c['expected'],'影响核对：'+c['impact']]
    before={p:p.read_bytes() for p in out.iterdir() if p.is_file()}
    run('build_documents.py','--plan',pp,'--evidence',efile,'--out',out,ok=False)
    assert all(p.read_bytes()==b for p,b in before.items())
    plan_one=copy.deepcopy(plan);plan_one.update(title='单测试点',cases=[c],questions=[])
    one=work/'one-plan.json';one.write_text(json.dumps(plan_one,ensure_ascii=False),encoding='utf-8')
    run('build_documents.py','--plan',one,'--evidence',efile,'--out',work/'one-case')
    spec=importlib.util.spec_from_file_location('exporter',HERE/'build_documents.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
    checks=0
    def reject(edit):
        nonlocal checks
        bad=copy.deepcopy(plan);edit(bad)
        try:mod.check_plan(bad,ev)
        except AssertionError:checks+=1
        else:raise AssertionError('invalid plan was accepted')
    reject(lambda p:p['cases'][0].update(ideas=['COM-999999']))
    reject(lambda p:p['cases'][1].update(questions=['Q999']))
    reject(lambda p:p['cases'][1].update(kind='需求验证'))
    reject(lambda p:p['cases'][0].update(actual='假定通过'))
    reject(lambda p:p['cases'][0].update(evidence=['R999999']))
    saved=req.read_bytes();req.write_bytes(saved+b'changed')
    reject(lambda p:None);req.write_bytes(saved)
    from self_test_status import exercise_status, check_saved_workbook
    status_checks=exercise_status()
    saved_contract=check_saved_workbook(out/(plan["title"]+"-AI验收表.xlsx"),2)
    target=work/'installed-skills';run('install.py','--skills-dir',target)
    installed=target/HERE.parent.name
    assert (installed/'SKILL.md').read_bytes()==(HERE.parent/'SKILL.md').read_bytes()
    assert (installed/'assets/通用游戏测试思路.xlsx').read_bytes()==(HERE.parent/'assets/通用游戏测试思路.xlsx').read_bytes()
    run('install.py','--skills-dir',target,ok=False)
    # Empty workspace must fall back to bundled source snapshots, without fixed drive paths.
    empty=work/'empty-workspace';empty.mkdir()
    run('read_sources.py','--workspace',empty,'--requirement',req,'--out',work/'fallback.json')
    fallback=json.loads((work/'fallback.json').read_text(encoding='utf-8'))
    assert any(Path(s['path'])==HERE.parent/'assets/通用游戏测试思路.xlsx' for s in fallback['sources'])
    from self_test_v2 import exercise_v2
    v2=exercise_v2(work/'v2')
    from self_test_v3 import exercise_v3
    v3=exercise_v3(work/'v3')
    from self_test_v5 import exercise_v5
    v5=exercise_v5(work/'v5')
    return {'v5':v5,'result':'passed','synthetic_exports':[2,1],'rejected_invalid_plans':checks,'feedback_formula_scenarios':status_checks,'v2':v2,'v3':v3,'saved_status_contract':saved_contract,'tested':['XLSX evidence','classic XMind notes','DOCX revisions','matching XMind/XLSX exports','no-Q plan','no overwrite','source hash protection','state feedback prompts','installer copy/refuse overwrite','bundled fallback'],'game_acceptance':'not performed','native_application_open':'not performed'}

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--work-dir');a=p.parse_args()
    if a.work_dir:
        work=Path(a.work_dir).resolve()
        if work.exists():raise SystemExit('Use a new work directory; existing files are preserved.')
        report=exercise(work)
        (work/'self-test-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    else:
        with tempfile.TemporaryDirectory(prefix='game-test-docs-') as tmp:report=exercise(Path(tmp))
    print(json.dumps(report,ensure_ascii=False,indent=2))
