"""Synthetic status and saved-formula checks; not native Excel recalculation."""
from itertools import product
import re
import openpyxl
from openpyxl.formula.translate import Translator
from acceptance_model import STATUSES, LEGACY_STATUSES, LAYER_RESULTS, SUMMARY_HEADERS, formula
from build_documents import eval_formula
from workbook_export import HEADERS, DETAIL_HEADERS


def exercise_status():
    base = dict(I='静态代码检查', K='入口校验积分后标记领取并发奖', L='synthetic/reward.py:12-24（合成）',
                M='已查入口；回调链路待补', N='合成奖励模块', T='自测', U='2026-11-11', V='fixture-v1',
                X='处理链路未查清', Y='未验证', Z='符合', AA='未验证')
    checks = 0
    def check(expected, **fields):
        nonlocal checks
        actual = eval_formula(formula(6), dict(base, **fields))
        assert actual == expected, (fields, expected, actual)
        checks += 1
    check('未填写状态', D='')
    check('状态值无效', D='随便通过')
    for state in LEGACY_STATUSES: check('旧状态待逐条复核', D=state)
    check('尚未开始', D='【未开始】', Z='未验证')
    check('状态冲突：已有检查结果', D='【未开始】')
    check('反馈已填写', D='【已通过】')
    check('反馈已填写', D='【已通过】', Y='符合', Z='未验证')
    check('反馈已填写', D='【已通过】', Y='符合')
    check('反馈已填写', D='【已通过】', M='', X='')
    check('反馈已填写', D='【已通过】', H='重复回调是否重复到账')
    check('待补关键处理的完整依据', D='【已通过】', Z='未验证')
    check('待补关键处理的完整依据', D='【已通过】', Z='部分验证')
    check('状态冲突：规则仍待确认', D='【已通过】', Y='待确认', R='批准了另一问题')
    check('待补规则批准记录', D='【已通过】', O='Q01')
    check('反馈已填写', D='【已通过】', O='Q01', R='合成批准口径/人/时间/版本')
    check('反馈已填写', D='【部分验收】', Z='部分验证')
    check('待补已检查内容', D='【部分验收】', Z='未验证')
    check('反馈已填写', D='【验收受阻】', Z='未验证', X='源码或配置缺失')
    check('待补具体缺口/原因/下一步', D='【验收受阻】', X='')
    check('反馈已填写', D='【待确认】', Z='待确认', O='Q01')
    check('反馈已填写', D='【未通过】', Z='不符', O='Q01')
    check('待补不符层及差异', D='【未通过】')
    check('反馈已填写', D='【不适用】', Y='不适用', Z='不适用', AA='不适用', X='范围不适用')
    check('状态冲突：整体不适用须说明各层', D='【不适用】')
    check('分层结果值无效', D='【已通过】', Y='已看过')
    check('待补三层结果（未采用填未验证）', D='【已通过】', Y='')
    for col in ['I','N','T','U','V']: check('待补检查方式/范围/基线', D='【已通过】', **{col:''})
    for col in ['K','L']: check('待补实现说明/定位证据', D='【已通过】', **{col:''})
    combinations = 0
    for values in product(LAYER_RESULTS, repeat=3):
        for state in STATUSES:
            answer = eval_formula(formula(6), dict(base, D=state, **dict(zip(['Y','Z','AA'],values))))
            if '不符' in values and state!='【未通过】': assert answer=='状态冲突：已记录不符'
            if state=='【已通过】' and answer=='反馈已填写':
                assert '符合' in values and set(values)<={'符合','不适用','未验证'}
            combinations += 1
    return {'representative_scenarios':checks,'state_layer_combinations':combinations}


def workbook_value(w, sheet, address):
    """Small independent evaluator for the actual emitted lookup/navigation formulas.

    Deliberately no native-compatibility claim. Reject unsupported tokens/functions.
    """
    value=w[sheet][address].value
    if not isinstance(value,str) or not value.startswith('='): return '' if value is None else value
    pattern=r'''"(?:[^"]|"")*"|(?:'[^']+'!)?\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|[A-Z]+|\d+|<>|[=(),&+\-]'''
    source=value[1:];tokens=re.findall(pattern,source)
    assert ''.join(tokens)==source, source
    pos=0
    def atom():
        nonlocal pos
        token=tokens[pos];pos+=1
        if token=='(':
            value=expr();assert tokens[pos]==')';pos+=1;return value
        if token.startswith('"'):return token[1:-1].replace('""','"')
        if token.isdigit():return int(token)
        if pos<len(tokens) and tokens[pos]=='(':
            pos+=1;args=[]
            while tokens[pos]!=')':
                args.append(expr())
                if tokens[pos]==',':pos+=1
            pos+=1
            if token=='IF':return args[1] if args[0] else args[2]
            if token=='OR':return any(args)
            if token=='AND':return all(args)
            if token=='MATCH':assert args[2]==0;return args[1].index(args[0])+1
            if token=='INDEX':return args[0][args[1]-1]
            if token=='HYPERLINK':return args[1]
            raise ValueError(token)
        target=sheet;ref=token
        if '!' in ref: target,ref=ref.split('!');target=target.strip("'")
        ref=ref.replace('$','')
        if ':' in ref:
            return [workbook_value(w,target,c.coordinate) for row in w[target][ref] for c in row]
        return workbook_value(w,target,ref)
    def expr():
        nonlocal pos
        left=atom()
        while pos<len(tokens) and tokens[pos] in ['&','+','-']:
            op=tokens[pos];pos+=1;right=atom()
            left=str(left)+str(right) if op=='&' else left+right if op=='+' else left-right
        if pos<len(tokens) and tokens[pos] in ['=','<>']:
            op=tokens[pos];pos+=1;right=expr()
            return left==right if op=='=' else left!=right
        return left
    result=expr();assert pos==len(tokens);return result


def check_saved_workbook(path, count):
    w=openpyxl.load_workbook(path,data_only=False);cache=openpyxl.load_workbook(path,data_only=True)
    main=w['AI实现检查'];detail=w['检查详情'];summary=w['验收说明与汇总']
    assert w.active.title==main.title and main.max_column==10 and detail.max_column==29
    assert [c.value for c in main[5]]==HEADERS and [c.value for c in detail[5]]==DETAIL_HEADERS
    assert not any(c in DETAIL_HEADERS for c in ['AI检查状态','实现说明','后续测试重点'])
    for ws,cols,choices in [(main,['F'],STATUSES),(detail,['Q','R','S'],LAYER_RESULTS)]:
        for col in cols:
            assert any(d.formula1=='"'+','.join(choices)+'"' and col+'6' in d.sqref for d in ws.data_validations.dataValidation)
    assert all(detail.cell(r,c).value is None for r in range(6,count+6) for c in range(14,27))
    assert all(main.cell(r,c).value is None for r in range(6,count+6) for c in [7])
    assert all(summary[f'B{r}'].value is None for r in range(17,22))
    for r in range(6,count+6):
        cid=detail.cell(r,1).value
        assert workbook_value(w,main.title,f'A{r}')==cid
        assert workbook_value(w,detail.title,f'AA{r}')==cache[detail.title].cell(r,27).value
        assert len(detail.cell(r,27).value)<8192
        for sh,col,target in [(main,10,detail.title),(detail,28,main.title)]:
            nav=sh.cell(r,col).value
            assert f'MATCH("{cid}",\'{target}\'!$A$6:$A${count+5},0)+5' in nav
    assert [summary.cell(24,c).value for c in range(1,16)]==SUMMARY_HEADERS
    totalrow=next(r for r in range(25,summary.max_row+1) if summary.cell(r,1).value=='合计')
    assert cache[summary.title].cell(totalrow,2).value==count
    assert sum(cache[summary.title].cell(totalrow,c).value for c in range(3,14))+cache[summary.title].cell(totalrow,15).value==count
    records=['']+STATUSES+LEGACY_STATUSES+['无效值'];computed=[]
    for col in range(3,14):
        criterion=re.findall(r'"([^"]*)"',summary.cell(25,col).value)[-1]
        computed.append(sum(v==('' if criterion=='=' else criterion) for v in records))
    assert computed==[1]*11 and len(records)-sum(computed)==1
    assert summary.cell(25,15).value=='=B25-SUM(C25:M25)'
    # Fill only a synthetic in-memory workbook: inherit baseline and verify real formulas.
    for r,value in [(17,'fixture-v1'),(18,'合成奖励模块'),(19,'静态代码检查'),(20,'自测'),(21,'2026-11-11')]:summary.cell(r,2,value)
    main['F6']='【已通过】';main['G6']='积分校验通过后发奖并标记';main['H6']='领取奖励 → 重登 → 检查资产不再增加'
    detail['N6']='synthetic/reward.py:12-24（合成）';detail['Q6']='未验证';detail['R6']='符合';detail['S6']='未验证'
    assert workbook_value(w,detail.title,'AA6')=='反馈已填写'
    summary['B17']=None
    assert workbook_value(w,detail.title,'AA6')=='待补检查方式/范围/基线'
    detail['Z6']='fixture-override'
    assert workbook_value(w,detail.title,'AA6')=='反馈已填写'
    detail['R6']='部分验证'
    assert workbook_value(w,detail.title,'AA6')=='待补关键处理的完整依据'
    detail['R6']='符合';main['I6']='Q-new'
    assert workbook_value(w,detail.title,'AA6')=='待补规则批准记录'
    detail['T6']='合成批准记录'
    assert workbook_value(w,detail.title,'AA6')=='反馈已填写'
    main['F6']='【待实测】'
    assert workbook_value(w,detail.title,'AA6')=='旧状态待逐条复核'
    main['F6']='【已通过】'
    # Swap main rows alone, then details alone, translating row-local formula refs.
    if count>1:
        def swap(ws):
            rows=[[c.value for c in ws[r]] for r in [6,7]]
            for dst,src,values in [(6,7,rows[1]),(7,6,rows[0])]:
                for col,val in enumerate(values,1):
                    if isinstance(val,str) and val.startswith('='):
                        val=Translator(val,origin=ws.cell(src,col).coordinate).translate_formula(ws.cell(dst,col).coordinate)
                    ws.cell(dst,col).value=val
        swap(main)
        assert workbook_value(w,detail.title,'AA6')=='反馈已填写'
        swap(detail)
        assert workbook_value(w,detail.title,'AA7')=='反馈已填写'
    return {'columns':10,'detail_columns':29,'summary_categories_checked':12,'layer_dropdowns':3,
            'initial_inputs':'blank results; pending status/questions and designed path may be populated','actual_saved_formulas':'shared/local baseline, status, Q and independent sort checked'}
