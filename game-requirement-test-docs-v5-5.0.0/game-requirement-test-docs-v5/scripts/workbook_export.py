"""Compact implementation-review workbook, with a single writable conclusion per case."""
import math
import json
import xlsxwriter
import openpyxl
from acceptance_model import STATUSES, LEGACY_STATUSES, LAYER_RESULTS, REASONS, SUMMARY_HEADERS, SUMMARY_WIDTHS, INSTRUCTIONS, formula
from design_model import tables
from review_design import detail_action, output_path
from presentation_model import display_module

HEADERS=['编号','模块','优先级','测试点','需求简述','AI检查状态','实现说明','最短验收路径','问题编号','详情']
WIDTHS=[90,120,65,220,300,115,330,380,100,75]
DETAIL_HEADERS=['编号','测试点','条件与操作','预期与观察','影响核对','规则性质','规则ID','覆盖项ID','需求证据','思路来源','完整备注','XMind路径','主题ID','实现检查证据','检查范围（例外）','检查方式（例外）','配置核对结果','代码核对结果','运行验证结果','批准记录','缺口/不适用原因','补充记录','缺陷/任务编号','检查者（例外）','检查时间（例外）','版本/提交（例外）','反馈提示','返回主表','模块']
DETAIL_WIDTHS=[90,220,440,440,300,110,130,150,550,550,800,400,220,600,350,180,130,130,130,350,180,400,180,150,150,350,280,100,120]
# On the overview these rows have dedicated shared input cells B:J.
BASELINE_ROWS={'V':17,'N':18,'I':19,'T':20,'U':21}
INK='#24364B';NAVY='#274867';AMBER='#FFF5DB'

def brief(c):return '\n'.join(c.get('xmind_points',[])) or c['title']
def pending(c):return bool(c['questions'] or c['kind']=='待确认')

def lookup(sheet, col, cid, last):
    return f'INDEX(\'{sheet}\'!${col}$6:${col}${last},MATCH("{cid}",\'{sheet}\'!$A$6:$A${last},0))'

def detail_formula(row,cid,last):
    # Identity-based references remain correct if either table is sorted independently.
    refs={k:lookup('AI实现检查',col,cid,last) for k,col in [('D','F'),('K','G'),('O','I')]}
    mapping={'L':'N','N':'O','I':'P','Y':'Q','Z':'R','AA':'S','R':'T','X':'U','M':'V','T':'X','U':'Y','V':'Z'}
    for key,col in mapping.items():
        local=f'{col}{row}'
        if key in BASELINE_ROWS:
            shared=f"'验收说明与汇总'!$B${BASELINE_ROWS[key]}"
            refs[key]=f'IF({local}="",{shared},{local})'
        else:refs[key]=local
    return formula(row,refs)

def jump_formula(sheet,cid,last,label):
    return f'=HYPERLINK("#\'{sheet}\'!A"&(MATCH("{cid}",\'{sheet}\'!$A$6:$A${last},0)+5),"{label}")'

def build_excel(plan,evidence,ideas,qs,nodes,path,notes,case_title,eval_formula):
    wb=xlsxwriter.Workbook(path,{'strings_to_urls':False,'strings_to_formulas':False});wb.set_calc_mode('auto')
    wb.set_properties({'title':plan['title']+'-AI验收表','author':'game-requirement-test-docs-v5','comments':'5.0.0 AI实现检查；生成时未执行实现核对'})
    names=['AI实现检查','验收说明与汇总','检查详情','待确认规则']
    sh={name:wb.add_worksheet(name) for name in names}
    def fmt(**kw):return wb.add_format({'font_name':'Microsoft YaHei','font_size':10,'font_color':INK,'valign':'top',**kw})
    head=fmt(bg_color=NAVY,font_color='white',bold=True,text_wrap=True,valign='vcenter')
    link=fmt(font_color='#1666A8',underline=True,text_wrap=True)
    def pixels(s):return sum(14 if '\u2e80'<=c<='\uffff' else 8 for c in str(s or ''))
    def base(ws,title,headers,widths,rows,table,inputs=()):
        ws.hide_gridlines(2);ws.set_zoom(85)
        for i,w in enumerate(widths):ws.set_column_pixels(i,i,w)
        ws.set_row_pixels(1,36);ws.write(1,0,title,fmt(font_size=16,bold=True))
        ws.set_row_pixels(4,46)
        ws.add_table(4,0,4+len(rows),len(headers)-1,{'name':table,'style':None,'columns':[{'header':h,'header_format':head} for h in headers]})
        for r,row in enumerate(rows,5):
            heights=[sum(max(1,math.ceil(pixels(line)/(widths[c]-20))) for line in str(v or '').split('\n'))*21+16 for c,v in enumerate(row)]
            ws.set_row_pixels(r,min(540,max(64,*heights)))
            for c,v in enumerate(row):ws.write(r,c,v or '',fmt(text_wrap=True,bg_color=AMBER if c in inputs else ('#F3F6FA' if r%2==0 else 'white')))
        ws.freeze_panes(5,1);ws.repeat_rows(4);ws.set_landscape();ws.fit_to_pages(1,0)
    cases=plan['cases'];last=len(cases)+5
    rows=[[c['id'],display_module(c),c['priority'],c['title'],brief(c),'【待确认】' if pending(c) else '','',c.get('acceptance_path',''),'、'.join(c['questions']),'查看'] for c in cases]
    main=sh['AI实现检查'];base(main,plan['title']+' · AI实现检查',HEADERS,WIDTHS,rows,'ImplementationReview',{5,6,7,8})
    main.write(2,0,'实现符合可由静态证据支持；详细设计与检查记录点“详情”。',fmt(font_color='#566C80'))
    main.data_validation(f'F6:F{last}',{'validate':'list','source':STATUSES,'error_type':'stop'})
    for status,color in zip(STATUSES,['#F0F3F7','#FFF5DB','#FFE8CC','#EDF0FC','#FCE8E6','#EAF3EC','#EAEAEA']):
        main.conditional_format(f'F6:F{last}',{'type':'text','criteria':'containing','value':status,'format':fmt(bg_color=color)})
    dr=[]
    def ref(r):return f'{r["file"]} · {r["locator"]}\n{r["text"]}'
    for c in cases:
        dr.append([c['id'],c['title'],detail_action(c),c['expected'],c['impact'],c['kind'],'、'.join(c.get('rule_ids',[])),'、'.join(c.get('coverage_ids',[])),
            '\n\n'.join(ref(evidence[e]) for e in c['evidence']),'\n\n'.join(ref(ideas[i]) for i in c['ideas']),notes(c,evidence,ideas),
            ' > '.join([plan['title']]+output_path(c)+[case_title(plan,c)]),nodes[c['id']]['id']]+['']*14+['返回',display_module(c)])
    detail=sh['检查详情'];base(detail,plan['title']+' · 检查详情',DETAIL_HEADERS,DETAIL_WIDTHS,dr,'ReviewDetails',set(range(13,26)))
    # Hidden full notes must not stretch every visible evidence row to half a screen.
    for r in range(5,last):detail.set_row_pixels(r,110)
    detail.set_column(2,12,None,None,{'level':1,'hidden':True});detail.set_column_pixels(13,13,600,None,{'collapsed':True})
    # Restore widths when grouping design columns; evidence and results remain immediately accessible.
    for col in range(2,13):detail.set_column_pixels(col,col,DETAIL_WIDTHS[col],None,{'level':1,'hidden':True})
    detail.data_validation(f'P6:P{last}',{'validate':'list','source':['配置核对','静态代码检查','代码与配置核对','运行实测','自动化测试','日志分析','多方式交叉验证']})
    detail.data_validation(f'U6:U{last}',{'validate':'list','source':REASONS,'error_type':'stop'})
    for col in ['Q','R','S']:detail.data_validation(f'{col}6:{col}{last}',{'validate':'list','source':LAYER_RESULTS,'error_type':'stop'})
    for r,c in enumerate(cases,6):
        for col,label in [(0,c['id']),(9,'查看')]:main.write_formula(r-1,col,jump_formula('检查详情',c['id'],last,label),link,label)
        # Detail IDs remain literal to provide an independent lookup key.
        detail.write_formula(r-1,27,jump_formula('AI实现检查',c['id'],last,'返回'),link,'返回')
        detail.write_formula(r-1,26,detail_formula(r,c['id'],last),fmt(text_wrap=True,bg_color='#F0F3F7'),eval_formula(formula(r),{'D':'【待确认】' if pending(c) else ''}))
    detail.conditional_format(f'AA6:AA{last}',{'type':'text','criteria':'not containing','value':'反馈已填写','format':fmt(font_color='#8C5A19')})
    qrows=[[q['id'],q['title'],q['gap'],q['ask'],(q['impact']+'\n' if q.get('impact') else '')+'、'.join(c['id'] for c in cases if q['id'] in c['questions']),'','','','','',''] for q in qs.values()]
    if qrows:base(sh['待确认规则'],plan['title']+' · 待确认规则',['问题ID','主题','原文与缺口','需要确认的问题','影响说明与用例','当前实现','定位证据','差异或影响','批准口径','确认人/日期/版本','备注'],[95,220,380,400,350,400,450,350,350,250,350],qrows,'PendingRules',range(5,11))
    else:sh['待确认规则'].write(1,0,'本次未登记待确认问题。',fmt())
    ov=sh['验收说明与汇总'];ov.hide_gridlines(2)
    for i,w in enumerate(SUMMARY_WIDTHS):ov.set_column_pixels(i,i,w)
    ov.merge_range(1,0,1,14,plan['title']+' · AI实现检查结果',fmt(font_size=16,bold=True));ov.set_row_pixels(1,38)
    for r,line in enumerate(INSTRUCTIONS,3):
        ov.merge_range(r,0,r,9,line,fmt(text_wrap=True));ov.set_row_pixels(r,max(42,math.ceil(pixels(line)/(sum(SUMMARY_WIDTHS[:10])-25))*22+12))
    # Status names already appear in the summary; avoid repeating the judgment rules.
    for r in range(8,14):ov.set_row_pixels(r,8)
    ov.merge_range(14,0,14,9,'共用检查基线（逐条例外写检查详情；未提供不猜测）',fmt(bold=True));ov.set_row_pixels(14,32)
    for key,label in [('V','版本/提交'),('N','检查范围'),('I','检查方式'),('T','检查者'),('U','检查时间')]:
        rr=BASELINE_ROWS[key]-1;ov.write(rr,0,label,fmt(bold=True));ov.merge_range(rr,1,rr,9,'',fmt(bg_color=AMBER,text_wrap=True));ov.set_row_pixels(rr,42)
    ov.write_row(23,0,SUMMARY_HEADERS,head);ov.set_row_pixels(23,46)
    modules=list(dict.fromkeys(display_module(c) for c in cases));totals=[0]*14
    for row,m in enumerate(modules,25):
        group=[c for c in cases if display_module(c)==m];p=sum(pending(c) for c in group)
        values=[len(group),len(group)-p]+[p if s=='【待确认】' else 0 for s in STATUSES+LEGACY_STATUSES]+[0,0]
        totals=[a+b for a,b in zip(totals,values)]
        ov.write(row-1,0,m,fmt());ov.set_row_pixels(row-1,32)
        mr=f"'AI实现检查'!$B$6:$B${last}";sr=f"'AI实现检查'!$F$6:$F${last}"
        fs=[f'=COUNTIFS({mr},A{row})',f'=COUNTIFS({mr},A{row},{sr},"=")']
        fs += [f'=COUNTIFS({mr},A{row},{sr},"{s}")' for s in STATUSES+LEGACY_STATUSES]
        fs += [f'=COUNTIFS(\'检查详情\'!$AC$6:$AC${last},A{row},\'检查详情\'!$AA$6:$AA${last},"反馈已填写")',f'=B{row}-SUM(C{row}:M{row})']
        for col,(f,v) in enumerate(zip(fs,values),1):ov.write_formula(row-1,col,f,fmt(),v)
    totalrow=25+len(modules);ov.write(totalrow-1,0,'合计',fmt(bold=True));ov.set_row_pixels(totalrow-1,32)
    for col,value in enumerate(totals,1):
        letter=xlsxwriter.utility.xl_col_to_name(col);ov.write_formula(totalrow-1,col,f'=SUM({letter}25:{letter}{totalrow-1})',fmt(bold=True),value)
    for number,(name,headers,widths,rows) in enumerate(tables(plan),1):
        names.append(name);base(wb.add_worksheet(name),plan['title']+' · '+name,headers,widths,rows,'DesignModel'+str(number))
    main.activate();main.set_first_sheet();wb.close()
    return names,totalrow,totals,24


def validate_excel(path,plan,nodes,notes,evidence,ideas,case_title):
    w=openpyxl.load_workbook(path,data_only=False);cache=openpyxl.load_workbook(path,data_only=True)
    main=w['AI实现检查'];detail=w['检查详情'];last=len(plan['cases'])+5
    assert main.max_column==10 and detail.max_column==29
    assert w.active.title=='AI实现检查' and main.freeze_panes=='B6'
    for row,c in enumerate(plan['cases'],6):
        assert cache['AI实现检查'].cell(row,1).value==c['id']
        for col,value in [(2,display_module(c)),(3,c['priority']),(4,c['title']),(5,brief(c)),(6,'【待确认】' if pending(c) else None),(9,'、'.join(c['questions']) or None)]:assert main.cell(row,col).value==value
        assert main.cell(row,7).value is None
        assert main.cell(row,8).value==(c.get('acceptance_path') or None)
        for col,value in [(1,c['id']),(3,detail_action(c)),(4,c['expected']),(5,c['impact']),(6,c['kind']),(11,notes(c,evidence,ideas)),(12,' > '.join([plan['title']]+output_path(c)+[case_title(plan,c)])),(13,nodes[c['id']]['id']),(29,display_module(c))]:assert detail.cell(row,col).value==value
        assert all(detail.cell(row,col).value is None for col in range(14,27))
        assert detail.cell(row,27).value==detail_formula(row,c['id'],last)
        assert main.cell(row,10).value==jump_formula('检查详情',c['id'],last,'查看')
        assert detail.cell(row,28).value==jump_formula('AI实现检查',c['id'],last,'返回')
    assert any(v.formula1=='"'+','.join(STATUSES)+'"' and 'F6' in v.sqref for v in main.data_validations.dataValidation)
    for col in ['Q','R','S']:assert any(v.formula1=='"'+','.join(LAYER_RESULTS)+'"' and col+'6' in v.sqref for v in detail.data_validations.dataValidation)
    return w,cache
