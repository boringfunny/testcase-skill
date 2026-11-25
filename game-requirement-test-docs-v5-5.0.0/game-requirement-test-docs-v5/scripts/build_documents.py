"""Export an evidence-backed AI-authored plan to matching XMind and AI acceptance XLSX."""
import argparse, collections, hashlib, json, math, re, uuid, zipfile
from pathlib import Path
import xml.etree.ElementTree as E
import xlsxwriter
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from design_model import prepare, tables
from review_design import validate_focus, focus_text, detail_action, output_path, review_report
from presentation_model import validate_display, enabled, template_report

from acceptance_model import STATUSES, LEGACY_STATUSES, LAYER_RESULTS, REASONS, SUMMARY_HEADERS, SUMMARY_WIDTHS, INSTRUCTIONS, formula
from workbook_export import HEADERS, WIDTHS, DETAIL_WIDTHS, build_excel as export_excel, validate_excel
INK='#24364B';BLUE='#548DD4';NAVY='#274867';AMBER='#FFF5DB'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def uid():return uuid.uuid4().hex
def node(title,kids=None,note=None,labels=None):
    t={'id':uid(),'class':'topic','title':title}
    if kids:t['children']={'attached':kids}
    if note:t['notes']={'plain':{'content':note}}
    if labels:t['labels']=labels
    return t
def kids(t):return t.setdefault('children',{}).setdefault('attached',[])
def walk(t):
    yield t
    for c in t.get('children',{}).get('attached',[]):yield from walk(c)
def pending(c):return bool(c['questions'] or c['kind']=='待确认')
def check_plan(plan,ev):
    assert isinstance(plan.get('title'),str) and plan['title'].strip(),'缺少title'
    assert plan.get('cases'),'没有测试点'
    evidence={r['id']:r for r in ev['records']};assert len(evidence)==len(ev['records'])
    ideas={r['idea_id']:r for r in ev['records'] if r.get('idea_id')}
    qs={q['id']:q for q in plan.get('questions',[])};assert len(qs)==len(plan.get('questions',[]))
    used=set()
    for q in qs.values():
        assert re.fullmatch(r'Q\d+',q['id']) and all(q.get(k) for k in ['title','gap','ask','evidence']),q
        assert all(e in evidence and evidence[e]['category']=='requirement' for e in q['evidence'])
    for c in plan['cases']:
        for key in ['id','path','title','action','expected','impact','priority','kind','evidence']:assert c.get(key),(c.get('id'),key)
        assert re.fullmatch(r'[A-Za-z][A-Za-z0-9]*-\d+',c['id']) and c['id'] not in used,c['id'];used.add(c['id'])
        assert isinstance(c['path'],list) and all(isinstance(s,str) and s.strip() for s in c['path'])
        assert c['priority'] in ['P0','P1','P2']
        assert c['kind'] in ['需求验证','边界派生','补充建议','待确认']
        c.setdefault('ideas',[]);c.setdefault('questions',[])
        assert all(q in qs for q in c['questions'])
        assert c['kind']!='待确认' or c['questions'],f'{c["id"]}待确认但没有Q依赖'
        assert not c['questions'] or c['kind']=='待确认',f'{c["id"]}有未确认规则须标待确认'
        assert all(e in evidence and evidence[e]['category']=='requirement' for e in c['evidence'])
        assert all(i in ideas for i in c['ideas'])
        assert c['ideas'] or plan.get('scope_notes'),'无适用思路时须说明'
        for k in ['status','actual','current_code','result','acceptance_status']:
            assert k not in c or not c[k],f'只制作文档，不接受预填实测字段：{k}'
    for s in ev['sources']:assert Path(s['path']).is_file() and sha(s['path'])==s['sha256'],f'源文件变化：{s["path"]}'
    return evidence,ideas,qs
def text_ref(r):return f'{r["file"]} · {r["locator"]}\n{r["text"]}'
def notes(c,evidence,ideas):
    return '\n'.join([f'编号：{c["id"]}','执行状态：未执行',f'性质：{c["kind"]}',f'建议优先级：{c["priority"]}',
        '需求证据：']+[text_ref(evidence[e]) for e in c['evidence']]+['通用思路：']+[text_ref(ideas[i]) for i in c['ideas']]+
        ['条件与操作（完整）：\n'+detail_action(c), '预期与观察（完整）：\n'+c['expected'], '影响核对：'+c['impact'],
         '导图简述：'+json.dumps(c.get('xmind_points',[]),ensure_ascii=False),
         '规则依赖：'+('、'.join(c['questions']) or '无'),
         '规则ID：'+'、'.join(c.get('rule_ids',[])), '覆盖项ID：'+'、'.join(c.get('coverage_ids',[])),
         '优先级理由：'+c.get('priority_reason','旧格式未提供'),
         '建议验证与准备：'+json.dumps(c.get('verification',{}),ensure_ascii=False),
         c.get('original_notes',''),'未执行代码检查或游戏验收。验收AI在Excel回填可观察行为、实际、证据和结论；逐例记录实际定位证据；静态证据不证明实际画面效果。'])

def compact_case(plan,c):
    return plan.get('schema_version')==2 and 'xmind_points' in c

def case_title(plan,c):
    return c['title']+('【待确认 '+ '、'.join(c['questions'])+'】' if compact_case(plan,c) and pending(c) else '')

def case_leaves(plan,c):
    if compact_case(plan,c):
        return c['xmind_points']
    label='判定（规则待确认）' if pending(c) else {'需求验证':'预期','边界派生':'预期（需求推导）','补充建议':'建议预期'}[c['kind']]
    return ['条件/操作：'+c['action'],label+'：'+c['expected'],'影响核对：'+c['impact']]+(['待确认：'+'、'.join(c['questions'])] if c['questions'] else [])

def presentation_report(plan):
    compact=[c['id'] for c in plan['cases'] if compact_case(plan,c)]
    missing=[c['id'] for c in plan['cases'] if plan.get('schema_version')==2 and not compact_case(plan,c)]
    return {'compact_cases':compact,'legacy_detail_cases':[c['id'] for c in plan['cases'] if not compact_case(plan,c)],
            'warnings':(['以下V2用例缺少xmind_points，已保留完整旧分支；需补写简述：'+ '、'.join(missing)] if missing else []),
            'long_summaries':[c['id'] for c in plan['cases'] if compact_case(plan,c) and (len(c['xmind_points'])>3 or any(len(s)>80 for s in c['xmind_points']))]}

def build_xmind(plan,evidence,ideas,qs,path):
    root=node(plan['title'],note='全部未执行。规则性质、来源与Q依赖见用例备注；仅测试设计，不代表实现或验收结果。')
    root['structureClass']='org.xmind.ui.map.unbalanced'
    roots={}
    case_nodes={}
    for c in plan['cases']:
        parent=root
        for seg in output_path(c):
            existing=next((t for t in kids(parent) if t['title']==seg and not t.get('labels')),None)
            if existing is None:existing=node(seg);kids(parent).append(existing)
            parent=existing
        leaves=[node(s) for s in case_leaves(plan,c)]
        labels=[c['id'],c['priority']]+([c['kind']] if compact_case(plan,c) else (['待确认'] if pending(c) else []))
        t=node(case_title(plan,c),leaves,notes(c,evidence,ideas),labels)
        if not compact_case(plan,c):t['branch']='folded'
        if pending(c):t['style']={'id':uid(),'properties':{'fo:color':'#9A5A00'}}
        kids(parent).append(t);case_nodes[c['id']]=t
    root['extensions']=[{'provider':'org.xmind.ui.map.unbalanced','content':[{'name':'right-number','content':str(max(1,math.ceil(len(kids(root))/2)))}]}]
    for b in kids(root):
        for sub in kids(b):
            if not sub.get('labels'):sub['branch']='folded'
    if enabled(plan):
        for branch in kids(root):
            for sub in walk(branch):
                if sub is not branch and sub.get('children') and not sub.get('labels'):
                    sub['branch']='folded'
    qroot=node(plan['title']+'—待确认规则');qroot['structureClass']='org.xmind.ui.logic.right'
    for q in qs.values():
        linked='、'.join(c['id'] for c in plan['cases'] if q['id'] in c['questions'])
        t=node(q['id']+' '+q['title'],[node('原文/缺口：'+q['gap']),node('需确认：'+q['ask']),node('影响用例：'+(q['impact']+'；' if q.get('impact') else '')+linked)],'\n'.join(text_ref(evidence[e]) for e in q['evidence']),['待确认']);t['branch']='folded';kids(qroot).append(t)
    if not qs:kids(qroot).append(node('本次设计未列待确认问题；不代表已实测或实现正确。'))
    info=node(plan['title']+'—依据与使用说明');info['structureClass']='org.xmind.ui.logic.right'
    if plan.get('schema_version')==2 and any(compact_case(plan,c) for c in plan['cases']):
        stats=plan['_coverage']
        kids(info).extend([
            node('阅读方式',[node('测试点后看简述；完整操作、预期和证据见测试点备注及Excel。'),node('待确认标识对应问题页；规则未明确时不预设结论。')]),
            node('设计概况',[node(f'{len(plan["cases"])}条用例，{stats["rules"]}条规则，{stats["items"]}项覆盖。'),
                node(f'待确认问题{len(qs)}项；待确认覆盖{stats["pending_planned"]}项，暂缓{stats["deferred"]}项，排除{stats["excluded"]}项。')]),
            node('使用限制',[node('全部未执行；推导和建议不等于原文明示要求。'),node('设计映射不证明需求完整覆盖；原生打开兼容性须另行验证。')],
                 '\n'.join(['读取范围：'+plan['source_review'],'需求版本：'+plan['requirement_version']]+plan.get('scope_notes',[])+['未知：']+plan.get('unknowns',[])+['失败尝试：']+plan.get('failed_attempts',[]))),
            node('详细资料位置',[node('Excel：规则与设计覆盖、执行准备、变更影响、检查详情。'),node('本页“使用限制”备注：读取范围、未确认信息及失败尝试。')])])
        warnings=presentation_report(plan)['warnings']
        if warnings:kids(info).append(node('部分测试点缺少简述，已保留完整旧展示。',note='\n'.join(warnings)))
    else:
        for title,lines in [('已确认的设计资料',[f'{len(plan["cases"])}条测试点；{len(qs)}项待确认规则。']+plan.get('scope_notes',[])),('推导与补充建议',['性质为边界派生/补充建议的条目是测试设计，不是原文明示或实测结论。']),('未知',plan.get('unknowns',[])+['当前代码处理、游戏执行结果、未批准口径均待验收AI填写。']),('失败的尝试',plan.get('failed_attempts',[]) or ['本轮未记录失败尝试。']),('交付边界',['原文件不修改；全部未执行；Excel总状态与三层结果按交接约定填写；本脚本未执行原生XMind打开校验。'])]:
            kids(info).append(node(title,[node(str(x)) for x in lines]))
    if plan.get('schema_version')==2 and not any(compact_case(plan,c) for c in plan['cases']):
        kids(info).append(node('需求版本与读取范围',[node(plan['requirement_version']),node(plan['source_review'])]))
        kids(info).append(node('设计覆盖说明',[node(plan['_coverage']['scope']),node(json.dumps({k:v for k,v in plan['_coverage'].items() if k not in ['links','scope']},ensure_ascii=False))]))
        for name,headers,widths,rows in tables(plan):
            kids(info).append(node(name,[node(str(row[0]),[node(h+'：'+str(v)) for h,v in zip(headers[1:],row[1:])]) for row in rows]))
        kids(info).append(node('简述待补充',note='\n'.join(presentation_report(plan)['warnings'])))
    theme={
      'map':{'id':uid(),'properties':{'svg:fill':'#FFFFFF','color-list':'#FFFFFF #24364B'}},
      'centralTopic':{'id':uid(),'properties':{'fo:font-family':'Microsoft YaHei','fo:font-size':'22pt','fo:color':INK,'svg:fill':'#DFEAF8','border-line-color':BLUE,'border-line-width':'3pt','shape-class':'org.xmind.topicShape.roundedRect'}},
      'mainTopic':{'id':uid(),'properties':{'fo:font-family':'Microsoft YaHei','fo:font-size':'15pt','svg:fill':'#DFEAF8','line-color':BLUE,'shape-class':'org.xmind.topicShape.roundedRect'}},
      'subTopic':{'id':uid(),'properties':{'fo:font-family':'Microsoft YaHei','fo:font-size':'12pt','line-color':BLUE,'line-class':'org.xmind.branchConnection.roundedfold'}}}
    if enabled(plan):
        theme['map']['properties'].update({'svg:fill':'#eee8e6','color-list':'#eee8e6 #2d221f'})
        for key in ('centralTopic','mainTopic','subTopic'):
            theme[key]['properties'].update({'fo:color':'#2d221f','svg:fill':'#c1aba5' if key!='subTopic' else '#e6ddda','line-color':'#88675e','border-line-color':'#88675e','shape-class':'org.xmind.topicShape.roundedRect','line-class':'org.xmind.branchConnection.curve'})
    sheets=[{'id':uid(),'class':'sheet','title':t,'rootTopic':r,'theme':theme} for t,r in [(plan['title']+'-测试点',root),('待确认规则',qroot),('依据与使用说明',info)]]
    ns='urn:xmind:xmap:xmlns:content:2.0';E.register_namespace('',ns)
    doc=E.Element('{'+ns+'}xmap-content',{'version':'2.0'})
    def xml_topic(parent,t,depth=0):
        attrs={'id':t['id'],'style-id':['central','main','sub'][min(depth,2)]}
        for a,b in [('branch','branch'),('structureClass','structure-class')]:
            if a in t:attrs[b]=t[a]
        e=E.SubElement(parent,'{'+ns+'}topic',attrs);E.SubElement(e,'{'+ns+'}title').text=t['title']
        if 'notes' in t:E.SubElement(E.SubElement(e,'{'+ns+'}notes'),'{'+ns+'}plain').text=t['notes']['plain']['content']
        if t.get('labels'):
            labels=E.SubElement(e,'{'+ns+'}labels')
            for v in t['labels']:E.SubElement(labels,'{'+ns+'}label').text=v
        if t.get('extensions'):
            es=E.SubElement(e,'{'+ns+'}extensions')
            for ext in t['extensions']:
                co=E.SubElement(E.SubElement(es,'{'+ns+'}extension',{'provider':ext['provider']}),'{'+ns+'}content')
                for v in ext['content']:E.SubElement(co,'{'+ns+'}'+v['name']).text=v['content']
        if t.get('children'):
            cs=E.SubElement(E.SubElement(e,'{'+ns+'}children'),'{'+ns+'}topics',{'type':'attached'})
            for c in kids(t):xml_topic(cs,c,depth+1)
    for s in sheets:
        se=E.SubElement(doc,'{'+ns+'}sheet',{'id':s['id'],**({'style-id':'sheet'} if enabled(plan) else {})});E.SubElement(se,'{'+ns+'}title').text=s['title'];xml_topic(se,s['rootTopic'])
    style_ns='urn:xmind:xmap:xmlns:style:2.0';styles=E.Element('{'+style_ns+'}xmap-styles',{'version':'2.0'});st=E.SubElement(styles,'{'+style_ns+'}styles')
    for sid,size,fill in [('central','22pt','#DFEAF8'),('main','15pt','#DFEAF8'),('sub','12pt','#FFFFFF')]:
        el=E.SubElement(st,'{'+style_ns+'}style',{'id':sid,'type':'topic'})
        E.SubElement(el,'{'+style_ns+'}topic-properties',{'{http://www.w3.org/1999/XSL/Format}font-family':'Microsoft YaHei','{http://www.w3.org/1999/XSL/Format}font-size':size,'{http://www.w3.org/2000/svg}fill':fill,'line-color':BLUE,'line-width':'1pt'})
    if enabled(plan):
        for style in st:
            props=style[0]
            props.set('{http://www.w3.org/2000/svg}fill','#c1aba5' if style.get('id')!='sub' else '#e6ddda')
            props.set('{http://www.w3.org/1999/XSL/Format}color','#2d221f')
            props.set('line-color','#88675e');props.set('line-class','org.xmind.branchConnection.curve')
            props.set('shape-class','org.xmind.topicShape.roundedRect')
        sheet_style=E.SubElement(st,'{'+style_ns+'}style',{'id':'sheet','type':'map'})
        E.SubElement(sheet_style,'{'+style_ns+'}map-properties',{'{http://www.w3.org/2000/svg}fill':'#eee8e6'})
    entries={'content.json':json.dumps(sheets,ensure_ascii=False).encode(),'content.xml':E.tostring(doc,encoding='utf-8',xml_declaration=True),'styles.xml':E.tostring(styles,encoding='utf-8',xml_declaration=True),'metadata.json':json.dumps({'dataStructureVersion':'3','creator':{'name':'game-requirement-test-docs-v5','version':'5.0.0'},'layoutEngineVersion':'5'}).encode()}
    entries['manifest.json']=json.dumps({'file-entries':{n:{} for n in entries}}).encode()
    manifest=E.Element('manifest',{'xmlns':'urn:xmind:xmap:xmlns:manifest:1.0'})
    for n in entries:E.SubElement(manifest,'file-entry',{'full-path':n,'media-type':'text/xml' if n.endswith('.xml') else 'application/json'})
    entries['META-INF/manifest.xml']=E.tostring(manifest,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in entries.items():z.writestr(n,b)
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        js=json.loads(z.read('content.json'));xml=E.fromstring(z.read('content.xml'))
        assert [(t['id'],t['title'],t.get('notes',{}).get('plain',{}).get('content','')) for s in js for t in walk(s['rootTopic'])]==[(t.attrib['id'],t.findtext('{'+ns+'}title'),t.findtext('{'+ns+'}notes/{'+ns+'}plain','')) for t in xml.iter('{'+ns+'}topic')]
        assert [(t['id'],t.get('labels',[]),[ch['id'] for ch in kids(t)]) for s in js for t in walk(s['rootTopic'])]==[(t.attrib['id'],[el.text for el in t.findall('{'+ns+'}labels/{'+ns+'}label')],[ch.attrib['id'] for ch in t.findall('{'+ns+'}children/{'+ns+'}topics/{'+ns+'}topic')]) for t in xml.iter('{'+ns+'}topic')]
        saved_cases={t['labels'][0]:t for t in walk(js[0]['rootTopic']) if t.get('labels')}
        assert set(saved_cases)=={c['id'] for c in plan['cases']}
        for c in plan['cases']:
            saved=saved_cases[c['id']]
            assert saved['title']==case_title(plan,c)
            assert [t['title'] for t in kids(saved)]==case_leaves(plan,c)
            assert saved['notes']['plain']['content']==notes(c,evidence,ideas)
    return case_nodes,root

def eval_formula(f,record):
    ts=re.findall(r'"[^"]*"|<>|=|\(|\)|,|[A-Z]+\d*',f.lstrip('='));pos=0
    def expr():
        nonlocal pos
        t=ts[pos];pos+=1
        if t.startswith('"'):v=t[1:-1]
        elif pos<len(ts) and ts[pos]=='(':
            pos+=1;args=[]
            while ts[pos]!=')':
                args.append(expr())
                if ts[pos]==',':pos+=1
            pos+=1
            if t=='IF':v=args[1] if args[0] else args[2]
            elif t=='AND':v=all(args)
            elif t=='OR':v=any(args)
            else:raise ValueError(t)
        else:v=record.get(re.sub(r'\d+$','',t),'')
        if pos<len(ts) and ts[pos] in ['=','<>']:
            op=ts[pos];pos+=1;rv=expr();v=(v==rv) if op=='=' else (v!=rv)
        return v
    result=expr();assert pos==len(ts);return result

def build_excel(plan,evidence,ideas,qs,nodes,path):
    return export_excel(plan,evidence,ideas,qs,nodes,path,notes,case_title,eval_formula)

def font_path():
    for p in [Path('C:/Windows/Fonts/msyh.ttc'),Path('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'),Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')]:
        if p.is_file():return str(p)
    return None
def previews(xlsx,root,support,plan,totalrow,start):
    fp=font_path();font=lambda size:ImageFont.truetype(fp,size) if fp else ImageFont.load_default()
    # Built from saved cells with explicit widths, not a claim of native rendering.
    w=openpyxl.load_workbook(xlsx,data_only=True)
    def shot(sn,r0,r1,cs,widths,name,expand=False):
        sh=w[sn];hs=[min(540,round((sh.row_dimensions[r].height or 15)*4/3)) for r in range(r0,r1+1)]
        if expand:
            measure=ImageDraw.Draw(Image.new('RGB',(1,1)))
            for ri,r in enumerate(range(r0,r1+1)):
                for ci,c in enumerate(cs):
                    count=0
                    for para in str(sh.cell(r,c).value or '').split('\n'):
                        line='';count+=1
                        for char in para:
                            if line and measure.textlength(line+char,font=font(14))>widths[ci]-16:count+=1;line=char
                            else:line+=char
                    hs[ri]=min(540,max(hs[ri],count*20+16))
        im=Image.new('RGB',(sum(widths),sum(hs)),'white');d=ImageDraw.Draw(im);y=0
        for ri,r in enumerate(range(r0,r1+1)):
            x=0
            for ci,c in enumerate(cs):
                cell=sh.cell(r,c);fill='#FFFFFF'
                merge=next((m for m in sh.merged_cells.ranges if m.min_row<=r<=m.max_row and m.min_col<=c<=m.max_col),None)
                if merge and (r!=merge.min_row or c!=merge.min_col):
                    x+=widths[ci];continue
                cell_width=sum(widths[j] for j,cc in enumerate(cs) if merge.min_col<=cc<=merge.max_col) if merge else widths[ci]
                if cell.fill.patternType=='solid' and cell.fill.fgColor.type=='rgb':fill='#'+cell.fill.fgColor.rgb[-6:]
                d.rectangle((x,y,x+cell_width,y+hs[ri]),fill=fill);text='' if cell.value is None else str(cell.value);f=font(14);color=INK
                if cell.font.color and cell.font.color.type=='rgb':color='#'+cell.font.color.rgb[-6:]
                lines=[]
                for para in text.split('\n'):
                    line=''
                    for char in para:
                        if line and d.textlength(line+char,font=f)>cell_width-16:lines.append(line);line=char
                        else:line+=char
                    lines.append(line)
                for j,line in enumerate(lines):
                    if j*20+26<=hs[ri]:d.text((x+8,y+7+j*20),line,font=f,fill=color)
                x+=widths[ci]
            y+=hs[ri]
        im.save(support/(name+'.png'))
    shot('AI实现检查',5,min(8,5+len(plan['cases'])),list(range(1,11)),WIDTHS,'Excel-主表')
    shot('验收说明与汇总',24,totalrow,list(range(1,16)),SUMMARY_WIDTHS,'Excel-汇总')
    shot('验收说明与汇总',3,21,list(range(1,11)),SUMMARY_WIDTHS[:10],'Excel-说明与基线')
    shot('检查详情',5,min(7,5+len(plan['cases'])),[1,14,15,16,17,18,19,27],[90,600,350,180,130,130,130,280],'Excel-检查详情')
    shot('检查详情',5,6,[1,3,4,5],[90,440,440,300],'Excel-完整设计',expand=True)
    focused=[r for r,c in enumerate(plan['cases'],6) if c.get('review_focus')]
    if focused:
        shot('检查详情',5,focused[0],[1,3,4],[90,650,440],'Excel-V3核对项',expand=True)
    if plan.get('questions'):shot('待确认规则',5,min(7,5+len(plan['questions'])),[1,2,3,4,5],[95,280,410,490,420],'Excel-规则')
    for name,headers,widths,rows in tables(plan):
        shot(name,5,min(7,5+len(rows)),list(range(1,min(6,len(headers))+1)),widths[:6],'Excel-'+name)
    branches=kids(root);height=max(500,80+len(branches)*100);im=Image.new('RGB',(1500,height),'white');d=ImageDraw.Draw(im)
    d.rounded_rectangle((30,height//2-40,350,height//2+40),radius=12,fill='#DFEAF8',outline=BLUE,width=3)
    d.text((45,height//2-16),plan['title'][:18],font=font(24),fill=INK)
    for i,b in enumerate(branches):
        y=60+i*100;d.line([(350,height//2),(440,y),(490,y)],fill=BLUE,width=2);d.rounded_rectangle((490,y-22,850,y+22),radius=8,fill='#DFEAF8',outline=BLUE)
        d.text((505,y-14),b['title'][:22],font=font(20),fill=INK);d.text((880,y-12),f'{sum(bool(t.get("labels")) for t in walk(b))} 条测试点，可展开',font=font(16),fill=INK)
    im.save(support/'XMind-概览.png')
    if enabled(plan):
        from template_preview import overview
        overview(root, support/'XMind-概览.png', font)
    compact=[c for c in plan['cases'] if compact_case(plan,c)]
    if compact:
        selected=[]
        for group in [compact[:2], [c for c in compact if not pending(c)][:2], [c for c in compact if len(c['xmind_points'])>1][:2],compact[-2:]]:
            for c in group:
                if c not in selected:selected.append(c)
        lines=[]
        for c in selected:
            lines.append((c['id']+'  '+' / '.join(output_path(c))+' / '+case_title(plan,c),True))
            lines.extend(('    → '+s,False) for s in c['xmind_points'])
        measure=ImageDraw.Draw(Image.new('RGB',(1,1)));layout=[]
        for text,is_title in lines:
            f=font(18 if is_title else 20);wrapped=[];line=''
            for ch in text:
                if line and measure.textlength(line+ch,font=f)>1430:wrapped.append(line);line=ch
                else:line+=ch
            wrapped.append(line)
            layout.append((wrapped,is_title))
        im=Image.new('RGB',(1500,50+sum(len(parts)*28+15 for parts,_ in layout)),'white');d=ImageDraw.Draw(im);y=25
        for wrapped,is_title in layout:
            f=font(18 if is_title else 20)
            for part in wrapped:d.text((25,y),part,font=f,fill=NAVY if is_title else INK);y+=28
            y+=15
        im.save(support/'XMind-简述抽样.png')

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--plan',required=True);p.add_argument('--evidence',required=True);p.add_argument('--out',required=True)
    p.add_argument('--require-v5',action='store_true',help='校验V5展示配置及V3必填设计')
    p.add_argument('--require-v3',action='store_true',help='新设计校验review_focus、acceptance_path及问题impact；旧计划兼容导出可省略')
    p.add_argument('--allow-legacy',action='store_true',help='仅已有导图保真转换，明确不计算v2覆盖')
    p.add_argument('--baseline',help='上次.support/design-baseline.json；只比较设计，不继承执行结果');a=p.parse_args()
    if (a.require_v3 or a.require_v5) and a.allow_legacy:p.error('--require-v3/--require-v5不能与--allow-legacy同时使用')
    ev=json.loads(Path(a.evidence).read_text(encoding='utf-8'))
    baseline=json.loads(Path(a.baseline).read_text(encoding='utf-8')) if a.baseline else None
    plan=prepare(json.loads(Path(a.plan).read_text(encoding='utf-8')),ev,a.allow_legacy,baseline);validate_focus(plan,a.require_v3);validate_display(plan,a.require_v5);template_report(plan);evidence,ideas,qs=check_plan(plan,ev)
    out=Path(a.out).resolve();out.mkdir(parents=True,exist_ok=True)
    title=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',plan['title']).strip(' .') or '游戏需求'
    xp=out/(title+'-测试用例.xmind');ep=out/(title+'-AI验收表.xlsx');support=out/'.support'
    if xp.exists() or ep.exists() or support.exists():raise SystemExit('输出已存在，拒绝覆盖；请选择新目录。')
    support.mkdir()
    nodes,root=build_xmind(plan,evidence,ideas,qs,xp)
    names,tr,totals,start=build_excel(plan,evidence,ideas,qs,nodes,ep)
    w,cache=validate_excel(ep,plan,nodes,notes,evidence,ideas,case_title);cs=plan['cases']
    assert w.sheetnames==names
    assert [cache['验收说明与汇总'].cell(tr,col).value for col in range(2,16)]==totals
    for name,headers,widths,rows in tables(plan):
        for ri,row in enumerate(rows,6):
            assert [(w[name].cell(ri,ci).value or '') for ci in range(1,len(headers)+1)]==[v or '' for v in row]
    for s in ev['sources']:assert sha(s['path'])==s['sha256']
    for f in [xp,ep]:
        with zipfile.ZipFile(f) as z:assert z.testzip() is None
    previews(ep,root,support,plan,tr,start)
    report={'files':[str(xp),str(ep)],'cases':len(cs),'questions':len(qs),'initial_statuses':{'blank':sum(not pending(c) for c in cs),'pending':sum(pending(c) for c in cs),'passed':0,'failed':0,'other':0},'checks':['JSON/XML主题一致','两文件逐项ID/详细内容一致；导图简述归属与完整备注一致','证据/Q引用存在','原件哈希未变','检查结果初始为空；最短验收路径为设计内容','下拉/筛选/冻结存在','统计缓存与独立计数一致','ZIP完整'],'native_application_open':'未执行','game_or_code_acceptance':'未执行','preview_engine':'Pillow辅助示意，非原生应用渲染','source_hashes':ev['sources'],'sha256':{f.name:sha(f) for f in [xp,ep]},'xmind_presentation':presentation_report(plan),'review_design':review_report(plan),'template_presentation':template_report(plan),'acceptance_contract':{'version':'5.0.0','statuses':STATUSES,'legacy_statuses':LEGACY_STATUSES,'layer_results':LAYER_RESULTS,'layer_sheet':'检查详情','layer_columns':['Q','R','S'],'main_sheet':'AI实现检查'}}
    if plan.get('schema_version')==2:
        report['design_coverage']=plan['_coverage'];report['changes']=plan['_changes']
        report['checks'].append('v2设计模型与三张附表一致；不保证语义覆盖完整')
        (support/'design-baseline.json').write_text(json.dumps(plan['_snapshot'],ensure_ascii=False,indent=2),encoding='utf-8')
    (support/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
