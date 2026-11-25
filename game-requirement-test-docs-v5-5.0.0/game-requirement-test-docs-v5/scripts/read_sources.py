"""Read selected design documents with stable locators; never infer business rules."""
import argparse, hashlib, json, posixpath, re, zipfile
from pathlib import Path
from urllib.parse import unquote
import xml.etree.ElementTree as E

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
N={'s':S,'r':R,'a':'http://schemas.openxmlformats.org/drawingml/2006/main','xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def resolve_part(base,target):
    target=unquote(target)
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.dirname(base)+'/'+target)
def relationships(z,part):
    rp=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
    return {x.attrib['Id']:x.attrib for x in E.fromstring(z.read(rp))} if rp in z.namelist() else {}
def extract_xlsx(p,category,emit,images,media):
    with zipfile.ZipFile(p) as z:
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ss=[''.join(t.text or '' for t in x.findall('.//s:t',N)) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
        rel=relationships(z,'xl/workbook.xml')
        for sheet in E.fromstring(z.read('xl/workbook.xml')).find('s:sheets',N):
            part=resolve_part('xl/workbook.xml',rel[sheet.attrib['{'+R+'}id']]['Target'])
            xml=E.fromstring(z.read(part));sn=sheet.attrib['name']; cells={}
            for c in xml.findall('.//s:sheetData/s:row/s:c',N):
                v=c.find('s:v',N);value=v.text or '' if v is not None else ''
                if c.attrib.get('t')=='s' and value:value=ss[int(value)]
                elif c.attrib.get('t')=='inlineStr':value=''.join(t.text or '' for t in c.findall('.//s:t',N))
                f=c.find('s:f',N)
                if f is not None:value=f'公式={f.text or ""}；缓存={value}'
                if value:
                    cells[c.attrib['r']]=value;emit(category,p,f'{sn}!{c.attrib["r"]}',value)
            if category=='ideas':
                byrow={}
                for addr,val in cells.items():byrow.setdefault(int(re.search(r'\d+',addr).group()),{})[re.match(r'[A-Z]+',addr).group()]=val
                for row,cs in byrow.items():
                    if re.fullmatch(r'[A-Z]{2,8}-\d+',cs.get('A','')):
                        emit(category,p,f'{sn}!A{row}:J{row}',' | '.join(f'{k}={v}' for k,v in cs.items()),idea_id=cs['A'])
            for rr in relationships(z,part).values():
                if rr.get('TargetMode')=='External':continue
                target=resolve_part(part,rr['Target'])
                if 'comments' in rr.get('Type','').lower():
                    for c in E.fromstring(z.read(target)).findall('.//s:comment',N):
                        emit(category,p,f'{sn}!{c.attrib.get("ref","?")} 批注',''.join(t.text or '' for t in c.findall('.//s:t',N)))
                if not rr.get('Type','').endswith('/drawing'):continue
                drel=relationships(z,target)
                for anchor in E.fromstring(z.read(target)):
                    origin=anchor.find('xdr:from',N)
                    xy={x.tag.split('}')[-1]:x.text for x in origin} if origin is not None else {}
                    loc=f'{sn} 图片锚点(0起算) {json.dumps(xy,ensure_ascii=False)}'
                    for blip in anchor.findall('.//a:blip',N):
                        br=drel.get(blip.attrib.get('{'+R+'}embed',''))
                        if not br or br.get('TargetMode')=='External':continue
                        mp=resolve_part(target,br['Target']);dest=media/(sha(p)[:10]+'_'+Path(mp).name)
                        dest.write_bytes(z.read(mp));rec=emit(category,p,loc,f'内嵌图：{mp}，需人工/视觉阅读，不能据路径推断内容')
                        images.append({'evidence':rec['id'],'file':str(dest),'locator':loc,'source':str(p)})
                    texts=''.join(t.text or '' for t in anchor.findall('.//a:t',N))
                    if texts:emit(category,p,loc+' 文本框',texts)

def extract(p,category,emit,images,media):
    ext=p.suffix.lower()
    if ext=='.xlsx':return extract_xlsx(p,category,emit,images,media)
    if ext=='.xmind':
        with zipfile.ZipFile(p) as z:
            if 'content.json' in z.namelist():
                def walk(t,path):
                    current=path+[t.get('title','')];txt=t.get('title','');note=t.get('notes',{}).get('plain',{}).get('content','')
                    emit(category,p,' > '.join(current),txt+('\n备注：'+note if note else ''))
                    for topics in t.get('children',{}).values():
                        for c in topics:walk(c,current)
                for sh in json.loads(z.read('content.json')):walk(sh['rootTopic'],[sh.get('title','')])
            else:
                ns={'x':'urn:xmind:xmap:xmlns:content:2.0'}
                def walk(t,path):
                    title=t.findtext('x:title','',ns);cur=path+[title]
                    note=t.findtext('x:notes/x:plain','',ns)
                    if not note:
                        html=t.find('x:notes/x:html',ns)
                        if html is not None:note=''.join(html.itertext())
                    emit(category,p,' > '.join(cur),title+('\n备注：'+note if note else ''))
                    for c in t.findall('x:children/x:topics/x:topic',ns):walk(c,cur)
                for sh in E.fromstring(z.read('content.xml')).findall('x:sheet',ns):walk(sh.find('x:topic',ns),[sh.findtext('x:title','',ns)])
            # Inline images may carry requirements; save resources without guessing semantics.
            for name in z.namelist():
                if Path(name).suffix.lower() in ['.png','.jpg','.jpeg','.webp'] and not name.startswith('Thumbnails/'):
                    dest=media/(sha(p)[:10]+'_'+hashlib.sha256(name.encode()).hexdigest()[:8]+Path(name).suffix)
                    dest.write_bytes(z.read(name));rec=emit(category,p,name,'XMind内嵌图，需查看其关联主题和图中文字')
                    images.append({'evidence':rec['id'],'file':str(dest),'locator':name,'source':str(p)})
        return
    if ext=='.docx':
        with zipfile.ZipFile(p) as z:
            ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for part in [n for n in z.namelist() if re.fullmatch(r'word/(document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml',n)]:
                for i,para in enumerate(E.fromstring(z.read(part)).findall('.//w:p',ns),1):
                    def revision_text(el):
                        name=el.tag.split('}')[-1]
                        if name in ['t','delText']:return el.text or ''
                        content=''.join(revision_text(x) for x in el)
                        if name in ['ins','del','moveFrom','moveTo']:
                            return '['+name+'修订：'+content+']'
                        if name=='tab':return '\t'
                        if name in ['br','cr']:return '\n'
                        return content
                    text=revision_text(para)
                    if text:emit(category,p,f'{part} 段落{i}',text)
            for name in z.namelist():
                if name.startswith('word/media/') and not name.endswith('/'):
                    dest=media/(sha(p)[:10]+'_'+Path(name).name);dest.write_bytes(z.read(name))
                    rec=emit(category,p,name,'DOCX内嵌媒体，需核对图片与正文定位/修订含义')
                    images.append({'evidence':rec['id'],'file':str(dest),'locator':name,'source':str(p)})
        return
    if ext=='.pdf':
        from pypdf import PdfReader
        for i,page in enumerate(PdfReader(p).pages,1):emit(category,p,f'第{i}页',page.extract_text() or '[无可提取文字：需视觉阅读或OCR]')
        return
    if ext in ['.txt','.md','.csv','.tsv']:
        for i,line in enumerate(p.read_text(encoding='utf-8-sig').splitlines(),1):
            if line.strip():emit(category,p,f'第{i}行',line)
        return
    raise ValueError(f'不支持直接读取{ext}，请使用对应读取工具，不能猜测内容：{p}')

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--workspace',required=True);p.add_argument('--requirement',required=True)
    p.add_argument('--ideas');p.add_argument('--style',action='append',default=[]);p.add_argument('--out',required=True)
    a=p.parse_args();workspace=Path(a.workspace).resolve();skill=Path(__file__).resolve().parents[1]
    req=Path(a.requirement).resolve();out=Path(a.out).resolve()
    if out.exists():raise SystemExit(f'证据文件已存在，未覆盖：{out}')
    if not req.is_file():raise SystemExit(f'需求文件不存在：{req}')
    if a.ideas:ideas=Path(a.ideas).resolve()
    else:
        candidates=[f for f in (workspace/'outputs').rglob('通用游戏测试思路.xlsx') if not f.name.startswith('~$')]
        groups={}
        for f in candidates:groups.setdefault(sha(f),f)
        if len(groups)>1:raise SystemExit('多个不同思路版本，需--ideas明确指定：'+str(list(map(str,groups.values()))))
        ideas=next(iter(groups.values())) if groups else skill/'assets/通用游戏测试思路.xlsx'
    if not ideas.is_file():raise SystemExit(f'思路表不存在：{ideas}')
    styles=[Path(x).resolve() for x in a.style]
    if not styles:
        # Only deterministic nearby samples, not a claim to select semantically closest files.
        folders=[f for f in workspace.iterdir() if f.is_dir() and '测试留存' in f.name]
        for folder in folders:
            for name in ['登录战令-测试用例.xmind','砍价活动-测试点.xmind','拯救萝莉-测试用例.xmind']:
                if (folder/name).is_file():styles.append(folder/name)
        if not styles:styles=sorted((skill/'assets/style-samples').glob('*.xmind'))
    out.parent.mkdir(parents=True,exist_ok=True);media=out.parent/(out.stem+'_media');media.mkdir(exist_ok=True)
    records=[];images=[];sources=[]
    def emit(cat,file,loc,text,**kw):
        r={'id':f'R{len(records)+1:06d}','category':cat,'file':str(file),'locator':loc,'text':str(text),**kw};records.append(r);return r
    for cat,file in [('requirement',req),('ideas',ideas)]+[('style',s) for s in styles]:
        if not file.is_file():raise SystemExit(f'源文件不存在：{file}')
        sources.append({'category':cat,'path':str(file),'sha256':sha(file)})
        extract(file,cat,emit,images,media)
    result={'sources':sources,'records':records,'images':images,'read_limits':[
        '提取文本与图片不等于已理解或已验收；图片与修订需单独检查。',
        'DOCX修订以ins/del/moveFrom/moveTo标注，未决定接受或拒绝；需确认有效版本。',
        'PDF只提取逐页文字，未提取或渲染图片；需另用可用PDF工具视觉阅读全部相关页。',
        '旧版XLS、WMF/EMF、公式渲染、图表/SmartArt及在线页面需额外读取工具；未读部分标未知。']}
    out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'evidence_file':str(out),'sources':sources,'records':len(records),'images':len(images),'idea_rows':sum('idea_id' in r for r in records)},ensure_ascii=False,indent=2))
if __name__=='__main__':main()
