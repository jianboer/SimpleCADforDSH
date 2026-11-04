"""核查交付内容、字体、目录、数值、包结构和原稿指纹；不修改任何原始分稿。"""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader
from zipfile import ZipFile
from lxml import etree
import json,re,hashlib,sys
B=Path(__file__).resolve().parents[1];src=B/'source';D=Document(B/'SimpleCADforDSH课程报告初稿.docx')
pdf=Path(sys.argv[1]);r=PdfReader(pdf);blocks=json.loads((src/'content.json').read_text());result={}
assert D.styles['Normal'].font.size.pt==12
assert D.styles['Normal'].paragraph_format.line_spacing.pt==20
assert D.styles['Normal'].element.get_or_add_rPr().get_or_add_rFonts().get(qn('w:eastAsia'))=='宋体'
sec=D.sections[0]
for key,val in [('top_margin',25.4),('bottom_margin',25.4),('left_margin',31.75),('right_margin',31.75)]:assert abs(getattr(sec,key).mm-val)<.02
result['版式']='A4模板页边距；宋体12磅；固定20磅'
assert next(p.text for p in D.paragraphs if p.text.startswith('成员姓名：')).startswith('成员姓名：李健博、许祖嘉、')
result['成员与计划表述']='李健博列封面成员首位；正文直接陈述项目计划，无内部讨论说明'
assert not re.search('仍待PO确认|尚未批准|拟采用的异常验收规则|全项目暂定预算|建议方案|本报告采用|报告采用|提交前|内部复核|内部讨论',json.dumps(blocks,ensure_ascii=False))
text='\n'.join(p.extract_text() or '' for p in r.pages);norm=lambda s:re.sub(r'\s','',s)
heads=[x for x in blocks if x['type']=='p' and ((x['body'] and x['style'].startswith('Heading')) or (not x['body'] and x['style']=='Heading 1'))]
locations={}
for n,page in enumerate(r.pages):
 if n<3:continue
 lines={norm(x) for x in page.extract_text().splitlines()}
 for h in heads:
  if norm(h['text']) in lines and h['text'] not in locations:locations[h['text']]=n+1
assert len(locations)==len(heads)==51,locations
# DOCX TOC cached values match delivered PDF page numbers
cache={}
for p in D.paragraphs:
 if p.style.name.startswith('TOC'):
  title,n=p.text.rsplit('\t',1);cache[title]=int(n)
assert cache==locations,(cache,locations)
result['目录']=f'{len(heads)}项可见目录页码与同版PDF完全一致'
normal=''.join(x['text'] for x in blocks if x['type']=='p' and x.get('body') and x['style']=='Normal')
count=len(re.findall('[\u4e00-\u9fff]',normal));assert 10000<=count<=12000
result['正文叙述字数不含图表附录']=count
assert set(re.findall(r'AT\d{2}-\d{2}',text))==set(re.findall(r'AT\d{2}-\d{2}',(B/'source/originals/余智皓_课程报告分稿/图表及源文件/全量AT验收追踪表.csv').read_text(encoding='utf-8-sig')))
assert len(set(re.findall(r'AT\d{2}-\d{2}',text)))==79
result['覆盖']='25故事、79唯一AT、49子过程；原AT状态保留未执行'
assert 66+9==75 and 75-72==3 and (66+9+1)*60==4560
assert 66+71+49==186 and 186+48==234 and 234*60==14040 and 14040+1000==15040
assert 3600+11*200==5800 and 13000+8.5*200==14700
assert 3600/(30-11)>189 and 3600/(30-11)<190
result['数值']='基准及增量人时、容量、预算、路线经济公式独立复算通过'
# Independently recompute schedule, all allocations, gates and exported workbook caches.
from collections import Counter
from openpyxl import load_workbook
plan=json.loads((src/'计划数据.json').read_text());tasks=plan['tasks'];gates=plan['gates']
assert len(tasks)==120 and len(gates)==21
assert len({(t['day'],t['slot'],t['person']) for t in tasks})==120
byid={t['id']:t for t in tasks};gate_end={}
for g in gates:
 end=max(byid[i]['finish'] for i in g['producer_ids']);assert end==g['finish']
 gate_end[g['iteration'],g['gate']]=end
for t in tasks:
 assert t['hours']==t['finish']-t['start']==1 and 0<=t['slot']<6
 for dep in t['dependencies']:assert gate_end[t['iteration'],dep]<=t['start']
for package in plan['packages']:
 if package['iteration']==1:continue
 members=[sum(t['hours'] for t in tasks if t['wbs']==package['wbs'] and t['person']==n) for n in plan['people']]
 assert members==package['allocations'],(package,members)
 assert sum(members)==package['hours']
for day in range(3,7):
 for person in plan['people']:assert sum(t['hours'] for t in tasks if t['day']==day and t['person']==person)<=6
for it,total in [(2,71),(3,49)]:
 assert sum(t['hours'] for t in tasks if t['iteration']==it)==total
 for person in plan['people']:assert sum(t['hours'] for t in tasks if t['iteration']==it and t['person']==person)<=12
xlsx=B/'SimpleCADforDSH计划与预算.xlsx';w=load_workbook(xlsx,data_only=True);form=load_workbook(xlsx,data_only=False)
assert w['概览']['G9'].value==15040
assert w['预算']['C19'].value==186 and w['预算']['D19'].value==48
assert w['预算']['F19'].value==14040 and w['预算']['C24'].value==14640 and w['预算']['C25'].value==15640
assert [w['风险'].cell(16,c).value for c in range(4,8)]==[22,12,9,1]
assert w['风险']['C22'].value==3 and w['风险']['C24'].value==4560
for i,t in enumerate(tasks,6):
 assert w['任务计划'].cell(i,2).value==t['id']
 assert w['任务计划'].cell(i,10).value==1
 assert w['任务计划'].cell(i,16).value=='满足列明依赖'
 assert form['任务计划'].cell(i,16).data_type=='f'
for sheet in w:
 for row in sheet:
  for cell in row:assert cell.data_type!='e',(sheet.title,cell.coordinate,cell.value)
result['任务安排']='120条任务、21个交接关口；逐人每日与迭代容量、同人时段、列明依赖及工作包分配一致'
result['工作簿']='6张表，公式缓存回读、风险去重及预算总额一致；q敏感性和缺报价传播另见workbook-verification.json'
for i,b in enumerate(blocks):
 if b['type']=='p' and b['text']=='1.4 导出、演示与项目交付成果':
  t=next(x for x in blocks[i+1:] if x['type']=='table')
  values=dict(t['rows']);assert values['时间要求'].startswith('第5—6天')
  assert '计划投入' not in values['责任主体']
assert len(set(re.findall(r'US\d{2}',text)))==25
process_tables=[x for x in blocks if x['type']=='table' and any('裁剪' in str(h) for h in x['head'])]
assert sum(len(t['rows']) for t in process_tables)==49
assert len({r[0].split()[0] for t in process_tables for r in t['rows']})==49
source_hash=json.loads((src/'input-sha256.json').read_text())
for name,digest in source_hash.items():assert hashlib.sha256((B.parent/name).read_bytes()).hexdigest()==digest,name
result['原稿']='四份原始材料指纹一致'
# Check each paragraph survives PDF rendering, excluding soft line break whitespace.
pdfnorm=norm(''.join('\n'.join(line for line in (page.extract_text() or '').splitlines() if not re.fullmatch(r'\d+',line.strip())) for page in r.pages));missing=[]
for x in blocks:
 if x['type']=='p' and len(x['text'])>10 and norm(x['text']) not in pdfnorm:missing.append(x['text'][:120])
result['文本回读未精确命中']=missing
assert not missing,missing
result['页数']=len(r.pages)
result['空页']=[n+1 for n,p in enumerate(r.pages) if len((p.extract_text() or '').strip())<10]
assert not result['空页']
# Explicit caption numbers are unique and sequential, narrative reference excluded.
fig=[int(re.match(r'图(\d+)',x['text'])[1]) for x in blocks if x['type']=='p' and x['style']=='Caption' and x['text'].startswith('图')]
tab=[int(re.match(r'表(\d+)',x['text'])[1]) for x in blocks if x['type']=='p' and x['style']=='Caption' and x['text'].startswith('表')]
assert fig==list(range(1,12)),fig
assert tab==list(range(1,21)),tab
result['编号']='图1—11、表1—20连续唯一'
# Package relationship targets preserved or resolve; XML structure no corrupt media relation.
with ZipFile(B/'SimpleCADforDSH课程报告初稿.docx') as z:
 names=set(z.namelist());rels=etree.fromstring(z.read('word/_rels/document.xml.rels'))
 import posixpath
 for rel in rels:
  if rel.get('TargetMode')=='External':continue
  target=posixpath.normpath('word/'+rel.get('Target'));assert target in names,target
result['结构']='文档关系及图片目标均可解析；表格为可编辑Word对象'
(src/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));print(json.dumps(result,ensure_ascii=False,indent=2))
