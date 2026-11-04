"""用已有模板、Markdown和流程图重建许祖嘉分稿。依赖python-docx。
从此脚本所在位置定位输入，不修改课程原模板，不生成任何测试或会议事实。
"""
from pathlib import Path
import copy,re
from docx import Document
from docx.shared import Pt,Mm,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
OUT=Path(__file__).resolve().parents[1]
REF=OUT.parent/'references/04课程报告模板.docx'
# 单独材料包也保留模板副本，以支持离线重建。
if not REF.exists():REF=Path(__file__).parent/'04课程报告模板.docx'

def configure_style(st,size=12,bold=False):
 st.font.name='Times New Roman';st.font.size=Pt(size);st.font.bold=bold;st.font.color.rgb=RGBColor(0,0,0)
 st.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'Songti SC')
 f=st.paragraph_format;f.line_spacing=Pt(20);f.space_before=Pt(0);f.space_after=Pt(0);f.first_line_indent=Pt(24);f.widow_control=True
 pp=st.element.find(qn('w:pPr'))
 if pp is not None:
  for el in list(pp.findall(qn('w:numPr'))):pp.remove(el)

def document(title):
 d=Document(REF)
 for el in list(d._element.body):
  if el.tag!=qn('w:sectPr'):d._element.body.remove(el)
 if 'Table Grid' not in d.styles:d.styles.element.append(copy.deepcopy(Document().styles['Table Grid'].element))
 if 'Title' not in d.styles:d.styles.add_style('Title',1)
 configure_style(d.styles['Normal'])
 for name in ['Heading 1','Heading 2','Heading 3']:
  configure_style(d.styles[name],12,True);f=d.styles[name].paragraph_format;f.first_line_indent=Pt(0);f.space_before=Pt(8);f.space_after=Pt(4);f.keep_with_next=True
 configure_style(d.styles['Title'],18,True);d.styles['Title'].paragraph_format.first_line_indent=Pt(0);d.styles['Title'].paragraph_format.space_after=Pt(10)
 d.core_properties.author='许祖嘉';d.core_properties.last_modified_by='许祖嘉';d.core_properties.title=title;d.core_properties.subject='课程报告分稿 XZJ-DRAFT-02'
 return d

def cell_text(s):return s.replace('`','')

def add_table(d,heads,rows,widths):
 t=d.add_table(rows=1,cols=len(heads));t.style='Table Grid';t.autofit=False
 for col,w in zip(t.columns,widths):col.width=Mm(w)
 for c,v in zip(t.rows[0].cells,heads):c.text=v
 t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
 for vals in rows:
  for c,v in zip(t.add_row().cells,vals):c.text=cell_text(v)
 for ri,row in enumerate(t.rows):
  row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
  for ci,c in enumerate(row.cells):
   c.width=Mm(widths[ci])
   if ri==0:
    shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'F2F2F2');c._tc.get_or_add_tcPr().append(shade)
   for p in c.paragraphs:
    f=p.paragraph_format;f.first_line_indent=Pt(0);f.line_spacing=Pt(16);f.space_before=Pt(3);f.space_after=Pt(3);f.keep_with_next=False
    for r in p.runs:
     r.font.size=Pt(10.5);r.font.bold=ri==0;r.font.name='Times New Roman';r._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'Songti SC')
 return t

def build_body():
 d=document('SimpleCADforDSH课程项目启动与计划报告');lines=(OUT/'正文.md').read_text().splitlines();i=0
 while i<len(lines):
  s=lines[i]
  if s.startswith('|'):
   rows=[]
   while i<len(lines) and lines[i].startswith('|'):
    rows.append([x.strip() for x in lines[i].strip('|').split('|')]);i+=1
   n=len(rows[0]);widths={2:[29,117],3:[31,55,60],4:[34,24,38,50]}[n]
   if rows[0][0]=='时点与负责人':widths=[34,64,48]
   if rows[0][0]=='对象':widths=[25,53,68]
   add_table(d,rows[0],rows[2:],widths);continue
  if s.startswith('!['):
   p=d.add_paragraph();f=p.paragraph_format;f.first_line_indent=Pt(0);f.line_spacing=1.;f.keep_with_next=True
   m=re.search(r'\]\(([^)]+)\)',s);r=p.add_run();r.add_picture(str(OUT/m[1]),width=Mm(146))
   dp=r._r.xpath('.//wp:docPr')
   if dp:dp[0].set('descr','项目目标流程：需求输入、生成或检索、几何与尺寸QA、有限编辑后复核、文件导出')
  elif s.startswith('# '):d.add_paragraph(s[2:],style='Title')
  elif s.startswith('#'):
   m=re.match(r'^(#+) (.*)',s);d.add_paragraph(m[2],style='Heading '+str(min(len(m[1])-1,3)))
  elif s.strip():
   p=d.add_paragraph(s)
   if s.startswith(('图1','表')):
    p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.keep_with_next=s.startswith('表');p.paragraph_format.space_after=Pt(4)
   if s.startswith(('[S','[E')):
    p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.alignment=0
    for r in p.runs:r.font.size=Pt(10.5)
  i+=1
 d.save(OUT/'正文.docx')

def build_processes():
 d=document('PMBOK第六版49子过程裁剪表');s=(OUT/'49子过程裁剪表.md').read_text()
 d.add_paragraph('PMBOK第六版49子过程裁剪表',style='Title')
 intro=s.split('## 整合管理')[0].split('\n\n')[1:]
 for p in intro:
  if p.strip():d.add_paragraph(p.strip())
 for block in re.split(r'^## ',s,flags=re.M)[1:]:
  name=block.splitlines()[0];d.add_paragraph(name,style='Heading 1')
  if name=='编制依据':
   for p in block.split('\n\n')[1:]:
    if p.strip():d.add_paragraph(p.strip())
   continue
  records=[]
  for line in block.splitlines():
   if re.match(r'^\| \d+\.\d+ ',line):
    r=[x.strip() for x in line.strip('|').split('|')];assert len(r)==8
    records.append([r[0]+' '+r[1]+'\n'+r[2],r[3]+'\n'+r[4],r[5]+'\n'+r[6]+'\n报告：'+r[7]])
  add_table(d,['编号与子过程／过程组','裁剪方式与项目理由','责任／产物／报告落点'],records,[37,60,49])
 d.save(OUT/'49子过程裁剪表.docx')

if __name__=='__main__':
 build_body();build_processes();print('正文与49子过程裁剪表已重建')
