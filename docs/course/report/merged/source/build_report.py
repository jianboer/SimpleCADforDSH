"""课程报告合稿。使用bundled python运行；原分稿只读。输出DOCX、来源映射与预算CSV。
渲染后可用 --toc-pdf <PDF> 回填同版目录页码，再次渲染至页码稳定。
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import load_workbook
from PIL import Image,ImageDraw,ImageFont
import copy,csv,re,json,hashlib,shutil,os,sys
OUT=Path(__file__).resolve().parents[1]; B=OUT.parent; AS=OUT/'assets'
X=B/'许祖嘉_课程报告分稿'; C=B/'常诚_课程报告分稿'; L=B/'刘子轩_课程报告分稿'; Y=B/'余智皓_课程报告分稿'
REF=B/'references/04课程报告模板.docx'; FILENAME='SimpleCADforDSH课程报告初稿'
blocks=[]; body=True; table_no=0;figure_no=0; titlemap={}; source='统稿新增'; bodytext=[]
def clean(s):
 s=s.replace('〔E03—E06〕','[1—3,14]').replace('〔E04、E07〕','[2,4,14]')
 s=re.sub(r'〔E[\d、—E]+〕','[14]',s)
 refs={'S1':'1','S2':'2','S3':'3','S4':'4','S5':'5','S6':'6','S7':'7','S8':'8','S9':'9','S11':'10','S12':'11','E1':'12'}
 s=s.replace('[S1—S5]','[1—5]')
 for a,b in refs.items():s=s.replace('['+a+']','['+b+']')
 s=s.replace('刘子轩在4.4展开','4.4展开').replace('其他三人已经提交或统稿完成','相关功能已经通过验收')
 s=s.replace('李健博维护开发计划及看板；常诚、刘子轩、余智皓分别编制范围、进度成本、质量风险沟通分稿；许祖嘉确认跨章节口径，报告编写职责不替代八人开发职责。','李健博维护开发计划及看板，各工作包负责人提供实现、自测与交接材料，许祖嘉确认跨章节口径和验收结论。')
 s=s.replace('许祖嘉收齐其余分稿后统一术语、图表及引用，复算跨表数字，核查各人AI记录再形成总报告；目前不宣称相关功能已经通过验收。','最终报告拟由八名成员共同复核范围、预算和验收依据，保留各项实际执行记录。报告材料的编制完成不等于软件能力已通过验收。')
 s=s.replace('余智皓、刘子轩区分原任务已含、新增、条件触发并映射负责人；PO确认取舍，不直接叠加或抵扣','风险22人时拆为12已含、9条件开发及1管理；开发全部触发为75/72人时，PO确认取舍')
 s=s.replace('22人时须映射既有任务；风险状态与阈值不视为实测结果','22人时已映射为12已含、9条件开发及1管理，归类仍待PO确认；风险状态与阈值不视为实测结果')
 s=s.replace('风险应对22人时还需逐项去重','风险应对22人时已按计划分为12已含、9条件开发和1管理，需由PO确认后执行')
 s=re.sub(r'(?<![A-Za-z0-9])E(?:0[1-9]|1[0-6])(?![A-Za-z0-9])', '[14]', s)
 s=s.replace('目前可核算的项目工作量仅覆盖首迭代','原讨论三可追溯的已估算工作量仅覆盖首迭代')
 s=s.replace('目前不宣称','本报告不宣称')
 s=s.replace('AT17-02预期结果仍需澄清','AT17-02原文不充分，按讨论四澄清提出可判定预期，PO正式确认状态另记')
 s=s.replace('AT17-02沿用讨论四已明确的','AT17-02依据讨论四澄清提出').replace('已完成79条AT与WBS','已完成79条AT的计划追踪关联，与WBS')
 s=s.replace('下列活动按D1—D6相对开发日安排。','下列活动按D1—D6相对开发日安排。79条AT追踪状态均为未执行，不据此计算产品测试通过率。')
 s=s.replace('全量追踪表列明79条AT','附录B的追踪摘要及随附源表列明79条AT')
 for old,new in titlemap.items():s=s.replace(old,new)
 return s

def p(s,style='Normal'):
 if not s.strip():return
 s=clean(s.strip());blocks.append({'type':'p','text':s,'style':style,'source':source,'body':body})
 if body and style=='Normal':bodytext.append(s)
def h(s,level=2):p(s,'Heading '+str(level))
def table(head,rows,widths=None,caption=None):
 global table_no
 if caption:
  table_no+=1;p('表'+str(table_no)+' '+caption,'Caption')
 blocks.append({'type':'table','head':head,'rows':[[clean(str(v)) for v in r] for r in rows],'widths':widths,'source':source,'body':body})
def pic(path,caption):
 global figure_no
 figure_no+=1;dest=AS/(f'figure-{figure_no}'+Path(path).suffix)
 if Path(path).resolve()!=dest.resolve():shutil.copy2(path,dest)
 blocks.append({'type':'image','path':str(dest.relative_to(OUT)),'source':source,'body':body})
 p('图'+str(figure_no)+' '+caption,'Caption')
def rows(t):
 return [[c.text for c in row.cells] for row in t.rows]
def csvrows(path):return list(csv.DictReader(Path(path).open(encoding='utf-8-sig')))

def from_md(text):
 lines=text.splitlines();i=0
 while i<len(lines):
  s=lines[i].strip()
  if s.startswith('|'):
   rs=[]
   while i<len(lines) and lines[i].strip().startswith('|'):
    rs.append([x.strip() for x in lines[i].strip().strip('|').split('|')]);i+=1
   table(rs[0],rs[2:]);continue
  if s.startswith('!['):
   path=re.search(r'\]\(([^)]+)',s)[1];pic(X/path,'项目目标流程 生成检索与质量核对闭环')
  elif s.startswith('#'):
   m=re.match(r'(#+) (.*)',s);h(m[2],min(len(m[1])-1,3))
  elif re.match(r'^表\d+',s):
   global table_no
   old=re.match(r'^表\d+',s)[0];table_no+=1;titlemap[old]='表'+str(table_no);p(re.sub(r'^表\d+',titlemap[old],s),'Caption')
  elif not s.startswith('图1'):p(s)
  i+=1

# 统一来源字典与前半正文
source='许祖嘉_课程报告分稿/正文.md'
s=X.joinpath('正文.md').read_text().split('## AI使用说明')[0]
s=s[s.index('## 1 项目概述'):]
from_md(s)
source='统稿新增 范围与WBS整理';h('4.2 范围管理')
h('4.2.1 收集需求',3)
p('需求来自课堂讨论形成的工程场景、用户故事与验收约束，而非已经开展的大规模企业调查。讨论一明确面向学生、初级设计人员和有CAD资产的小型团队；讨论二将目标转为25个故事及79条AT；讨论三安排任务、责任与依赖；讨论四通过扑克牌澄清复杂度；讨论五分析数据、接口、资源与交付风险。每次转化都应保留原编号，便于追溯为什么安排这项工作，以及以什么结果认定它完成。[1—5]')
p('需求整理以同一零件闭环为单位。输入支撑板描述时，工程尺寸和单位先成为结构化约束，再驱动建模；显示参数、几何检查、修改参数及导出应围绕同一个模型版本。只实现孤立按钮不能说明闭环成立。检索输入则需要样例模型、参考图及来源、规格和版本，结果加载后仍要核对尺寸和几何，而不是将相似度分值等同于工程合格。')
p('需求粒度分为目标、故事和核验任务。G目标描述阶段结果，US描述用户可获得的能力，AT给出可重复执行的条件与预期；WBS表达实现和管理交付物。AT编号不是工作量单位，同一工作包可以支撑多个AT，重复关联不重复记工时。首迭代唯一故事主责沿用原任务表，故事主责、测试执行人和PO接收人分别记录，后续迭代责任在启动前复核。')
cd=Document(C/'4.2.docx');story=[]
newsp=[13,8,13,3,8,8,13,5,3,3,3,5,5,5,2,3,3,5,5,5,5,3]
for row in cd.tables[0].rows:
 vals=[c.text for c in row.cells]
 if re.fullmatch(r'US\d{2}',vals[0]):
  n=int(vals[0][2:]);story.append([vals[0],vals[1],vals[4],vals[5],str(newsp[n-1]) if n<=22 else '三格式合并5',vals[6]])
table(['故事','用户能力','优先级','原点数','重估点数','迭代'],story,[14,52,16,16,29,19],'25项用户故事及估算口径')
p('故事表保留讨论三原估算，讨论四重估作为复杂度变化记录。首迭代由58点升至85点，第二迭代由40点调整至41点；三种格式导出由三故事11点变为组合估算5点，不能将组合估算填入每个故事后再相加。已有原工时只有首迭代66小时，因此重估并不自动说明原排期仍可行，4.3和4.4分别分析容量与成本。')
h('4.2.2 范围定义与项目范围说明书',3)
p('交付范围包括多模态CAD生成与参数展示、样例资产检索与详情、几何和尺寸QA、清单内参数编辑及基础CAD操作、STEP/STL/GLB导出与兼容核验，以及复现说明、固定演示剧本和管理材料。开发基于已有插件demo补齐能力，不承诺构建完整通用CAD、企业级PDM或敏感企业资产平台。原版本规划中的故障诊断和多源数据融合不属于本项目。')
p('迭代一对应Version 0.5，形成US01—US12生成、检索及QA基础增量；迭代二对应Version 0.8，形成US13—US22参数编辑与基础操作；迭代三对应Version 1.0，完成US23—US25导出、总回归、演示和归档。版本名称仅表示计划交付关口，不代表这些版本已经发布，原第二迭代18点的文字误写统一按表内40点修正。')
table(['范围边界','本期安排与判定'],[
 ['生成与检索','固定简单零件与可获得资产；图像缺尺度或视图矛盾时提示补充，不猜测真实尺寸'],
 ['编辑清单','保留US13—US22及原AT，最终迁移清单由PO确认；不泛化为任意复杂CAD编辑'],
 ['几何与文件','几何合法和尺寸符合分别判定；导出须在目标查看工具中打开并核对主体几何'],
 ['明确排除','从零训练基础模型、重建企业资产库、通用几何修复、无限编辑序列及完整CAD/PDM'],
 ['先决条件','合法样例、固定提示词、可复现环境、模型服务和必要额度可用；不成立时记录影响并变更'],
 ['接收规则','全部AT与G1—G5阶段目标分别举证；未执行、失败与限制如实记录，PO确认接收']
],[28,118],'范围边界与接收规则')
p('范围控制以“新增需求是否改变交付物、接口、人员工时或验收依据”为判断入口。对任意曲面开孔、复杂几何修复或变半径圆角等扩展，不因现有内核提供相关函数就视为已纳入范围。主责人先列出影响故事、任务、资源和测试，再由PO作出接受、延期或拒绝的决定，并同步WBS和看板。原有承诺未变更前仍需完整验收，不能通过删除验收项使进度看起来达标。')
p('模型来源和训练数据来源分别记录。资产元数据中的来源说明不能证明所用生成模型的训练数据组成，未知信息应显示未知。自建样例需保存建模参数、对应图片和预期尺寸；公开资产需保存许可和版本信息。样例目录准备完成后，先核对目标资产是否能读取和被检索，再开展排序效果评估，避免仅有一份文件清单却没有可用检索基准。')
h('4.2.3 创建WBS及WBS字典',3)
p('WBS以交付物为组织依据，根项0.0下分1.1项目管理与过程资料、1.2生成检索与QA、1.3参数编辑与基础操作、1.4导出演示与交付。管理活动覆盖全周期，开发前建立需求基线；功能工作包依赖已有模型、接口和样例。完整字典见附录A，包含工作内容、责任、时间、依赖和质量要求，字典中的百分比是验收目标，不是实测成功率。')
p('最小工作包用于分派和核验，首迭代仍使用I1原任务号记录人时。工作包汇总不再次叠加到子任务工时。原字典部分表头将1.3误写为生成检索成果、将1.3.3重复写为轮廓与拉伸，本稿按其正文定义统一为参数编辑与基础CAD操作、贯穿孔与布尔并集。原审批栏若无独立确认记录，只能作为来源信息，不能推导计划已正式获批。')
# WBS tree re-rendered from drawio text using its node labels, split 4 branches
import xml.etree.ElementTree as ET,html
xml=ET.parse(C/'WBS结构树.drawio');labels=[]
for node in xml.iter('mxCell'):
 v=node.get('value','');v=re.sub('<[^>]*>',' ',html.unescape(v)).strip()
 if v:labels.append(v)
for node in xml.iter('mxCell'):
 v=node.get('value','')
 m=re.search(r'(?<![0-9.])([1-4])(?:\.([1-6]))?(?=[ .\u4e00-\u9fff<]|$)',v)
 if m:
  revised='1.'+m[1]+('.'+m[2] if m[2] else '')
  v=v[:m.start()]+revised+v[m.end():];node.set('value',v)
xml.write(AS/'统一WBS结构树.drawio',encoding='utf-8',xml_declaration=True)
(AS/'wbs-drawio-labels.json').write_text(json.dumps(labels,ensure_ascii=False,indent=2))
fontfile='/System/Library/Fonts/Supplemental/Songti.ttc'
font=ImageFont.truetype(fontfile,30)
def figure(title,lines,path):
 im=Image.new('RGB',(1800,110+len(lines)*75),'white');dr=ImageDraw.Draw(im)
 dr.text((35,25),title,font=ImageFont.truetype(fontfile,38),fill='black')
 for n,s in enumerate(lines):
  y=100+n*75;dr.rectangle((35,y,1760,y+60),outline='#888888',width=2);dr.text((55,y+10),s,font=font,fill='black')
 im.save(path)
wbsdict=[]
for t in cd.tables[5:]:
 kv={r.cells[0].text:r.cells[1].text for r in t.rows[1:]};wbsdict.append(kv)
for prefix in ['1.1','1.2','1.3','1.4']:
 records=[x for x in wbsdict if x.get('WBS编号','').startswith(prefix)]
 figure(prefix+' '+records[0]['WBS名称'],[x['WBS编号']+' '+x['WBS名称'] for x in records[1:]],AS/(prefix+'.png'))
 pic(AS/(prefix+'.png'),'WBS '+prefix+'分支（按drawio及配套字典重排）')
h('4.2.4 需求追踪与范围确认',3)
p('追踪关系按G目标、REQ分类、US、AT、WBS、迭代、主责和测试执行任务建立。附录B给出25项需求映射和79条AT的输入、操作、预期及执行状态；原始CSV保留CodeArts平台ID和链接。开发中模型参数或接口约定变化，应先标出受影响AT，再更新正反例和回归集，不能只在正文更改一句描述。')
p('AT17-02原文仅写“执行开孔”，不足以判断异常是否正确处理。本稿依据讨论四明确的范围，提出无效参数应提示且原模型不变的验收预期。AT20-03的过大倒角距离等步骤属于质量分稿补充的计划，保留原始字段并由PO正式确认。参数读取也按实际实现表述为从结构化描述与脚本PARAMS取得可编辑值，不承诺从任意STEP恢复历史参数。')

# Import Liu/Yu editable content, normalizing all table styles and extracting native images.
def import_doc(path):
 global source,figure_no,table_no
 source=str(path.relative_to(B));d=Document(path)
 # caption map first so narrative references use final numbering
 tn=table_no;fn=figure_no
 for para in d.paragraphs:
  s=para.text.strip();m=re.match(r'^(表|图)(\d+\.\d+-\d+)(?=\s)',s)
  if m:
   if m[1]=='表':tn+=1;titlemap[m[0]]='表'+str(tn)
   else:fn+=1;titlemap[m[0]]='图'+str(fn)
 for el in d.iter_inner_content():
  if hasattr(el,'text'):
   s=el.text.strip()
   for node in el._p.xpath('.//a:blip'):
    rid=node.get(qn('r:embed'));part=d.part.related_parts[rid];dest=AS/('embedded-'+hashlib.sha256(part.blob).hexdigest()[:12]+'.'+part.content_type.split('/')[-1].replace('jpeg','jpg'));dest.write_bytes(part.blob)
    blocks.append({'type':'image','path':str(dest.relative_to(OUT)),'source':source,'body':body})
   if not s:continue
   if re.match(r'^(表|图)4\.\d+-\d+\s',s):
    if s.startswith('表'):table_no+=1
    else:figure_no+=1
    p(s,'Caption');continue
   if re.match(r'^4\.[34567](?:\.\d+)? ',s):h(s,3 if re.match(r'^4\.\d+\.\d+ ',s) else 2)
   else:p(s)
  else:
   rs=rows(el);table(rs[0],rs[1:])
import_doc(L/'正文.docx')
# Remove obsolete budget paragraphs and replace risk accounting wording.
for b in blocks:
 if b['source'].startswith('刘子轩') and b['type']=='p':
  t=b['text']
  if t.startswith('后两迭代只有WBS') or '后两迭代只有WBS责任' in t:
   b['text']=t.replace('后两迭代只有WBS责任与时间窗，没有活动人时及个人容量估算，因此尚不能称为资源平衡已通过。','本次据现有WBS补做后两迭代暂定工时，见4.4.3；个人超载仍需PO确认调整。')
  if t.startswith('若10小时增量全部发生'):
   b['text']='若9开发人时增量全部发生，开发任务为66+9=75人时，比72人时有效容量多3小时；另有1小时PO管理增量，合计76人时，按60元折算为4560元。管理增量不挤占六名开发人员容量，全部触发也不是期望损失。原人员关键链路满载，PO需按技能和依赖调整任务或范围，不能仅用总余量冲抵。'
  if t.startswith('现金成本用实际数量'):
   b['text']='现金成本按实际数量和单价核算，原始材料尚无API账单、GPU时长或许可报价。本次增加项目级暂定估算，不将缺失实付信息写为0元；计费条件确定前，新增输入仅用于情景讨论，不能作为批准支出。'
  if t.startswith('现阶段可将'):
   b['text']=t.replace('工作簿','原工作簿')
for b in blocks:
 if b['source'].startswith('刘子轩') and b['type']=='table':
  for r in b['rows']:
   for i,s in enumerate(r):
    r[i]=s.replace('10h全部触发时超过分散余量6h，需先换出工作或缩减范围。','开发增量9h，75/72缺口3h；管理增量1h另列，均须PO确认。')
source='统稿新增 暂定预算 MERGE-BUDGET-02';h('4.4.3 全项目暂定预算与资源平衡',3)
p('本节补做预算情景MERGE-BUDGET-02，所有新增人时、调用数量及单价输入均为暂定估算，待许祖嘉确认。首迭代66人时沿用原任务表；第二、三迭代根据已有工作包逐项分解实现、自测、联调和资料整理，不从故事点换算。它反映按原责任分工执行的资源压力，而非已经批准的新开发基线。开发人员会议和协调已在每迭代每人预留4小时中扣除，不再次计作净开发任务。')
people=['常诚','刘子轩','余智皓','严翊航','侯永森','庞佳泽']
budget=[
 ('1.3.1','参数编辑与模型重建',2,[7,3,3,1],[4,3,2,1,3,1]),
 ('1.3.2','矩形轮廓与实体拉伸',2,[5,2,1,1],[0,0,1,1,2,5]),
 ('1.3.3','贯穿孔与布尔并集',2,[6,3,2,1],[0,2,2,3,2,3]),
 ('1.3.4','边圆角与边倒角',2,[5,3,2,1],[3,2,2,1,2,1]),
 ('1.3.5','撤销与模型树',2,[5,2,2,1],[2,3,2,1,1,1]),
 ('1.3.6','编辑功能联调与验收',2,[0,3,10,2],[2,2,3,5,2,1]),
 ('1.4.1','STEP导出与兼容验证',3,[5,3,3,1],[1,1,3,4,1,2]),
 ('1.4.2','STL与GLB导出',3,[4,2,2,1],[0,1,2,3,0,3]),
 ('1.4.3','全流程回归与缺陷关闭',3,[2,3,8,1],[2,2,3,3,2,2]),
 ('1.4.4','发布与固定剧本演示',3,[2,1,3,2],[1,1,3,1,1,1]),
 ('1.4.5','最终验收与资料归档支持',3,[0,1,1,4],[1,1,1,1,1,1])]
for r in budget:assert sum(r[3])==sum(r[4])
with (OUT/'source/暂定预算明细.csv').open('w') as f:
 w=csv.writer(f);w.writerow(['WBS','工作包','迭代','实现h','自测h','联调h','资料h']+people+['总h','状态'])
 for id,n,it,parts,alloc in budget:w.writerow([id,n,it]+parts+alloc+[sum(parts),'计划估算'])
table(['工作包','实现','自测','联调','资料','合计h'],[[r[0]+' '+r[1]]+r[3]+[sum(r[3])] for r in budget],[61,17,17,17,17,17],'迭代二三暂定自下而上工时')
allocs=[[12,12,10,10,11,11]]+[[sum(r[4][i] for r in budget if r[2]==it) for i in range(6)] for it in [2,3]]
assert sum(allocs[1])==71 and sum(allocs[2])==49
p('工时估计考虑复用已有内核与导出链路：重建需验证参数、孔位、回滚；特征操作需准备有效实体并检查错误输入；回归按接口和模型状态组织，而不是再实现一遍功能。表中1.3.6和1.4.3主要是跨模块联调及复测，因此实现工时较少。各人的分配是对原WBS主责及协作关系的暂定展开，供团队核对技能和实际可投入时间；主责人负责交付，不意味着每项工作都必须亲自完成。')
table(['成员','迭代1h','迭代2h','迭代3h','每迭代容量h'],[[people[i]]+[a[i] for a in allocs]+[12] for i in range(6)]+[['合计']+[sum(a) for a in allocs]+[72]],[34,28,28,28,28],'逐人容量核对（后两迭代为暂定）')
p('迭代二总计71小时低于72小时，但侯永森24小时、庞佳泽16小时，分别超出个人有效容量12和4小时；余智皓12小时也没有余量。其余成员的17小时余量不能自动解决这16小时超载，需确认技能、接口交接及活动顺序。建议开工前将可独立验证和共用接口工作拆给有余量成员，迁移主责继续复核内核操作；若不能形成同窗口可执行分配，则应调整交付清单或工期。现有甘特图保持原六日目标窗口，不宣称已完成资源平衡。')
p('第三迭代暂定49开发人时，其中余智皓12小时、严翊航13小时，严翊航超1小时。可评估由庞佳泽协助一小时格式导出核验，但更改仍须本人和PO确认。首迭代的9小时条件开发风险不能因为后两迭代存在余量就直接后移：若它阻断G1—G3，后续编辑和导出仍受影响，必须先分析依赖。所有重排同时更新WBS、甘特图、任务责任和预算，不能只在成本表移动数字。')
p('管理人时暂按PO与SM每人每迭代8小时估算，各含计划与范围确认、协调跟踪、评审接收和归档；两人合计每迭代16小时、全期48小时，不计入六名开发人员每迭代72小时容量。开发基准66+71+49=186小时，人力机会成本11160元；加管理48小时后234小时、14040元。首迭代原66小时对应3960元仍保留，新增16管理小时对应960元为本次暂定输入，不追溯为原已批准预算。')
table(['迭代','开发h','管理h 暂定','机会成本元','调用次数 暂定'],[['1',66,16,4920,500],['2',71,16,5220,300],['3',49,16,3900,200],['合计',186,48,14040,1000]],[22,25,32,34,33],'分迭代人力与现金用量情景')
p('现金预算不预设具体服务商报价，暂以1000次完整任务调用作数量情景，分别500、300、200次，单次有效调用费q元需根据实际模型、输入输出Token、图片和失败重试计费核对。仅为便于讨论，以q=1元这一分析输入得到1000元调用情景，取q=0.5、1、3元时分别为500、1000、3000元。这些是预算敏感性假设，不是官方报价；如果采用Token计费，应将q替换为每次实际Token成本后再汇总。')
p('基准假设复用现有设备和团队自建样例、不新增硬件与付费许可；若改用云GPU，以8小时压力测试/备用推理情景计8g元，g为待确认小时单价；购买数据或许可另列L元。全期现金表达式为1000q+8g+L，云GPU未触发时不计该条件项。基准取q=1且不触发云GPU及付费许可时，机会成本加现金为15040元；不具备资源报价或许可依据的支出仍待确认，不能由这个情景取得采购授权。')
p('风险全部触发的追加为9开发小时加1管理小时，机会成本追加600元，全期244小时、14640元；若q=1且无其他资源增项，合计15640元。此项只追加未包含的10小时，不再次增加已含的12小时，也不将其视为概率加权储备金。路线B的初始60小时属于经济路线比较模型，已与项目预算分别核算，不能再叠加到186开发小时上。')
p('预算控制拟由主责人每次交接报告实际工时、剩余工时和调用消耗，SM按WBS汇总，PO比较暂定基线与完成预测。若费用数量、活动剩余量或资源约束变化，先说明导致变化的输入及对验收的影响，保留变更前后版本。实际账单、用量和工时尚未产生时，本节只用于计划审阅；团队确认估算、报价和付款承担方式后，才形成正式预算及支出上限。')
import_doc(Y/'正文.docx')
# CodeArts gate control adds a real corrective plan, never fabricated test fixes.
source='统稿新增 CodeArts质量门禁处理'
# insert block directly after 4.5.2 preceding risk heading later
extra='CodeArts本轮全量检查执行成功，但198项问题中有2项严重问题，门禁未通过，不能写为代码质量验收完成。拟由集成组按问题位置和规则分派给模块主责，先定位两项严重问题，再区分真实缺陷、规则适配和重复告警；保留原始问题与处理理由。修改需关联提交并执行受影响AT和代码复检，只有严重问题达到门禁要求且回归证据充分，PO才考虑放行。当前没有修复和复检结果，保留未解决状态。'
i=next(i for i,b in enumerate(blocks) if b['type']=='p' and b['text']=='4.6 风险管理')
blocks.insert(i,{'type':'p','text':extra,'style':'Normal','source':source,'body':True});bodytext.append(extra)
# replace Liu old capacity statement which is in existing sentence and remove extra raw Fig references old mapping
for b in blocks:
 if b['type']=='p':
  b['text']=b['text'].replace('完整公式和敏感性见工作簿。','完整公式和敏感性见原工作簿及本稿4.4。')
# AI summary without assigning drafting duties in final report
body=False;source='四人AI使用记录与本次统稿操作'
h('AI使用说明',1)
p('本报告使用OpenAI Codex辅助资料核对、需求追踪、经济计算、风险归类、图表制作和版式整理。项目选题、目标、角色、范围、预算确认、验收和最终结论由学生承担。本次统稿修复了错贴项目内容、估算汇总和容量口径，并保留未执行测试、未通过门禁及未批准预算的真实状态。')
table(['环节','AI提供的辅助','已有人工判断与本次处理','最终责任'],[
 ['需求与范围','整理故事、AT及追踪映射','常诚记录逐条对照讨论二，纠正目标口径；本次统一25故事及79AT、WBS错名和范围边界','PO确认范围和异常验收预期'],
 ['经济与进度','生成公式、敏感性、工时与甘特图','刘子轩记录核对66/72与估算差异；本次分开9开发与1管理，补暂定预算并暴露超载','PO与主责确认估算，不视为已批准'],
 ['质量风险与CodeArts','核对台账、截图及工具记录','余智皓补充风险会议方法并记录登录权限和操作授权；保留198项问题及失败门禁','真实操作由参与学生负责；验收由PO确认'],
 ['论证与管理裁剪','比较资料、提出49过程安排','许祖嘉确认角色、经济分析方式；人工最终复核尚需八人共同完善','项目选题和关键决定由学生承担'],
 ['合稿与排版','按模板整合、复算和渲染检查','本次统稿未代替学生宣布新预算获批、AT通过或严重问题已解决','八名成员共同审阅后形成最终稿']
],[22,32,63,29],'AI辅助用途及人工决策边界')
p('原始AI记录随各分稿保留，记录工具、输入材料、建议、人工核查修改及决策人。本报告中的新增人时及现金用量假设须由许祖嘉和相关主责审阅，不能把AI生成的计算或文本直接当作本人的已执行成果。答辩时可由学生用US、AT、WBS和版本解释范围、估算、风险和计划取舍。')
h('参考资料',1)
refs=[
 '本组讨论一《软件开发梦想秀》：目标、章程、角色与阶段要求；仓库基线ce5b0203。',
 '本组讨论二《编写用户故事及其验收测试标准》：US01—US25与79条AT；同一仓库基线。',
 '本组讨论三《项目计划会议》《首迭代用户故事与任务分配》《各组用户故事估算与优先级》：WBS、任务、人时与依赖。',
 '本组讨论四《扑克牌估算工作量》：重估及范围澄清，原始出牌记录保持原样。',
 '本组讨论五《风险管理最佳实践》：24候选、投票图、十大风险和22人时估算。',
 '项目params.py、cad_cli.py、dsh-plugin/index.js及models目录；许祖嘉2026年11月2日静态核查记录，非AT执行结果。',
 '《课程报告要求及评分标准2026—2027》与《04课程报告模板》，本报告references目录。',
 'FreeCAD官方功能说明：https://freecad.github.io/Website/features/ ；许祖嘉分稿访问记录2026年11月2日。',
 'build123d官方文档：https://build123d.readthedocs.io/en/stable/index.html ；许祖嘉分稿访问记录2026年11月2日。',
 '许祖嘉角色及分析方式确认、余智皓会议方法补充记录，见两人原始证据索引与AI使用记录；本次未补造独立会议。',
 'Onshape FeatureScript MCP官方说明：https://www.onshape.com/en/features/featurescript-mcp ；许祖嘉分稿检索记录2026年11月2日。',
 '经济情景XZJ-ECO-01，许祖嘉《经济分析假设与计算》及刘子轩《进度与成本分析.xlsx》，2026年10月30日至11月2日分稿。',
 '常诚《4.2.docx》《需求追踪表.xlsx》及WBS结构树.drawio；原审批栏未独立核验。',
 '余智皓《正文.docx》、全量AT、十大风险、阶段目标、协作、CodeArts CSV及截图，2026年11月2日材料包。',
 'CodeArts项目f6d21d4cbbd24bab9665dd3f98385838（华北-北京四）；Repo 3110839，基线ce5b0203ee446ef25ced27cda7f856e8197d8403。',
 'CodeArts Check任务2777245d88b846f185edf5e4fea8af07，Job d6368a47e4ce4229babae2f663446c0e，2026年11月2日00:06:46。原始报告位于平台任务#71216773附件；本次以本地截图与索引核查。',
 '本次暂定预算MERGE-BUDGET-02：新增人时及现金数量为统稿情景，未批准；完整公式与分配见source/暂定预算明细.csv及source/预算计算.json。'
]
for i,r in enumerate(refs,1):p(f'[{i}] {r}')
# Full WBS dictionary, compact one table per level-one branch, avoiding 26 repeated project titles.
source='常诚_课程报告分稿/4.2.docx'
h('附录A 完整WBS字典',1)
p('以下按原WBS归并重复字段，工作内容、依赖、责任、质量与时窗均予保留；百分比为目标。原审批人和2026年10月6日日期未经独立核验，本稿不将其作为本次预算或修订批准凭据。新增预算见4.4.3，不改写原字典历史工时。')
for parent in ['0.0','1.1','1.2','1.3','1.4']:
 rec=[x for x in wbsdict if x.get('WBS编号')==parent or (parent!='0.0' and x.get('WBS编号','').startswith(parent+'.'))]
 h(parent+' '+rec[0]['WBS名称'],2)
 for x in rec:

  if x['WBS编号']!=parent:h(x['WBS编号']+' '+x['WBS名称'],3)
  rr=[]
  for key in ['父级WBS编号/名称','子级WBS编号/名称','工作内容描述','责任主体','质量要求','时间要求','依赖关系']:
   rr.append([key,x.get(key,'原稿未提供')])
  rr.append(['估算来源','首迭代对照I1任务；迭代二三暂定拆分见4.4.3。原责任与时窗依原字典，开工前复核。'])
  table(['字段','内容'],rr,[27,119])
# Requirements tracking Excel copy into native table; discard artificial Column headers
source='常诚_课程报告分稿/需求追踪表.xlsx';h('附录B 需求与验收追踪',1)
wb=load_workbook(C/'需求追踪表.xlsx',read_only=True,data_only=True);r=list(wb['2'].values)
table(['目标／REQ／US','AT范围','WBS／迭代','故事主责／核验责任'],[[str(x[0])+'/'+str(x[1])+'/'+str(x[2]),str(x[3]),str(x[5])+'／'+str(x[6]),str(x[7])+'／'+str(x[4])] for x in r[2:] if x[2]],[36,33,33,44],'25项需求追踪摘要')
ats=csvrows(Y/'图表及源文件/全量AT验收追踪表.csv');keys=list(ats[0]);assert len(ats)==79
(OUT/'source/AT列名.json').write_text(json.dumps(keys,ensure_ascii=False))
for us in sorted({list(a.values())[0] for a in ats}):
 h(us+'验收计划',2); ar=[]
 for a in ats:
  v=list(a.values())
  if v[0]!=us:continue
  ar.append([v[1],v[2]+'\n'+v[3],v[4]+('\n计划补充：'+v[5] if v[5] else '')])
 table(['AT编号','输入与操作','预期及补充'],ar,[23,59,64])
 p('本故事AT状态：未执行。责任、WBS、时窗及CodeArts任务ID见同版全量AT源表。')
# 49 processes grouped
source='许祖嘉_课程报告分稿/49子过程裁剪表.md';h('附录C PMBOK第六版49子过程裁剪',1)
p('配置管理纳入整合管理，不另算第50项。过程组数量为启动2、规划24、执行10、监控12、收尾1；未来过程是计划，不是已完成证明。')
process_count=0
for part in re.split(r'^## ',(X/'49子过程裁剪表.md').read_text(),flags=re.M)[1:]:
 name=part.splitlines()[0];rs=[]
 for line in part.splitlines():
  if re.match(r'^\| \d+\.\d+ ',line):
   r=[x.strip() for x in line.strip('|').split('|')];process_count+=1
   rs.append([r[0]+' '+r[1]+'\n'+r[2],r[3]+'\n'+r[4],r[5]+'\n'+r[6]+'\n报告：'+r[7]])
 if rs:h(name,2);table(['编号／过程／过程组','裁剪与项目理由','责任／产物／落点'],rs,[35,65,46])
assert process_count==49
source='余智皓_课程报告分稿/图表及源文件';h('附录D 阶段目标与风险应对',1)
stages=csvrows(Y/'图表及源文件/阶段目标独立核验表.csv')
(OUT/'source/阶段字段.json').write_text(json.dumps(stages,ensure_ascii=False,indent=2))
for a in stages:
 vals=list(a.values());h(vals[0]+' 阶段核验',2)
 table(['字段','内容'],[[k,v] for k,v in a.items() if v],[33,113])
risks=csvrows(Y/'图表及源文件/十大风险登记册.csv')
for a in risks:
 h(a['风险编号']+' '+a['风险名称'],2)
 table(['字段','内容'],[[k,a[k]] for k in ['责任人','触发条件','应对措施','WBS关联','任务关联','AT或阶段目标','原应对估算人时','已含基线人时','条件或管理增量人时','增量分类','计入口径与限制','原表状态','状态来源日期说明']],[34,112])
 p('原表状态仅按来源保留，无本次风险已关闭或措施有效的推断。')
from release_revision import apply
release_plan=apply(blocks,OUT,budget,people)
# Preserve source assets + map, never modify originals.
for who in [X,C,L,Y]:
 for f in who.rglob('*'):
  if f.is_file() and f.suffix.lower() in ['.csv','.xlsx','.drawio','.svg']:
   dest=OUT/'source'/'originals'/who.name/f.relative_to(who);dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(f,dest)
inputs={str(p.relative_to(B)):hashlib.sha256(p.read_bytes()).hexdigest() for who in [X,C,L,Y] for p in who.rglob('*') if p.is_file() and p.name!='.DS_Store' and not p.name.startswith('~$')}
(OUT/'source/input-sha256.json').write_text(json.dumps(inputs,ensure_ascii=False,indent=2))
(OUT/'source/content.json').write_text(json.dumps(blocks,ensure_ascii=False,indent=2))
(OUT/'source/来源映射.json').write_text(json.dumps([{'序号':i+1,'类型':b['type'],'位置':b.get('text','')[:80],'来源':b['source']} for i,b in enumerate(blocks)],ensure_ascii=False,indent=2))
(OUT/'source/预算计算.json').write_text(json.dumps({'版本':'MERGE-BUDGET-02 计划预算','人员顺序':people,'开发工时分配':allocs,'开发合计':186,'管理合计':48,'机会成本率':60,'机会成本基准':14040,'条件增量开发':9,'条件增量管理':1,'含风险机会成本':14640,'现金公式':'1000q + 8g(云GPU触发时) + L(许可触发时)','q等于1无其他触发时基准总额':15040,'q等于1含风险无其他触发总额':15640},ensure_ascii=False,indent=2))
# DOCX authoritative template clone
D=Document(REF)
for el in list(D._element.body):
 if el.tag!=qn('w:sectPr'):D._element.body.remove(el)
# Template section & footer remain, body replaced per slot contract.
if 'Table Grid' not in D.styles:D.styles.element.append(copy.deepcopy(Document().styles['Table Grid'].element))
for name in ['Title','Caption']:
 if name not in D.styles:D.styles.add_style(name,1)
def style(st,size=12,bold=False):
 st.font.name='Times New Roman';st.font.size=Pt(size);st.font.bold=bold;st.font.color.rgb=RGBColor(0,0,0)
 rf=st.element.get_or_add_rPr().get_or_add_rFonts();rf.set(qn('w:eastAsia'),'宋体');rf.set(qn('w:ascii'),'Times New Roman');rf.set(qn('w:hAnsi'),'Times New Roman')
 for a in ['asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme']:rf.attrib.pop(qn('w:'+a),None)
 f=st.paragraph_format;f.line_spacing=Pt(20);f.first_line_indent=Pt(24);f.space_before=Pt(0);f.space_after=Pt(0);f.widow_control=True
 pp=st.element.get_or_add_pPr()
 for node in list(pp.findall(qn('w:numPr'))):pp.remove(node)
style(D.styles['Normal'])
for i,size in [(1,18),(2,15),(3,12)]:
 st=D.styles['Heading '+str(i)];style(st,size,True);f=st.paragraph_format;f.first_line_indent=Pt(0);f.keep_with_next=True;f.space_before=Pt(10);f.space_after=Pt(5)
style(D.styles['Caption'],10.5);D.styles['Caption'].paragraph_format.first_line_indent=Pt(0)
style(D.styles['Title'],36);D.styles['Title'].paragraph_format.first_line_indent=Pt(0)
for name in ['TOC 1','TOC 2']:
 if name not in D.styles:D.styles.add_style(name,1)
 style(D.styles[name],11);D.styles[name].paragraph_format.first_line_indent=Pt(0 if name=='TOC 1' else 12)
for section in D.sections:
 section.page_width=Mm(210);section.page_height=Mm(297);section.top_margin=Mm(25.4);section.bottom_margin=Mm(25.4);section.left_margin=Mm(31.75);section.right_margin=Mm(31.75)
 for para in section.footer.paragraphs:
  para.alignment=1
  for run in para.runs:
   run.font.name='Times New Roman';run.font.size=Pt(10)
# Cover explicit and compact keeps original title text, centered.
for txt,size,before,after in [('工业软件课程项目',36,85,16),('开发启动、计划报告',36,0,60),('项目名称：SimpleCADforDSH',18,0,12),('面向DeepSeek Harness的智能参数化CAD生成检索插件',15,0,42),('2课堂12组',15,0,20),('成员姓名：李健博、许祖嘉、常诚、刘子轩、',15,0,2),('侯永森、庞佳泽、余智皓、严翊航',15,0,45),('2026年11月',15,0,0)]:
 para=D.add_paragraph();para.alignment=1;f=para.paragraph_format;f.first_line_indent=Pt(0);f.line_spacing=Pt(size*1.5);f.space_before=Pt(before);f.space_after=Pt(after)
 run=para.add_run(txt);run.font.size=Pt(size);run.font.name='Times New Roman';run._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'宋体')
D.add_page_break();para=D.add_paragraph('目录',style='Heading 1');para.paragraph_format.page_break_before=False
# cached static TOC, updated by supplied PDF and followed by fixed pagebreak
heads=[b for b in blocks if b['type']=='p' and ((b.get('body') and b['style'].startswith('Heading')) or (not b.get('body') and b['style']=='Heading 1'))]
toc_pages={}
if '--toc-pdf' in sys.argv:
 from pypdf import PdfReader
 pdf=PdfReader(sys.argv[sys.argv.index('--toc-pdf')+1])
 for pageid,page in enumerate(pdf.pages):
  text=page.extract_text() or '';compact=re.sub(r'\s','',text)
  # page2/3 are TOC, body starts after them; matching only exact heading lines avoids narrative references
  if pageid<3:continue
  lines={re.sub(r'\s','',line) for line in text.splitlines()}
  for b in heads:
   key=b['text'];norm=re.sub(r'\s','',key)
   if norm in lines and key not in toc_pages:toc_pages[key]=pageid+1
# Consistently 3 TOC pages (about 66 heading rows) split at fixed 27 rows.
for idx,b in enumerate(heads):
 if idx and idx%27==0:D.add_page_break()
 para=D.add_paragraph(style='TOC 1' if b['style']=='Heading 1' else 'TOC 2');para.paragraph_format.line_spacing=Pt(20)
 tabs=para.paragraph_format.tab_stops
 from docx.enum.text import WD_TAB_ALIGNMENT,WD_TAB_LEADER
 tabs.add_tab_stop(Mm(145),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
 para.add_run(b['text']+'\t'+str(toc_pages.get(b['text'],'—')))
D.add_page_break()
# body renderer
for b in blocks:
 if b['type']=='p':
  para=D.add_paragraph(b['text'],style=b['style'])
  if b['style']=='Heading 1':para.paragraph_format.page_break_before=True
  if b['style']=='Caption':
   para.paragraph_format.alignment=1;para.paragraph_format.keep_with_next=b['text'].startswith('表');para.paragraph_format.space_after=Pt(4)
  if b['text'].startswith('[') and not b['body']:
   para.paragraph_format.first_line_indent=Pt(0)
  if b['text'].startswith('附录') and b['style']=='Heading 1':para.paragraph_format.page_break_before=True
 elif b['type']=='image':
  im=Image.open(OUT/b['path']);w,h=im.size;maxh=175
  width=min(146,146*maxh/(146*h/w))
  para=D.add_paragraph();f=para.paragraph_format;f.first_line_indent=Pt(0);f.line_spacing=1.;f.keep_with_next=True;para.alignment=1
  para.add_run().add_picture(str(OUT/b['path']),width=Mm(width))
 else:
  n=len(b['head']);widths=b['widths'] or ({2:[34,112],3:[33,57,56],4:[30,25,46,45],5:[28,26,26,33,33],6:[23,41,20,18,24,20]}.get(n,[146/n]*n))
  if b['head']==['核验对象','判定依据','责任与时点','记录要求']:widths=[26,54,32,34]
  if b['head']==['里程碑','计划窗口','主要核验内容','责任/决策']:widths=[28,23,64,31]
  if b['head'][0]=='编号与风险':widths=[35,18,23,50,20]
  if b['head'][0]=='方案':widths=[23,32,25,23,20,23]
  t=D.add_table(rows=1,cols=n);t.style='Table Grid';t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
  for col,w in zip(t.columns,widths):col.width=Mm(w)
  for c,val in zip(t.rows[0].cells,b['head']):c.text=str(val)
  t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
  for rr in b['rows']:
   for c,val in zip(t.add_row().cells,rr):c.text=str(val)
  for ri,row in enumerate(t.rows):
   row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
   for ci,c in enumerate(row.cells):
    c.width=Mm(widths[ci])
    if ri==0:
     shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'F2F2F2');c._tc.get_or_add_tcPr().append(shade)
    for para in c.paragraphs:
     f=para.paragraph_format;f.first_line_indent=Pt(0);f.line_spacing=Pt(16);f.space_before=Pt(2);f.space_after=Pt(2);f.keep_with_next=(b['head']==['字段','内容'] and ri<4)
     for run in para.runs:
      run.font.name='Times New Roman';run.font.size=Pt(10.5);run.font.bold=ri==0
      run._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'宋体')
# Disable inherited auto numbering (explicit semantic heading numbers in text).
for para in D.paragraphs:
 if para.style.name.startswith('Heading'):
  pp=para._p.get_or_add_pPr();np=OxmlElement('w:numPr');ni=OxmlElement('w:numId');ni.set(qn('w:val'),'0');np.append(ni);pp.append(np)
D.core_properties.title='SimpleCADforDSH课程项目开发启动与计划报告';D.core_properties.author='SimpleCADforDSH 第12组';D.core_properties.last_modified_by='SimpleCADforDSH 第12组';D.core_properties.comments=''
D.save(OUT/(FILENAME+'.docx'))
normal=''.join(b['text'] for b in blocks if b['type']=='p' and b.get('body') and b['style']=='Normal')
count=len(re.findall(r'[\u4e00-\u9fff]',normal))
(OUT/'source/构建统计.json').write_text(json.dumps({'正文叙述汉字数不含表图附录':count,'US':len(story),'AT':len(ats),'子过程':process_count,'表数':table_no,'图数':figure_no,'目录条目':len(heads),'已回填页码':len(toc_pages)},ensure_ascii=False,indent=2))
print('DOCX built',count,'正文叙述汉字',len(heads),'目录项',len(toc_pages),'页码')
