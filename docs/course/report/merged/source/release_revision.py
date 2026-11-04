"""2026-11-04提交初稿收口：建议资源方案、任务依赖、图形、正文整理。"""
from pathlib import Path
import json,csv,re,copy,html
from collections import Counter
from PIL import Image,ImageDraw,ImageFont
import xml.etree.ElementTree as ET
PEOPLE=['常诚','刘子轩','余智皓','严翊航','侯永森','庞佳泽']
S2=[['31','31','35','31','31','34','35','34','34','36',None,'36'],
    ['31','35','35','31','31','33','34','35','33','34','36','36'],
    ['31','35','31','33','34','32','33','35','34','36','36','36'],
    ['35','34','32','31','36','36','36','33','33','33','36','36'],
    ['31','31','31','32','32','33','33','34','34','35','36','36'],
    ['31','32','32','32','32','32','33','33','33','34','35','36']]
D2=[
 ['参数表单与字段校验','无效数值与保存接口','操作历史接口设计','参数保存联接核对','重建失败提示与回滚核验','选边界面接入','历史与模型树状态绑定','半径距离前置校验','边处理异常提示','模型状态与QA联调','余量','阶段全链路回归'],
 ['工程参数协议样例','历史记录数据设计','版本状态字段实现','参数映射样例核对','参数重建测试输入','开孔布尔正反例准备','边处理极值样例准备','历史序列测试输入','布尔输入条件核对','边处理错误分支核对','编辑序列回归','阶段全链路回归'],
 ['当前模型状态契约','历史ID与模型关联','重建QA钩子接入','双实体选择接口','边映射报告接口准备','拉伸QA适配','布尔QA断言准备','历史状态刷新适配','边处理报告联接准备','开孔布尔QA联调','边处理结果QA联调','阶段全链路回归'],
 ['撤销与模型树测试设计','圆角倒角样例设计','轮廓拉伸测试设计','重建尺寸核对','重建接口契约联调','联调执行步骤准备','轮廓拉伸联调','布尔量测样例准备','布尔错误输入预检','贯穿性与融合结果核对','边处理组合回归','阶段全链路回归'],
 ['参数到内核重建','孔位与尺寸保持','内核异常回滚','轮廓构建接口','拉伸有效实体接入','贯穿孔内核操作','布尔内核接入','圆角内核与选边','倒角内核与失败保护','撤销恢复内核接入','编辑内核结果复核','阶段全链路回归'],
 ['参数模型样例准备','矩形轮廓UI接入','闭合与零尺寸校验','拉伸UI接入','轮廓实体衔接','轮廓拉伸结果核对','开孔布尔界面准备','开孔布尔调用接入','孔与实体结果核对','圆角倒角结果核对','撤销模型树组合核对','阶段全链路回归']]
S3=[['41',None,None,None,'43',None,'43',None,None,'44',None,'45'],
 ['42','41',None,None,None,'43',None,'43',None,'44',None,'45'],
 ['41','41','41','42','42','43','43','43','44','44','44','45'],
 ['41','41','41','41','42','42','42','43','43','43','44','45'],
 ['41',None,None,None,None,'43',None,'43',None,None,'44','45'],
 ['41','41','42','42','42',None,'43',None,'43',None,'44','45']]
NAMES={'31':'参数编辑与模型重建','32':'矩形轮廓与实体拉伸','33':'贯穿孔与布尔并集','34':'边圆角与边倒角','35':'撤销与模型树','36':'编辑联调与验收','41':'STEP导出与兼容验证','42':'STL与GLB导出','43':'全流程回归与缺陷关闭','44':'发布与固定剧本演示','45':'最终验收与资料归档支持'}
# Gates derive from identified producer tasks; no manual finish times feed validation.
GATES2={
 '参数契约':[(0,1),(1,0),(2,0)],'重建内核':[(4,2),(2,2)],
 '重建样例':[(0,4),(3,3)],'拉伸内核':[(4,4)],'拉伸样例':[(5,5)],
 '选择接口':[(2,3)],'布尔内核':[(4,6)],'开孔布尔样例':[(5,8)],
 '边处理内核':[(4,8)],'边处理样例':[(5,9),(1,9)],
 '历史契约':[(0,2),(1,2),(2,1)],'撤销内核':[(4,9),(2,7)],'撤销样例':[(5,10)],
 'M2阶段关口':[(i,11) for i in range(6)]}
GATES3={'STEP导出接口':[(2,2),(5,1),(3,2)],'网格导出接口':[(2,4),(5,4)],'STEP样例':[(3,3),(2,2)],'三格式样例':[(3,6),(2,4)],'回归材料':[(3,9),(2,7),(5,8)],'固定演示':[(i,10) for i in [2,3,4,5]],'M3归档':[(i,11) for i in range(6)]}
def task_id(it,i,t):return f'I{it}-{i+1:02}-{t+1:02}'
def plan():
 tasks=[]
 for it,sched in [(2,S2),(3,S3)]:
  for i,arr in enumerate(sched):
   for t,code in enumerate(arr):
    if code is None:continue
    dep=[]
    if it==2:
     if code=='31' and t>=3:dep=['参数契约','重建内核']
     if code=='32' and i==5 and t==5:dep=['拉伸内核']
     if code=='32' and i==4:dep=['重建内核']
     if code=='33' and i==4:dep=['拉伸内核']
     if code=='34' and i==4:dep=['布尔内核']
     if code=='35' and i==4:dep=['历史契约','边处理内核']
     if code=='33' and i==5 and t>=7:dep=['选择接口','布尔内核']
     if code=='33' and i==3 and t==9:dep=['开孔布尔样例']
     if code=='34' and t>=9:dep=['边处理内核']
     if code=='35' and t>=6:dep=['历史契约']
     if code=='35' and i==5 and t==10:dep+=['撤销内核']
     if code=='36':
      dep=['重建内核'] if t<=5 else (['拉伸样例'] if t==6 else (['开孔布尔样例'] if t==9 else ['边处理样例']))
      if t==11:dep=['重建样例','拉伸样例','开孔布尔样例','边处理样例','撤销样例']
     name=D2[i][t]
    else:
     name={'41':'STEP接口/样例准备','42':'网格与GLB接口/样例准备','43':'回归样例与预期准备','44':'启动及恢复说明准备','45':'版本资料归档'}[code]
     if code=='41':
      name={0:'STEP调用与异常接口设计',1:'STEP参数及模型传递',2:'STEP输出与错误保护实现',3:'STEP打开规则准备',4:'STEP内核导出接口复核',5:'STEP导出样例与兼容接入'}[i]
      if t==1:name={0:'STEP调用与异常接口设计',1:'STEP参数及模型传递',2:'STEP输出与错误保护实现',3:'STEP兼容性用例准备',4:'STEP内核导出接口复核',5:'STEP导出样例接入实现'}[i]
      if i==3 and t==2:name='STEP加载与尺寸核对脚本'
     if code=='42':name={0:'网格导出接口协作',1:'STL与GLB导出协议',2:'STL/GLB输出及异常实现',3:'三格式几何核验准备',4:'网格内核协作',5:'网格导出调用与样例接入'}[i]
     if code=='41' and t==3:name='STEP外部打开与主体尺寸核对';dep=['STEP导出接口']
     if code=='42' and t==6:name='三格式文件及几何结果核对';dep=['STEP样例','网格导出接口']
     if code=='43' and t>=5:
      name='已就绪格式及编辑链路回归';dep=['STEP样例']
      if t>=7:dep=['三格式样例']
     if code=='44':name='发布与恢复步骤演练' if t<10 else '固定剧本与接收材料核对';dep=['三格式样例'] if t<10 else ['回归材料']
     if code=='45':dep=['固定演示']
    tasks.append({'id':task_id(it,i,t),'iteration':it,'person':PEOPLE[i],'person_index':i,'day':(3 if it==2 else 5)+t//6,'slot':t%6,'start':t,'finish':t+1,'hours':1,'wbs':f'1.{code[0]}.{code[1]}','package':NAMES[code],'task':name,'deliverable':name+'记录/接口或样例（计划）','dependencies':dep,'state':'计划'})
 gates=[]
 for it,gs in [(2,GATES2),(3,GATES3)]:
  byid={x['id']:x for x in tasks if x['iteration']==it}
  for name,producers in gs.items():
   ids=[task_id(it,i,t) for i,t in producers];assert all(x in byid for x in ids)
   finish=max(byid[x]['finish'] for x in ids)
   gates.append({'iteration':it,'gate':name,'finish':finish,'producer_ids':ids})
  gatetime={g['gate']:g['finish'] for g in gates if g['iteration']==it}
  for x in byid.values():
   for dep in x['dependencies']:assert gatetime[dep]<=x['start'],(x,dep,gatetime[dep])
 for it,total in [(2,71),(3,49)]:
  assert sum(x['hours'] for x in tasks if x['iteration']==it)==total
  for person in PEOPLE:
   assert sum(x['hours'] for x in tasks if x['iteration']==it and x['person']==person)<=12
   for day in [3,4,5,6]:assert sum(x['hours'] for x in tasks if x['day']==day and x['person']==person)<=6
 assert len({(x['day'],x['slot'],x['person']) for x in tasks})==len(tasks)
 return tasks,gates

def boxdraw(title,nodes,parents,path):
 # 1600px canvas, explicit root/branch/leaf connectors; save editable SVG beside PNG.
 H=max(y+h for x,y,w,h,text in nodes.values())+35;W=max(x+w for x,y,w,h,text in nodes.values())+20
 im=Image.new('RGB',(W,H),'white');draw=ImageDraw.Draw(im);fontpath='/System/Library/Fonts/Supplemental/Songti.ttc';font=ImageFont.truetype(fontpath,30)
 svg=[f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}"><rect width="100%" height="100%" fill="white"/>']
 for child,parent in parents.items():
  a=nodes[parent];b=nodes[child];x1=a[0]+a[2]/2;y1=a[1]+a[3];x2=b[0]+b[2]/2;y2=b[1]
  points=[(x1,y1),(x1,(y1+y2)/2),(x2,(y1+y2)/2),(x2,y2)]
  if parent!='root':
   bus=950;mid=b[1]+b[3]/2;side=b[0]+b[2] if b[0]<bus else b[0]
   points=[(x1,y1),(x1,y1+25),(bus,y1+25),(bus,mid),(side,mid)]
  draw.line(points,fill='#444444',width=2);svg.append('<polyline points="'+' '.join(f'{x},{y}' for x,y in points)+'" fill="none" stroke="#444" stroke-width="2"/>')
 for k,(x,y,w,h,text) in nodes.items():
  draw.rectangle((x,y,x+w,y+h),fill='#f6f6f6',outline='#666666',width=2)
  lines=text.split('\n');nodefont=ImageFont.truetype(fontpath,26 if w<500 else 30)
  for j,line in enumerate(lines):draw.text((x+14,y+12+36*j),line,font=nodefont,fill='black')
  svg.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#f6f6f6" stroke="#666"/>')
  for j,line in enumerate(lines):svg.append(f'<text x="{x+14}" y="{y+38+36*j}" font-family="Songti SC,serif" font-size="{26 if w<500 else 30}">{html.escape(line)}</text>')
 im.save(path);Path(path).with_suffix('.svg').write_text(''.join(svg)+'</svg>')

def diagrams(out,blocks,tasks):
 asset=out/'assets';wbs={}
 # Extract names from dictionary headings, no independent name list drift.
 for b in blocks:
  if b['type']=='p':
   m=re.match(r'^(1\.[1-4](?:\.[1-6])?) (.+)',b['text'])
   if m and b['style'].startswith('Heading'):wbs[m[1]]=m[2]
 parentnames={'1.1':'项目管理与过程资料','1.2':'生成、检索与QA成果','1.3':'参数编辑与基础CAD操作成果','1.4':'导出、演示与项目交付成果'}
 for fig,prefix in zip(range(2,6),parentnames):
  nodes={'root':(420,20,860,65,'0.0 SimpleCADforDSH项目交付')};parents={}
  if prefix=='1.1':
   for i,(id,name) in enumerate(parentnames.items()):
    nodes[id]=(20+i*420,145,400,105,id+'\n'+name);parents[id]='root'
   y=320
  else:
   nodes[prefix]=(420,145,860,70,prefix+' '+parentnames[prefix]);parents[prefix]='root';y=290
  for n,(id,name) in enumerate((x for x in wbs.items() if x[0].startswith(prefix+'.'))):
   nodes[id]=(180+(n%2)*800,y+(n//2)*140,740,75,id+' '+name);parents[id]=prefix
  boxdraw(prefix,nodes,parents,asset/f'figure-{fig}.png')
 # A whole drawio diagram with root and all package nodes/edges; values match report.
 mx=ET.Element('mxfile',host='app.diagrams.net');di=ET.SubElement(mx,'diagram',name='WBS');model=ET.SubElement(di,'mxGraphModel');root=ET.SubElement(model,'root');ET.SubElement(root,'mxCell',id='0');ET.SubElement(root,'mxCell',id='1',parent='0')
 ids={'0.0':'SimpleCADforDSH'}|parentnames|{k:v for k,v in wbs.items() if k.count('.')==2}
 for idx,(id,name) in enumerate(ids.items()):
  cell=ET.SubElement(root,'mxCell',id=id,value=id+' '+name,vertex='1',parent='1',style='rounded=0;whiteSpace=wrap;html=1;');
  if id=='0.0':x,y=900,0
  elif id.count('.')==1:x,y=(int(id[-1])-1)*600,180
  else:x,y=(int(id.split('.')[1])-1)*600,350+(int(id.split('.')[2])-1)*120
  ET.SubElement(cell,'mxGeometry',x=str(x),y=str(y),width='500',height='90',attrib={'as':'geometry'})
  if id!='0.0':
   parent='0.0' if id.count('.')==1 else id.rsplit('.',1)[0]
   edge=ET.SubElement(root,'mxCell',id='edge-'+id,edge='1',parent='1',source=parent,target=id,style='edgeStyle=orthogonalEdgeStyle;');ET.SubElement(edge,'mxGeometry',relative='1',attrib={'as':'geometry'})
 ET.ElementTree(mx).write(asset/'统一WBS结构树.drawio',encoding='utf-8',xml_declaration=True)
 # resource-level timeline, one hour per slot for D3-D6; original D1-D2 baseline retained.
 from PIL import ImageFont
 font=ImageFont.truetype('/System/Library/Fonts/Supplemental/Songti.ttc',27)
 im=Image.new('RGB',(1900,1010),'white');d=ImageDraw.Draw(im)
 d.text((30,20),'六日开发目标与资源交接安排',font=ImageFont.truetype('/System/Library/Fonts/Supplemental/Songti.ttc',42),fill='black')
 d.text((30,80),'D1—D2：沿用66小时原任务基线；D3—D6：每格1有效人时，空白为余量',font=font,fill='black')
 colors={'31':'#dfe9ff','32':'#e4f0e3','33':'#fff1d8','34':'#ffe3df','35':'#eee4ff','36':'#d9edf3','41':'#dfe9ff','42':'#e4f0e3','43':'#fff1d8','44':'#eee4ff','45':'#d9edf3'}
 for it,ys,sch in [(2,160,S2),(3,570,S3)]:
  startday=3 if it==2 else 5
  for day in range(2):
   d.text((190+day*810,ys),f'D{startday+day} 上午（3h）              下午（3h）',font=font,fill='black')
  for i,name in enumerate(PEOPLE):
   y=ys+50+i*48;d.text((30,y+8),name,font=font,fill='black')
   for t,code in enumerate(sch[i]):
    x=180+t*137;d.rectangle((x,y,x+130,y+44),fill=colors.get(code,'white'),outline='#888888');d.text((x+10,y+8),('1.'+code[0]+'.'+code[1]) if code else '余量',font=font,fill='black')
 d.text((30,940),'交接关口：D3重建/拉伸 → D4开孔布尔/边处理/撤销 → D5格式样例 → D6回归/演示/归档',font=font,fill='black')
 path=asset/'resource-gantt.png';im.save(path)
 for b in blocks:
  if b['type']=='image' and b['source'].startswith('刘子轩'):b['path']=str(path.relative_to(out))


def adopt_report_language(blocks):
 """用户认可当前估算、安排与验收口径；保留提交前内部复核及真实执行状态。"""
 prefixes={
 '项目预算采用MERGE-BUDGET-02建议方案。':'项目预算采用MERGE-BUDGET-02计划。首迭代66开发人时沿用既有任务表，后两迭代按工作包分解为71和49开发人时，管理人时另列。新增人时、调用用量与计费参数作为本报告的计划估算，提交前由团队内部复核；执行中按实际资源和计费情况更新。开发人员会议及协调已在每迭代每人预留4小时中扣除，不再次计入净开发任务。',
 '现阶段可将“继续评估现有插件扩展B”':'本项目采用现有插件扩展B路线及XZJ-ECO-01经济假设。该路线复用既有插件基线，但N=200时仅比人工方案低200元，工时上升可能使优势消失。因此按工作包跟踪投入，并结合实际用量更新经济分析；路线与资源取舍在提交前由团队内部复核。',
 '预算控制拟由主责人':'预算控制由主责人每次交接报告实际工时、剩余工时和调用消耗，SM按WBS汇总，PO比较计划基线与完成预测。若费用数量、活动剩余量或资源约束变化，应说明输入变化及对验收的影响，保留变更前后版本。本报告采用的估算在提交前内部复核；实际现金支出按资源报价、账单和用量核算，采购及付款按项目变更与费用管理程序办理。',
 '本附录为六天目标下的建议执行安排':'本附录列出六天目标下本报告采用的执行计划，提交前由团队内部复核。D3—D6每天6有效小时，分上午3小时与下午3小时；小时序号表示净任务时段，会议和协调另按原容量预留。当前安排通过逐人时段冲突、每日容量和列明产物依赖检查；开工前检查技能、环境与样例，执行中按关口跟踪交接。',
 '完整WBS字典以交付物为组织依据':'完整WBS字典以交付物为组织依据，保留工作内容、质量目标、责任、依赖及交接窗口。首迭代责任依据既有任务表，后两迭代采用本报告计划；具体人员时段见附录E及计划预算工作簿，提交前由团队内部复核。百分比为验收目标，未据此认定实际通过。原资料中的历史签署信息不作为本轮计划采用的依据。',
 '[17] 项目计划预算MERGE-BUDGET-02':'[17] 项目计划预算MERGE-BUDGET-02，2026年11月4日。配套SimpleCADforDSH计划与预算.xlsx及source/计划数据.json；按当前报告采用的估算与安排编制，提交前团队内部复核。',
 }
 replacements={
 '需由PO确认后执行':'按本报告计划执行',
 '归类仍待PO确认':'归类按本报告计划采用',
 '本计划经PO与主责复核后生效':'本报告采用上述计划，提交前由团队内部复核',
 '开工前由PO确认该清单及验收规则':'本报告采用该清单及验收规则，提交前内部复核',
 '拟采用的异常验收规则':'本报告采用的异常验收规则',
 '上述预期基于讨论四及质量计划提出，开工前由PO确认':'上述预期依据讨论四及质量计划补充，作为本报告验收规则，提交前内部复核',
 '但该口径仍待PO确认':'本报告采用该口径并在提交前内部复核',
 'PO正式确认状态另记':'采用补充预期，提交前内部复核',
 '后两迭代建议安排':'后两迭代计划安排',
 '图6和附录E展示建议时段':'图6和附录E展示计划时段',
 '日历起止日期尚未确认':'实际日历按开工安排确定',
 '开发日历未确定时使用第1—6天':'开发日程按相对第1—6天编制',
 '开发期限、签署与批准状态据实维护':'开发日历与历史签署记录据实维护，管理计划按本报告采用口径组织',
 '新增API、算力等费用上限、承担方式、批准人和实际开发起止日期待确认，不代填签署':'新增API、算力按实际计费参数测算，PO统筹费用与资源；开发期按D1—D6相对日程组织，实际开工时排定日历',
 '批准状态另行记录':'按本报告角色安排执行',
 '候选分配':'人员分配',
 '全项目暂定预算':'全项目计划预算',
 '迭代二三暂定自下而上工时':'迭代二三自下而上计划工时',
 '后两迭代为暂定':'后两迭代为计划估算',
 '计划暂定输入，不追溯为原已批准预算':'当前计划估算输入，历史估算记录另行保留',
 '管理人时暂按':'管理人时按',
 '管理h 暂定':'管理h 计划',
 '调用次数 暂定':'调用次数 计划',
 '项目按暂定输入测算费用':'项目按计划输入测算费用',
 'g为待确认小时单价':'g为实际小时单价参数',
 '不具备资源报价或许可依据的支出仍待确认，不能由这个情景取得采购授权':'资源报价和许可费用按实际使用条件核定，该情景不构成真实报价或支出上限',
 '同版建议安排':'同版计划安排',
 '本报告建议安排':'本报告计划安排',
 '建议投入':'计划投入',
 '既有WBS与建议任务拆分':'既有WBS与计划任务拆分',
 '迭代二三暂定拆分':'迭代二三计划拆分',
 '日历待确认':'日历按实际开工安排',
 '范围变更和待确认事项':'范围变更和复核事项',
 '最终迁移清单由PO确认':'迁移清单按US13—US22执行，提交前内部复核',
 'PO确认G4清单':'按US13—US22清单执行',
 '费用上限待批准':'现金参数化估算，实际费用按账单和用量核定',
 '均须PO确认':'按触发升级机制处理',
 'PO确认取舍':'按触发升级机制处理',
 '建议按9开发与1管理分别计入':'按9开发与1管理分别计入',
 '正式计划仍由学生审阅决定':'本报告采用的计划提交前由学生内部复核',
 '本报告中的新增人时及现金用量假设须由许祖嘉和相关主责审阅':'本报告采用的新增人时及现金用量假设在提交前由许祖嘉和相关主责内部复核',
 'PO与主责确认估算，不视为已批准':'PO与主责在提交前内部复核计划估算',
 '本次统稿未代替学生宣布新预算获批、AT通过或严重问题已解决':'本报告采用资源与预算计划，提交前由团队复核；AT未执行、严重问题未解决的状态如实保留',
 }
 formal={'本报告采用该清单及验收规则，提交前内部复核，': '', '，按本报告计划执行': '', '管理计划按本报告采用口径组织，': '', '采用补充预期，提交前内部复核': '异常输入预期为提示无效并保护原模型', '归类按本报告计划采用；': '', '迁移清单按US13—US22执行，提交前内部复核': '迁移清单按US13—US22执行', '；本报告采用上述计划，提交前由团队内部复核': '', '本报告采用的异常验收规则：': '异常验收规则：', '作为本报告验收规则，提交前内部复核，': '', '，本报告采用该口径并在提交前内部复核': '', '；路线与资源取舍在提交前由团队内部复核': '', '新增人时、调用用量与计费参数作为本报告的计划估算，提交前由团队内部复核；执行中按实际资源和计费情况更新。': '人时和调用用量为计划估算，执行中按实际资源和计费情况更新。', '本报告采用的估算在提交前内部复核；': '', '；本报告采用的计划提交前由学生内部复核': '', 'PO与主责在提交前内部复核计划估算': 'PO与主责维护估算及变更记录', '本报告采用资源与预算计划，提交前由团队复核；': '', '本报告采用的新增人时及现金用量假设在提交前由许祖嘉和相关主责内部复核': '新增人时及现金用量假设以工作包分解和经济情景为依据', '；按当前报告采用的估算与安排编制，提交前团队内部复核': '', '后两迭代采用本报告计划': '后两迭代按工作包分解和资源容量安排', '，提交前由团队内部复核': '', '原资料中的历史签署信息不作为本轮计划采用的依据。': '', '本附录列出六天目标下本报告采用的执行计划': '本附录列出六天开发期的任务安排与交接关口', '八名成员共同审阅后形成最终稿': '学生对最终结论负责', '人工最终复核尚需八人共同完善': '学生负责项目角色与经济分析决策', '最终报告拟由八名成员共同复核范围、预算和验收依据，保留各项实际执行记录。': '项目范围、预算和验收结论以相应计划及实际执行记录为依据。'}
 for b in blocks:
  def f(v):
   if not isinstance(v,str):return v
   for prefix,new in prefixes.items():
    if v.startswith(prefix):v=new;break
   for old,new in replacements.items():v=v.replace(old,new)
   for old,new in formal.items():v=v.replace(old,new)
   return v
  if b['type']=='p':b['text']=f(b['text'])
  elif b['type']=='table':b['head']=[f(v) for v in b['head']];b['rows']=[[f(v) for v in row] for row in b['rows']]


def apply(blocks,out,budget,people):
 tasks,gates=plan();maps={x[0]:x for x in budget};changed=[]
 # Current report adopts the plan per user instruction; historical signatures and actual test results are unchanged.
 replace={
 '目标沿用G1—G5：G1覆盖文本及单、多视角生成；G2覆盖文本与图像检索复用；G3显示几何与尺寸检查；G4提供清单内有限编辑；G5实现可恢复的固定剧本。US01—US25保留稳定编号，讨论二AT是故事验收依据，阶段目标另行核对。早期“最多6项编辑能力”是讨论一设想，须与后续US13—US22协调，由PO确认，不能静默删减故事。[1][2]':'目标沿用G1—G5：G1覆盖文本及单、多视角生成；G2覆盖文本与图像检索复用；G3显示几何与尺寸检查；G4提供清单内有限编辑；G5实现可恢复的固定剧本。US01—US25为故事及验收范围，G4按US13—US22对应清单逐项核验，阶段目标与故事AT分别记录。开工前由PO确认该清单及验收规则，范围变更须同步计划。[1][2]',
 }
 for b in blocks:
  if b['type']!='p':continue
  old=b['text'];s=replace.get(old,old)
  if s.startswith('迭代一对应Version'):s='迭代一对应Version 0.5，形成US01—US12生成、检索及QA基础增量；迭代二对应Version 0.8，形成US13—US22参数编辑与基础操作；迭代三对应Version 1.0，完成US23—US25导出、总回归、演示和归档。版本名称表示计划交付关口，原故事估算分别为58、40、11点，重估记录与资源排期分别管理。'
  if s.startswith('最小工作包用于分派'):s='最小工作包用于分派和核验，首迭代使用I1原任务号记录人时，工作包汇总不再次叠加到子任务工时。1.3为参数编辑与基础CAD操作，1.3.3为贯穿孔与布尔并集。各工作包的新增任务分配、时间窗和交接条件统一见同版计划预算工作簿及附录E；本计划经PO与主责复核后生效。'
  if s.startswith('本节补做预算情景'):s='项目预算采用MERGE-BUDGET-02建议方案。首迭代66开发人时沿用既有任务表，后两迭代按工作包分解为71和49开发人时，管理人时另列。新增人时、调用用量与计费参数是计划估算，尚未作为批准预算；开工前由PO与各主责核实任务、资源可得性和费用承担方式后确认。开发人员会议及协调已在每迭代每人预留4小时中扣除，不再次计入净开发任务。'
  if s.startswith('迭代二总计71小时'):s='第二迭代总计71人时，按专业工作与协作任务重新分配为常诚11小时，其余五名开发成员各12小时。侯永森、庞佳泽负责CAD内核及关键几何结果复核，其他成员承担参数校验、模型状态接口、测试输入和集成记录。D3上午先形成参数及历史契约，重建内核就绪后开展联调；D3结束前交付轮廓拉伸样例，D4依次交接开孔布尔、边处理和撤销样例，最后一小时六人共同完成阶段回归。附录E列明逐人时段及依赖，不用总余量代替个人容量检查。'
  if s.startswith('第三迭代暂定49开发人时'):s='第三迭代总计49人时，候选分配为常诚5、刘子轩6、余智皓12、严翊航12、侯永森5、庞佳泽9小时。庞佳泽协助一小时格式核验，严翊航继续主导格式和尺寸复核。D5推进STEP及STL/GLB接口、样例和外部读取，D6完成三格式联调、全链路回归、固定演示及归档。资源计划按每人每天最多6有效小时展开，人员不重复占用；技能可用性和实际执行效率仍须开工前复核。首迭代9小时条件开发增量若阻断G1—G3，应当期升级，不能直接用后两迭代余量抵扣。'
  if s.startswith('工时估计考虑复用已有内核'):s='工作包估算考虑复用已有内核与导出链路，保留实现、自测、联调和资料四类工时。参数和历史接口设计可以先于几何结果，异常样例准备可以与内核接入并行；只有对应内核及模型样例交接后才安排结果核验。各项任务均为计划，模型尚未就绪时的用例准备不等同于AT已经执行。附录E按产物关口列依赖，区分准备、接入和结果核验，形成可供团队复核的六日安排。'
  if s.startswith('本稿已完成') or s.startswith('本轮已完成79条AT'):s='需求追踪覆盖79条AT，关联WBS、任务、责任及时窗；AT17-02提出无效参数提示和原模型保护的可判定预期，AT20-03补充过大倒角距离的异常步骤。成本与风险均采用12已含、9条件开发和1管理的计入口径。协作记录C01—C07保留范围、风险、需求映射和工具检查等事项的来源与处理，作为计划编制过程依据。[14]'
  if s.startswith('图6由工作簿'):s='图6按同版任务计划生成。D1—D2沿用既有首迭代基线，D3—D6以每格一有效人时展示人员、工作包和余量；附录E及工作簿保存具体任务、交接产物和前置关口。重建、拉伸、特征及撤销分别交接，最终阶段回归在全部样例就绪后进行。'
  if s.startswith('图6 ') or s.startswith('图6  '):s='图6 六日开发目标下的D3—D6资源与工作包安排'
  if s.startswith('图2 WBS'):s='图2 WBS项目根项、四类交付及管理分支'
  if s.startswith('图3 WBS'):s='图3 WBS生成、检索与QA分支'
  if s.startswith('图4 WBS'):s='图4 WBS参数编辑与基础操作分支'
  if s.startswith('图5 WBS'):s='图5 WBS导出、演示与交付分支'
  if s.startswith('以下按原WBS归并'):s='完整WBS字典以交付物为组织依据，保留工作内容、质量目标、责任、依赖及交接窗口。首迭代责任依据既有任务表，后两迭代采用本报告建议安排；具体人员时段见附录E及计划预算工作簿。百分比为验收目标，未据此认定实际通过。原资料中的审批栏未作为本次修订的批准依据。'
  if s.startswith('经济比较要求三路线'):s=s.replace('完整公式和敏感性见原工作簿及本稿4.4。','公式和敏感性见同版计划预算工作簿“经济比较”表及本节。')
  if s.startswith('工作簿把WBS窗口'):s='任务安排使用D1—D6相对开发日。D1—D2沿用既有66人时基线；同版工作簿以一小时任务列出D3—D6的人员占用与产物依赖，图6和附录E展示建议时段。D4末完成编辑阶段回归，D6依次完成全链路回归、固定演示和归档。日历起止日期尚未确认，课程报告提交日期不作为开发日期。'
  if s.startswith(('[17] 项目暂定预算','[17] 本次暂定预算')):s='[17] 项目计划预算MERGE-BUDGET-02，2026年11月4日。配套SimpleCADforDSH计划与预算.xlsx及source/计划数据.json；新增估算与安排为建议，尚未批准。'
  if s.startswith('本报告使用OpenAI Codex'):s='本报告使用OpenAI Codex辅助资料核对、需求追踪、经济计算、风险归类、图表制作和版式整理。学生对项目选题、目标、角色、范围、预算确认、验收和最终结论负责。辅助计算检查了故事点、人时和风险计入方式，提出资源重排及预算情景；正式计划仍由学生审阅决定。AT未执行、CodeArts门禁未通过等状态按已有证据保留。'
  s=s.replace('MERGE-BUDGET-01','MERGE-BUDGET-02').replace('本次暂定预算','项目暂定预算').replace('本次统稿','报告整合').replace('本次保持','本期保持').replace('本次增加项目级暂定估算','项目按暂定输入测算费用').replace('本次暂定输入','计划暂定输入').replace('原资料中的审批栏未作为本次修订的批准依据。','计划生效以独立批准记录为准，原资料审批栏不作为当前建议的批准依据。').replace('本次据现有WBS补做后两迭代暂定工时，见4.4.3；个人超载仍需PO确认调整。','后两迭代建议安排经逐人每日容量及交接关口检查，见4.4.3与附录E；执行前须复核技能和资源条件。')
  if not b.get('body') and s.startswith('原表状态仅按来源保留'):s='风险业务状态按2026年11月1日读取的来源保留；实际触发及关闭以执行记录为准。'
  if old!=s:changed.append({'before':old,'after':s});b['text']=s
 # Replace original special WBS timings and staffing with the same plan used in task allocation.
 current=None
 for b in blocks:
  if b['type']=='p':
   m=re.match(r'^(0\.0|1\.[1-4](?:\.\d+)?) ',b['text'])
   if m and b['style'].startswith('Heading'):current=m[1]
  if b['type']=='table':
   for r in b['rows']:
    for i,v in enumerate(r):
     if isinstance(v,str):r[i]=v.replace('作为章程输入；本次许祖嘉确认兼任项目经理与PO，原记录保持原样','作为章程输入；许祖嘉角色补充说明其兼任项目经理与PO，批准状态另行记录').replace('几何判定准确率90%；','AT11—AT12的正反例逐项满足预期；').replace('本次分开9开发与1管理，补暂定预算并暴露超载','建议按9开发与1管理分别计入，联动预算及容量升级机制')
   if b['head'][0]=='成员' and '迭代1h' in b['head']:
    a1=[12,12,10,10,11,11];a2=[11,12,12,12,12,12];a3=[5,6,12,12,5,9]
    b['rows']=[[n,a1[i],a2[i],a3[i],12] for i,n in enumerate(people)]+[['合计',66,71,49,72]]
   if b['head']==['字段','内容'] and current in maps:
    row=maps[current];alloc=row[4]
    for r in b['rows']:
     if r[0]=='责任主体':r[1]=r[1]+'；建议投入：'+'、'.join(f'{n}{h}h' for n,h in zip(people,alloc) if h)
     if r[0]=='时间要求':
      ts=[x for x in tasks if x['wbs']==current];lo=min(x['day'] for x in ts);hi=max(x['day'] for x in ts)
      r[1]=f'D{lo}—D{hi}；准备与实现分步交接，具体依赖与验收时段见附录E。'
     if r[0]=='估算来源':r[1]='既有WBS与建议任务拆分；MERGE-BUDGET-02及同版工作簿。任务重排保持各工作包总工时不变，执行前由PO和主责复核。'
 # Explicit contract for AT exceptions once, preserving full original rows in appendix.
 for b in blocks:
  if b['type']=='p' and b['text'].startswith('AT17-02原文仅写'):
   b['text']='拟采用的异常验收规则：AT17-02在正常支撑板上输入无法形成有效孔的参数，预期显示参数无效且原模型可继续查看，尺寸、几何和模型标识不发生错误改变；AT20-03在100×80×20 mm支撑板目标边输入100 mm倒角距离，预期提示几何无法成立并保护原模型。上述预期基于讨论四及质量计划提出，开工前由PO确认，原始AT与补充步骤分栏保留。可编辑参数从工程描述和脚本PARAMS读取，不承诺由任意STEP恢复完整设计历史。'
 # Existing risk/cost tables are coherent; ensure reference to integrated current workbook.
 for b in blocks:
  if b['type']=='p':b['text']=b['text'].replace('原工作簿','同版计划预算工作簿').replace('xlsx表格里的“活动计划”','同版工作簿“任务计划”')
 # Add resource details as deliverable appendix, not issue register.
 def P(text,style='Normal'):blocks.append({'type':'p','text':text,'style':style,'source':'MERGE-BUDGET-02资源计划','body':False})
 def T(head,rows,widths):blocks.append({'type':'table','head':head,'rows':rows,'widths':widths,'source':'MERGE-BUDGET-02资源计划','body':False})
 P('附录E 资源安排与交接关口','Heading 1')
 P('本附录为六天目标下的建议执行安排，未作为已批准或已完成事实。D3—D6每天6有效小时，分上午3小时与下午3小时；小时序号表示净任务时段，会议和协调另按原容量预留。建议方案通过逐人时段冲突、每日容量和列明产物依赖检查，开工前仍需主责确认技能、环境与样例。')
 for day in range(3,7):
  P(f'D{day}逐人工作安排','Heading 2')
  rows=[]
  for person in people:
   tt=[x for x in tasks if x['day']==day and x['person']==person]
   desc=[]
   for half in [0,1]:
    items=[]
    for slot in range(half*3,half*3+3):
     x=next((x for x in tt if x['slot']==slot),None);items.append(f'{slot+1}h '+(x['wbs']+' '+x['task'] if x else '余量'))
    desc.append('\n'.join(items))
   rows.append([person]+desc+[str(len(tt))])
  T(['成员','上午3有效小时','下午3有效小时','合计h'],rows,[20,57,57,12])
 P('产物关口与阶段放行','Heading 2')
 rows=[]
 for g in gates:
  t=g['finish'];day=(3 if g['iteration']==2 else 5)+(t-1)//6;hour=(t-1)%6+1
  rows.append([f'I{g["iteration"]} '+g['gate'],f'D{day}第{hour}有效小时结束',','.join(g['producer_ids'])])
 T(['关口','计划就绪时点','交接任务ID'],rows,[37,37,72])
 P('每个关口须交付接口/代码版本、固定输入及预期、当前模型和报告关联。若到时未就绪，则依赖该产物的结果核验不得开始；准备工作可以并行，但不能登记AT通过。SM记录影响，PO按剩余容量和课程目标决定改派、重新排序或变更。首迭代条件开发增量全部发生时75/72人时，需在当期处理；后续余量不构成自动放行条件。')
 adopt_report_language(blocks)
 diagrams(out,blocks,tasks)
 # Export master dataset for workbook and acceptance check; no Excel authoring in Python.
 daily=[]
 for day in range(3,7):
  for n in people:daily.append({'day':day,'person':n,'hours':sum(x['hours'] for x in tasks if x['day']==day and x['person']==n)})
 data={'version':'MERGE-BUDGET-02','people':people,'tasks':tasks,'gates':gates,'daily':daily,'packages':[{'wbs':r[0],'name':r[1],'iteration':r[2],'effort_parts':r[3],'allocations':r[4],'hours':sum(r[3])} for r in budget],'baseline1':[12,12,10,10,11,11],'conditions':'开工前检查技能与资源；AT未执行，CodeArts门禁未通过。'}
 data['risks']=list(csv.DictReader((out.parent/'余智皓_课程报告分稿/图表及源文件/十大风险登记册.csv').open(encoding='utf-8-sig')))
 (out/'source/计划数据.json').write_text(json.dumps(data,ensure_ascii=False,indent=2))
 (out/'source/正文修订对照.json').write_text(json.dumps(changed,ensure_ascii=False,indent=2))
 with (out/'source/逐人任务安排.csv').open('w') as f:
  w=csv.DictWriter(f,fieldnames=list(tasks[0]));w.writeheader()
  for x in tasks:w.writerow({**x,'dependencies':'、'.join(x['dependencies'])})
 return data
