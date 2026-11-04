import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
const src=path.dirname(fileURLToPath(import.meta.url));const out=path.dirname(src);
const rt='/tmp/simplecad-report-workbook-runtime';await fs.mkdir(rt,{recursive:true});
try{await fs.symlink('/Users/pano/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules',rt+'/node_modules');}catch(e){if(e.code!=='EEXIST')throw e;}
const req=createRequire(rt+'/entry.mjs');const {Workbook,SpreadsheetFile}=await import(req.resolve('@oai/artifact-tool'));
const d=JSON.parse(await fs.readFile(src+'/计划数据.json','utf8'));const wb=Workbook.create();
const names=['概览','任务计划','容量','预算','风险','经济比较'];const sh={};for(const name of names)sh[name]=wb.worksheets.add(name);
function set(s,address,value){sh[s].getRange(address).values=[[value]];}
function fx(s,address,formula){sh[s].getRange(address).formulas=[[formula]];}
function head(s,row,labels){sh[s].getRangeByIndexes(row-1,1,1,labels.length).values=[labels];const r=sh[s].getRangeByIndexes(row-1,1,1,labels.length);r.format.fill='#24445C';r.format.font.color='#FFFFFF';r.format.font.bold=true;r.format.rowHeight=30;r.format.wrapText=true;r.format.horizontalAlignment='center';}
for(const n of names){const s=sh[n];s.showGridLines=false;const r=s.getRange('B2:Q170');r.format.font.name='Songti SC';r.format.font.size=11;r.format.rowHeight=24;r.format.columnWidth=17;r.format.verticalAlignment='center';s.getRange('B2').values=[['SimpleCADforDSH '+n]];s.getRange('B2').format.font.size=16;s.getRange('B2').format.font.bold=true;}
set('概览','B3','MERGE-BUDGET-02 项目计划与预算。AT未执行，CodeArts门禁未通过。');sh['概览'].getRange('B3:M3').format.font.size=10;
// Task schedule and derived handoff gates.
head('任务计划',5,['任务ID','迭代','成员','开发日','开始h','结束h','全期开始h','全期结束h','人时','WBS','任务','交付物','依赖关口','依赖就绪h','检查']);
const start=6,end=start+d.tasks.length-1;const taskRow=new Map(d.tasks.map((t,i)=>[t.id,start+i]));
head('任务计划',135,['关口','迭代','就绪h','交接任务ID']);
const gateCell=new Map();d.gates.forEach((g,i)=>{const r=136+i;gateCell.set(g.iteration+':'+g.gate,'D'+r);set('任务计划','B'+r,g.gate);set('任务计划','C'+r,g.iteration);fx('任务计划','D'+r,'=MAX('+g.producer_ids.map(id=>'I'+taskRow.get(id)).join(',')+')');set('任务计划','E'+r,g.producer_ids.join(', '));});
d.tasks.forEach((t,i)=>{const r=start+i;sh['任务计划'].getRange('B'+r+':P'+r).values=[[t.id,t.iteration,t.person,t.day,t.slot,t.slot+1,null,null,null,t.wbs,t.task,t.deliverable,t.dependencies.join('、'),null,null]];fx('任务计划','H'+r,`=(E${r}-1)*6+F${r}`);fx('任务计划','I'+r,`=(E${r}-1)*6+G${r}`);fx('任务计划','J'+r,`=G${r}-F${r}`);fx('任务计划','O'+r,t.dependencies.length?'=MAX('+t.dependencies.map(g=>gateCell.get(t.iteration+':'+g)).join(',')+')':'=0');fx('任务计划','P'+r,`=IF(H${r}>=O${r},"满足列明依赖","依赖未就绪")`);});
// Gate formulas use finish column I, duration column J.
const taskSheet="'任务计划'!";const dur=taskSheet+`$J$${start}:$J$${end}`;const who=taskSheet+`$D$${start}:$D$${end}`;const it=taskSheet+`$C$${start}:$C$${end}`;const day=taskSheet+`$E$${start}:$E$${end}`;
sh['任务计划'].getRange('L6:N125').format.wrapText=true;sh['任务计划'].getRange('L6:N125').format.rowHeight=68;sh['任务计划'].getRange('L:L').format.columnWidth=34;sh['任务计划'].getRange('M:M').format.columnWidth=37;sh['任务计划'].getRange('N:N').format.columnWidth=34;for(const [c,w] of Object.entries({C:8,D:12,E:10,F:10,G:10,H:14,I:14,J:8,K:12,O:14,P:22}))sh['任务计划'].getRange(c+':'+c).format.columnWidth=w;
sh['任务计划'].freezePanes.freezeRows(5);sh['任务计划'].getRange('P6:P125').conditionalFormats.add('containsText',{text:'未就绪',format:{fill:'#FDE9E7',font:{color:'#A52222'}}});
// Per iteration and per day capacity (only net development; managers are separate).
head('容量',5,['成员','迭代1h 原基线','迭代2h 计划','迭代3h 计划','每迭代容量h','迭代2余量h','迭代3余量h']);
d.people.forEach((n,i)=>{const r=6+i;set('容量','B'+r,n);set('容量','C'+r,d.baseline1[i]);fx('容量','D'+r,`=SUMIFS(${dur},${who},B${r},${it},2)`);fx('容量','E'+r,`=SUMIFS(${dur},${who},B${r},${it},3)`);set('容量','F'+r,12);fx('容量','G'+r,`=F${r}-D${r}`);fx('容量','H'+r,`=F${r}-E${r}`);});
set('容量','B12','合计');for(const c of ['C','D','E','F','G','H'])fx('容量',c+'12',`=SUM(${c}6:${c}11)`);
head('容量',16,['成员','D3 h','D4 h','D5 h','D6 h','每日容量h','最高日负荷h']);d.people.forEach((n,i)=>{const r=17+i;set('容量','B'+r,n);[3,4,5,6].forEach((a,j)=>fx('容量',String.fromCharCode(67+j)+r,`=SUMIFS(${dur},${who},B${r},${day},${a})`));set('容量','G'+r,6);fx('容量','H'+r,`=MAX(C${r}:F${r})`);});
set('容量','B25','检查按当前任务时段及列明交接关口计算；技能、环境与实际效率须开工前复核。');sh['容量'].getRange('C6:H22').setNumberFormat('0');
// Assumptions and budget. Blank cloud/licence rates propagate when selected.
head('预算',4,['输入/控制','数值','含义/来源']);const inputs=[['机会成本 元/h',60,'XZJ-ECO-01，非工资'],['单次调用费q 元',1,'敏感性输入，非官方报价'],['云GPU触发 0/1',0,'0未触发，1触发'],['备用GPU小时',8,'用量情景'],['GPU单价g 元/h',null,'触发时需真实报价'],['许可触发 0/1',0,'0不购买，1购买'],['许可费用L 元',null,'触发时需真实报价'],['每迭代管理h',16,'PO/SM各8h，计划估算']];sh['预算'].getRange('B5:D12').values=inputs;sh['预算'].getRange('C5:C12').format.fill='#FFF3CE';sh['预算'].getRange('C5:C12').format.font.color='#0000CC';sh['预算'].getRange('D5:D12').format.columnWidth=38;
sh['预算'].getRange('C7').dataValidation={rule:{type:'list',values:['0','1']}};sh['预算'].getRange('C10').dataValidation={rule:{type:'list',values:['0','1']}};
head('预算',15,['迭代','开发h','管理h','合计h','机会成本元','调用次数','调用现金元','云GPU元','许可元','现金合计元','综合情景元']);
for(let i=0;i<3;i++){const r=16+i;set('预算','B'+r,i+1);fx('预算','C'+r,i===0?"='容量'!C12":`=SUMIFS(${dur},${it},${i+1})`);fx('预算','D'+r,'=$C$12');fx('预算','E'+r,`=SUM(C${r}:D${r})`);fx('预算','F'+r,`=E${r}*$C$5`);set('预算','G'+r,[500,300,200][i]);fx('预算','H'+r,`=IF(ISNUMBER($C$6),G${r}*$C$6,"待计费参数")`);fx('预算','I'+r,i===0?'=IF($C$7=0,0,IF(ISNUMBER($C$9),$C$8*$C$9,"待GPU报价"))':'=0');fx('预算','J'+r,i===0?'=IF($C$10=0,0,IF(ISNUMBER($C$11),$C$11,"待许可报价"))':'=0');fx('预算','K'+r,`=IF(AND(ISNUMBER(H${r}),ISNUMBER(I${r}),ISNUMBER(J${r})),SUM(H${r}:J${r}),"待报价")`);fx('预算','L'+r,`=IF(ISNUMBER(K${r}),F${r}+K${r},"待报价")`);}
set('预算','B19','合计');for(const c of ['C','D','E','F','G'])fx('预算',c+'19',`=SUM(${c}16:${c}18)`);for(const c of ['H','I','J','K','L'])fx('预算',c+'19',`=IF(COUNT(${c}16:${c}18)=3,SUM(${c}16:${c}18),"待报价")`);
set('预算','B22','首迭代条件开发增量h');fx('预算','C22',"='风险'!F16");set('预算','B23','条件管理增量h');fx('预算','C23',"='风险'!G16");set('预算','B24','全期含风险机会成本元');fx('预算','C24','=F19+SUM(C22:C23)*C5');set('预算','B25','全期含风险综合情景元');fx('预算','C25','=IF(ISNUMBER(K19),C24+K19,"待报价")');
head('预算',28,['q 元/次 敏感性','调用现金元']);[.5,1,3].forEach((q,i)=>{set('预算','B'+(29+i),q);fx('预算','C'+(29+i),`=B${29+i}*$G$19`);});
set('预算','B34','来源：讨论三66h；后两迭代WBS拆分与2026-11-04计划重排；管理16h/迭代及调用数量为分析假设。');
// Risk table actual source, separate inclusion and incremental buckets.
head('风险',5,['风险编号','风险','原应对h','已含h','条件开发h','管理增量h','触发条件']);
d.risks.forEach((a,i)=>{const r=6+i;const inc=Number(a['条件或管理增量人时']);const mgmt=a['增量分类']==='管理';sh['风险'].getRange('B'+r+':H'+r).values=[[a['风险编号'],a['风险名称'],Number(a['原应对估算人时']),Number(a['已含基线人时']),mgmt?0:inc,mgmt?inc:0,a['触发条件']]];});set('风险','B16','合计');for(const c of ['D','E','F','G'])fx('风险',c+'16',`=SUM(${c}6:${c}15)`);set('风险','B19','首迭代基线开发h');set('风险','C19',66);set('风险','B20','全部条件开发触发h');fx('风险','C20','=C19+F16');set('风险','B21','首迭代开发容量h');set('风险','C21',72);set('风险','B22','开发容量缺口h');fx('风险','C22','=MAX(0,C20-C21)');set('风险','B23','含管理总人时h');fx('风险','C23','=C20+G16');set('风险','B24','含管理机会成本元');fx('风险','C24',"=C23*'预算'!C5");set('风险','B27','来源：讨论五十大风险与22h；12已含、9条件开发、1管理来自逐任务去重，未视为风险已发生。');
sh['风险'].getRange('C6:C15').format.columnWidth=28;sh['风险'].getRange('H6:H15').format.columnWidth=53;
// Economic scenario shared across report, no project budget double count.
set('经济比较','B5','同类合格任务数N');set('经济比较','C5',200);sh['经济比较'].getRange('C5').format.fill='#FFF3CE';
head('经济比较',7,['路线','初始人工h','初始现金元','单任务人工min','单任务现金元','总人工h','机会成本元','现金元','总成本元']);
const routes=[['A 人工CAD',0,0,30,0],['B 扩展插件',60,0,10,1],['C 自研/训练',200,1000,8,.5]];
routes.forEach((a,i)=>{const r=8+i;sh['经济比较'].getRange('B'+r+':F'+r).values=[a];fx('经济比较','G'+r,`=C${r}+$C$5*E${r}/60`);fx('经济比较','H'+r,`=G${r}*'预算'!$C$5`);fx('经济比较','I'+r,`=D${r}+$C$5*F${r}`);fx('经济比较','J'+r,`=SUM(H${r}:I${r})`);});
set('经济比较','B13','B相对A回收任务数');fx('经济比较','C13',"=ROUNDUP((C9*'预算'!C5+D9)/(E8/60*'预算'!C5+F8-E9/60*'预算'!C5-F9),0)");set('经济比较','B14','C与B成本相等任务数');fx('经济比较','C14',"=(C10*'预算'!C5+D10-C9*'预算'!C5-D9)/(E9/60*'预算'!C5+F9-E10/60*'预算'!C5-F10)");set('经济比较','B15','B在N情景净省人工h');fx('经济比较','C15','=G8-G9');set('经济比较','B16','B在N情景净省成本元');fx('经济比较','C16','=J8-J9');set('经济比较','B19','来源：XZJ-ECO-01，所有费率及时间均为情景输入，不是工资、实测收益或报价。路线初始工时不叠加到项目186开发小时。');
sh['经济比较'].getRange('G8:J10').setNumberFormat('#,##0.00');
// Overview links outputs, never from audit cells.
head('概览',5,['迭代','开发人时','管理人时','机会成本元','现金情景元','综合情景元']);for(let i=0;i<3;i++){const r=6+i,br=16+i;fx('概览','B'+r,`='预算'!B${br}`);fx('概览','C'+r,`='预算'!C${br}`);fx('概览','D'+r,`='预算'!D${br}`);fx('概览','E'+r,`='预算'!F${br}`);fx('概览','F'+r,`='预算'!K${br}`);fx('概览','G'+r,`='预算'!L${br}`);}set('概览','B9','合计');for(const [c,b]of [['C','C'],['D','D'],['E','F'],['F','K'],['G','L']])fx('概览',c+'9',`='预算'!${b}19`);
set('概览','B12','首迭代风险全部触发的开发缺口h');fx('概览','E12',"='风险'!C22");set('概览','B13','全期含风险机会成本元');fx('概览','E13',"='预算'!C24");set('概览','B14','含风险综合情景元');fx('概览','E14',"='预算'!C25");set('概览','B17','D3—D6当前时段安排无同人冲突；每人每日≤6h、每迭代≤12h。准备任务不等同于AT已执行。');set('概览','B19','未触发GPU/许可时无此类增量；触发后缺报价则显示待报价，不计作零元。');
// Fit bounded inputs and results. Long schedule and note cells have explicit widths.
for(const n of names){const s=sh[n];s.getRange('B:B').format.columnWidth=27;s.getRange('B2:Q2').format.rowHeight=32;s.getRange('B3:Q3').format.rowHeight=28;}
sh['任务计划'].getRange('B:B').format.columnWidth=20;sh['预算'].getRange('B5:B12').format.columnWidth=32;sh['风险'].getRange('B:B').format.columnWidth=30;sh['容量'].getRange('C:H').format.columnWidth=19;sh['概览'].getRange('C:G').format.columnWidth=21;
sh['经济比较'].getRange('C15').setNumberFormat('0.00');sh['预算'].getRange('F16:L19').setNumberFormat('#,##0.00');sh['概览'].getRange('E6:G14').setNumberFormat('#,##0.00');
// Perturbation checks: q updates costs; selecting missing cloud price propagates missing input; actual zero is valid.
wb.recalculate();const checks={};checks.base=sh['概览'].getRange('G9').values[0][0];if(checks.base!==15040)throw Error('Base mismatch '+JSON.stringify(checks));
set('预算','C6',2);wb.recalculate();checks.q2=sh['概览'].getRange('G9').values[0][0];if(checks.q2!==16040)throw Error('q sensitivity');set('预算','C6',1);set('预算','C7',1);wb.recalculate();checks.missingGpu=sh['概览'].getRange('G9').values[0][0];if(checks.missingGpu!=='待报价')throw Error('missing GPU swallowed');set('预算','C9',0);wb.recalculate();checks.zeroGpu=sh['概览'].getRange('G9').values[0][0];if(checks.zeroGpu!==15040)throw Error('zero GPU invalid');set('预算','C7',0);set('预算','C9',null);wb.recalculate();
checks.current=sh['概览'].getRange('B5:G14').values;checks.risks=sh['风险'].getRange('D16:G16').values;checks.capacity=sh['容量'].getRange('B6:H12').values;
if(JSON.stringify(checks.risks)!==JSON.stringify([[22,12,9,1]]))throw Error('Risk buckets');
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'Formula error scan'});checks.errorScan=errors.ndjson;
for(const n of names){const range={概览:'B2:G19',任务计划:'B5:P11',容量:'B5:H22',预算:'B4:L19',风险:'B5:H16',经济比较:'B5:J16'}[n];const blob=await wb.render({sheetName:n,range,scale:1.5,format:'png'});await fs.writeFile('/tmp/simplecad-'+n+'.png',new Uint8Array(await blob.arrayBuffer()));}
const gatePreview=await wb.render({sheetName:'任务计划',range:'B135:P142',scale:1.5,format:'png'});await fs.writeFile('/tmp/simplecad-交接关口.png',new Uint8Array(await gatePreview.arrayBuffer()));
const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(out+'/SimpleCADforDSH计划与预算.xlsx');try{await fs.rename(out+'/SimpleCADforDSH计划与预算.xlsx.inspect.ndjson',src+'/工作簿公式检查.ndjson');}catch(e){if(e.code!=='ENOENT')throw e;}await fs.writeFile(src+'/workbook-verification.json',JSON.stringify(checks,null,2));console.log(JSON.stringify({base:checks.base,q2:checks.q2,missingGpu:checks.missingGpu,zeroGpu:checks.zeroGpu,risks:checks.risks,errors:checks.errorScan}));
