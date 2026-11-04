# 合稿来源与复现

本目录保留合稿脚本、结构化正文、来源映射、原始附件副本、当前计划数据及验证记录。四份中文姓名目录下的原始分稿只读，input-sha256.json记录输入指纹。

当前计划版本为MERGE-BUDGET-02（2026年11月4日），按工作包、人时和资源容量组织计划与预算。权威输入为计划数据.json，计划与预算工作簿由该文件生成。originals/中的历史工作簿、CSV和讨论资料保留原口径，不能代替本轮同版工作簿；风险统一按12已含、9条件开发、1管理处理。

## 文件职责

- build_report.py：读取四份原稿与课程模板，重组正文和附录，生成DOCX、content.json、来源映射.json、预算CSV/JSON和图表；adopt_report_language统一正式表述。
- release_revision.py：对正式表述和WBS字典收口；生成D3—D6共120条一小时任务、21个交接关口、逐人安排、WBS连线源图和资源甘特图。任务估算及人员投入均未被记为实际执行。
- build_workbook.mjs：通过bundled @oai/artifact-tool创建SimpleCADforDSH计划与预算.xlsx，包含概览、任务计划、容量、预算、风险、经济比较六表。保留公式、情景输入、缺报价传播及敏感性检查。
- verify_report.py：独立复算任务时段、工作包分配、依赖、容量、预算，回读Excel公式缓存，核对目录、模板格式、正文回读与原稿指纹。
- 满分标准核查.md：评分范围与报告落点，不预测成绩。

## 重建步骤

使用Codex workspace dependency loader返回的bundled Python/Node；本次路径分别为：

```
/Users/pano/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
/Users/pano/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node
```

1. 用bundled Python运行source/build_report.py；它调用release_revision.py并输出计划数据。
2. 用bundled Node运行source/build_workbook.mjs。依赖从bundled node_modules解析，临时链接在/tmp/simplecad-report-workbook-runtime；不安装仓库依赖。
3. 设置FONTCONFIG_FILE为source/fonts.conf，用documents技能render_docx.py生成页面PNG及PDF。Fontconfig将宋体映射到macOS Songti SC，DOCX仍保留宋体12磅和固定20磅。渲染使用bundled LibreOffice。
4. 再执行build_report.py --toc-pdf <上次PDF>，重新渲染，直至51项目录页码稳定。目录为与PDF相符的可见页码缓存；编辑后按此步骤重建。当前目录不依赖Word刷新域，不承诺F9可自动重算这些可见页码。
5. 逐页查看PNG，检查内容、表格、图像和留白；将同版PDF复制到merged目录，用bundled Python执行verify_report.py <最终PDF>。

模板合同见template-contract.json和artifact.md；交付统计见构建统计.json和verification.json；工作簿情景检查见workbook-verification.json。临时渲染页图在/tmp，不随交付保存。WBS图源与甘特图在../assets/，任务分配和来源同时保留CSV/JSON。

预算、人员任务细分及异常AT规则与报告保持一致。CodeArts检查门禁未通过，AT仍为未执行。产品开发、门禁修复与测试执行结果依据各自实际记录维护。
