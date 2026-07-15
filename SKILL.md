---
name: write-feasibility-study-sgcc
description: 编写、修订、排版、绘图、编制预算并质检国家电网有限公司（SGCC）科技项目中文可行性研究报告。适用于用户要求新建或修改国网、国家电网、SGCC科技项目可研，包括DOCX版本管理、章节撰写、原生MathType公式与变量、参考文献、技术路线图、进度计划、预算附件以及交付前质量检查。进入Word交付模式时，本技能自动加载word-mathtype-native-equations生成Equation.DSMT4对象和MathType自动编号，用户无需另行调用。本技能仅适用于国网科技项目，不得套用南方电网或其他通用可研格式。
---

# 国网科技项目可研写作

完成国网科技项目可研全文，或在保持已确认Word模板的前提下修订指定内容。将内容质量、Word格式、公式、图表及文档完整性作为一个统一流程处理。

## 自动加载原生MathType依赖

用户只需调用本技能，不得要求用户在提示词中再次调用`word-mathtype-native-equations`。

进入Word交付模式后，在检查或写入正文之前自动执行：

1. 从当前可用技能目录定位`word-mathtype-native-equations/SKILL.md`；当前本机默认位于`$CODEX_HOME/skills/word-mathtype-native-equations/SKILL.md`。
2. 完整读取该`SKILL.md`。任务涉及Word公式、变量、自动编号、分栏或变量解释时，继续完整读取其中指定的`references/layout-numbering.md`和`references/automation.md`。
3. 将该技能作为内部必需依赖执行，无需也不得向用户索要第二次技能调用。
4. 本技能负责可研内容和国网配置，依赖技能负责MathType对象生成、Word COM、自动编号和验证；发生冲突时，MathType实现细节以依赖技能最新规则为准，国网编号口径和报告内容以本技能为准。
5. 无法定位依赖或MathType预检失败时，停止最终Word公式阶段并报告具体原因；禁止退回静态编号、Word OMath或LaTeX文本冒充完成。

用户明确只要求“直接在窗口发我”且不创建DOCX时，不加载该依赖。

## 选择交付方式

- 用户要求“直接在窗口发我”时，仅在聊天中撰写所需文字，不创建或修改DOCX。
- 新建完整可研时，先取得或检索已确认的国网模板，建立项目简报，完成各模块草稿，再生成`V1.docx`。
- 修订既有可研时，将用户最新手工修改版作为唯一母版，复制为下一版本，只修改用户指定范围。
- 禁止覆盖用户文件，禁止无提示地退回旧版本。
- 除非用户明确要求，否则不生成PDF。
- 禁止使用界面自动化或computer-use插件。普通内容可采用OOXML或独立后台Word COM批量编辑；公式和数学变量必须调用`word-mathtype-native-equations`规定的静默Word COM与MathType宏流程。不得执行`.ps1`文件。

## 按任务加载规则

只读取当前任务所需的规则；创建或编辑DOCX前，必须读取`references/word-format-spec.md`。

- 撰写或改写章节时，读取`references/module-writing-rules.md`。
- 选择Word母版或本地参考资料前，读取`references/template-and-source-routing.md`。
- 在用户本地环境工作时，读取`references/local-gold-sources.md`；公开版本只提供空白登记模板，实际路径必须由使用者在私有副本中填写并核验。
- 处理DOCX版式、样式、段落、表格、图题表题、分页和版本时，读取`references/word-format-spec.md`。
- 出现公式、数学变量、参考文献或研究现状时，读取`references/formula-and-citation-spec.md`。Word模式中的MathType依赖由本技能按上一节自动加载，不要求用户另行调用。
- 新建或替换任何图片前，读取`references/figure-spec.md`。
- 处理预算、附件、成果或进度计划时，读取`references/budget-and-schedule-spec.md`。
- 编写完整可研或进行大范围修订前，读取`references/user-rule-registry.md`。
- 最终交付前，读取`references/qa-rules.md`。

## 建立项目简报

在正式报告外记录以下事实：

- 项目名称、申报单位、项目地点和既有基础；
- 起止时间和总经费；
- 已确认的研究内容标题及其内部结构；
- 预期成果和可量化验收指标；
- 最新DOCX或已确认国网模板路径；
- 用户明确认可的本地参考可研和项目资料；
- 仍需确认的关键缺失信息。

以`assets/project-brief-template.yaml`为字段框架。只针对阻断性事实提问，每次不超过三个问题；非关键细节可依据已确认的本地材料合理推断。将假设和核实事项留在正式报告之外，禁止把面向助手的免责声明写入可研。

## 检索证据

优先使用用户的本地科研索引，再考虑较大范围的文件检索。仅以只读方式打开完成任务所需的少量、已确认国网可研和项目材料。严格区分：

- 当前项目已经核实的事实；
- 可以复用的国网写法与格式规律；
- 只属于其他项目、不得移植的案例事实。

政策、标准、现行制度或机构事实可能变化时，检索并核验权威的一手来源。禁止虚构政策文件、项目基础、实验室、承担单位、人员、平台、数据源或性能结果。

## 先设计研究逻辑

建立完整逻辑链：

`实际需求 → 待解决问题 → 研究内容 → 技术路线 → 创新点 → 成果 → 验收指标`

先确认研究内容标题和内部任务结构，再绘图或扩展公式。正式“研究内容”小节原则上每项写成一个完整段落，除非模板另有要求。

## 撰写报告

遵循`references/module-writing-rules.md`，按下列顺序形成内容：

1. 项目背景与意义；
2. 含引用的国内外研究现状；
3. 理论依据与实践依据；
4. 研究内容；
5. 实施方案与技术路线；
6. 创新点；
7. 预期目标、成果及指标；
8. 组织实施、条件、进度与预算；
9. 参考文献和附件。

禁止按照经费比例分配篇幅，禁止为了对应研究点而机械拆分成果。正式报告不得出现助手解释、保留意见、资料核实提醒或其他元话语。

## 处理公式、变量与引用

将职责分开：本技能决定公式内容、变量定义、国网编号配置和引用格式；自动加载的`word-mathtype-native-equations`负责把公式与变量写成原生`Equation.DSMT4`对象、插入MathType自动编号字段、计算制表位并验证DOCX。

- 聊天草稿可以使用LaTeX表达数学内容；最终Word不得残留`$...$`、`\[...\]`或MathType占位符。
- 新建国网可研默认采用半角括号、按章编号：MathType模板`(#C1-#E1)`，预设`ascii-chapter-dash`；从第一章开始在每个主章边界推进MathType章号，即使该章没有公式也不得跳过。若用户最新母版已有可用MathType编号和交叉引用，则优先保持原格式，不得擅自重建。
- 正文变量、显示公式和“式中，”变量解释里的变量均转换为原生MathType OLE对象；单位、中文术语、标点和普通运算说明保留为Word文本。
- “式中，”变量解释段落左缩进、首行缩进和字符单位缩进均为零。
- 对每个新增变量说明含义、单位、取值范围或物理意义。
- 参考文献引用继续执行`references/formula-and-citation-spec.md`中的上标顺序编码规则。
- 若`word-mathtype-native-equations`不可用或MathType预检失败，停止最终公式写入，仅保留未交付的内部草稿并说明阻断原因；禁止退回静态编号、Word OMath或LaTeX文本冒充完成结果。

## 绘制图片

执行`references/figure-spec.md`。先确定文字内容和逻辑结构，再绘制图片。保留可编辑SVG或draw.io源文件，向Word插入高分辨率PNG。用户确认某张分项图后，只按要求局部修改，不重新绘制整套图片。

## 安全编辑Word

修订既有报告时：

1. 对母版运行`scripts/inspect_docx.py`；
2. 运行`scripts/clone_version.py`生成下一版本；
3. 只在副本中修改文字、媒体和明确指定的公式范围；
4. 保留分节符、样式、表格、页眉页脚、关系文件、既有OLE对象和嵌入对象；新增公式导致的对象变化必须与预期数量一致；
5. 普通内容优先使用Range、Styles、Find、整表或连续区域批量处理；避免Selection及逐字符、逐段落、逐单元格循环；
6. MathType转换只使用新建的不可见Word实例，关闭提示和屏幕刷新，不附着、不结束用户已有Word实例；
7. 运行`scripts/qa_docx.py`完成国网结构检查，再运行`word-mathtype-native-equations/scripts/validate_mathtype_docx.py`完成原生MathType检查；
8. 使用Windows PowerShell 5.1的`-NoProfile -Command`内联执行Word COM，不创建或执行`.ps1`文件；保存后以新的隐藏Word实例只读复核。

新建报告时，应复制用户确认的国网模板，复用其样式和表格，不得从空白Word文档拼装，也不得以南网或通用可研作为格式母版。

## 交付前校验

执行`references/qa-rules.md`，至少检查：

- 研究逻辑与标题是否一致；
- 公式是否为`Equation.DSMT4`、Word OMath是否为零、MathType自动编号字段与可见编号是否正确、变量解释是否完整；
- 标题、正文、图题、表格和参考文献格式是否一致；
- 图片文字是否清晰，箭头是否对齐；
- 项目时间是否一致，经费计算是否平衡；
- 用户未要求时不得保留目录；
- 图片、OLE对象、媒体和嵌入对象是否完整，新增MathType对象数量是否与计划一致；
- DOCX压缩包结构是否有效；具备Word时能否正常打开并保存。

只交付最新版本，并附简洁的修改摘要和绝对路径链接。
