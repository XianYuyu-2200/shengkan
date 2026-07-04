# Engineering Competition Course Integration Writing Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a province-journal-ready Chinese teaching-reform paper on the “four-in-one, dual-cycle” course-competition integration model for engineering competitions such as Robomaster, Robocon, and the National Undergraduate Electronics Design Contest.

**Architecture:** The manuscript will be built as a staged academic writing project: confirmed design specification, case-material intake, detailed outline, section-by-section draft, citation and policy support, revision, and submission package. The writing follows a “mode construction + practice path + experience summary” structure, with anonymous local applied undergraduate university practice as the case support.

**Tech Stack:** Markdown manuscript drafts, Chinese academic writing conventions, CNKI/Wanfang-style literature retrieval when references are needed, GB/T 7714-style reference formatting, optional Word `.docx` export after the manuscript stabilizes.

## Global Constraints

- Target positioning: province-journal publication first, with structure reserved for later upgrade toward Chinese core-journal quality.
- Topic scope: higher-education engineering discipline competitions, especially Robomaster, Robocon, and the National Undergraduate Electronics Design Contest.
- Paper type: teaching-reform and practice research, not pure empirical research.
- Case identity: use anonymous wording such as “某地方应用型本科高校工科学院”.
- Core model: “课程教学、项目训练、竞赛实践、成果转化” four-in-one structure plus “课程支撑竞赛、竞赛反哺课程” dual-cycle mechanism.
- Data rule: do not invent concrete numbers; if exact data are unavailable, use experience-summary expressions and mark where real data can strengthen the argument.
- Writing focus: competition serves education; do not turn the paper into a competition award report.

---

## File Structure

- Modify: `docs/superpowers/specs/2026-07-04-engineering-competition-course-integration-design.md`
  - Role: approved design source. Only update if the user changes the agreed topic, model, or target journal positioning.
- Create: `manuscript/outline.md`
  - Role: detailed paper outline with title, abstract, keywords, first-level and second-level headings, and paragraph-level writing notes.
- Create: `manuscript/case-materials.md`
  - Role: structured intake sheet for school background, competitions, courses, teams, achievements, and available evidence.
- Create: `manuscript/draft.md`
  - Role: full Chinese manuscript draft.
- Create: `manuscript/references.md`
  - Role: candidate literature list, policy references, retrieval notes, and GB/T 7714-style formatted references.
- Create: `manuscript/revision-checklist.md`
  - Role: self-review checklist for province-journal fit, argument coherence, data risk, citation sufficiency, and language polish.
- Optional create after draft approval: `manuscript/submission-version.docx`
  - Role: Word version for journal submission.

## Task 1: Case Material Intake

**Files:**
- Create: `manuscript/case-materials.md`

**Interfaces:**
- Consumes: approved design spec from `docs/superpowers/specs/2026-07-04-engineering-competition-course-integration-design.md`
- Produces: structured case facts that later tasks use to write the practice path and effectiveness sections

- [ ] **Step 1: Create the manuscript directory**

Run:

```powershell
New-Item -ItemType Directory -Force -Path manuscript | Out-Null
```

Expected: `manuscript/` exists.

- [ ] **Step 2: Create the case material intake document**

Create `manuscript/case-materials.md` with this content:

```markdown
# 案例材料收集表

## 1. 案例单位匿名表述

建议使用：某地方应用型本科高校工科学院。

可替代表述：

1. 某地方高校工程类专业群
2. 某应用型本科高校智能制造与电子信息相关专业群

最终稿默认使用第一种表述。

## 2. 涉及竞赛

| 竞赛名称 | 可写入论文的角色 | 可支撑的能力培养 |
| --- | --- | --- |
| Robomaster | 综合工程实践载体 | 机械结构、嵌入式控制、视觉识别、团队协作 |
| Robocon | 机器人系统设计载体 | 机械设计、运动控制、任务分解、系统集成 |
| 全国大学生电子设计竞赛 | 电子信息实践载体 | 电路设计、嵌入式开发、调试测试、创新设计 |

## 3. 课程与竞赛衔接素材

| 课程或课程群 | 可衔接的竞赛任务 | 可写入论文的表达 |
| --- | --- | --- |
| 电路与模拟电子技术 | 传感器采集、电源模块、信号调理 | 将基础电路知识转化为竞赛硬件模块设计任务 |
| 单片机/嵌入式系统 | 控制系统、通信模块、底盘控制 | 以项目任务强化嵌入式系统综合应用能力 |
| 自动控制原理 | 运动控制、闭环控制、PID 调试 | 将控制理论融入机器人运动控制实践 |
| 机械设计基础 | 机构设计、传动设计、结构优化 | 支撑机器人机械结构设计与迭代优化 |
| 计算机视觉/人工智能基础 | 目标识别、图像处理、智能决策 | 服务 Robomaster 等赛事的视觉识别与智能控制任务 |

## 4. 可用成果材料

如有真实材料，补充在对应行；如暂无精确数据，正式论文采用概括性表述。

| 材料类型 | 是否已有 | 可采用的写法 |
| --- | --- | --- |
| 参赛年份 | 不确定 | 近年来持续组织学生参加工程类学科竞赛 |
| 参赛人数 | 不确定 | 形成跨年级、跨专业学生参与机制 |
| 获奖情况 | 不确定 | 在相关竞赛中取得一定成效 |
| 课程项目数量 | 不确定 | 形成若干来源于竞赛任务的课程项目案例 |
| 工作室/训练平台 | 不确定 | 依托专业实验室和学生创新工作室开展项目训练 |
| 教师团队 | 不确定 | 组建由不同专业教师参与的导师组 |

## 5. 数据使用原则

1. 未经确认的数字不写入正文。
2. 没有精确数据时，使用“逐步形成”“有效促进”“一定程度提升”等审慎表达。
3. 若后续提供真实数据，可优先补入参赛人数、获奖数量、课程项目数量、学生创新创业项目数量和毕业设计转化数量。
```

- [ ] **Step 3: Verify intake document**

Run:

```powershell
Get-Content -Raw -Encoding UTF8 manuscript\case-materials.md
```

Expected: the table headings display correctly in Chinese, with no garbled characters.

- [ ] **Step 4: Commit**

Run:

```powershell
git add manuscript/case-materials.md
git commit -m "docs: add case material intake for competition paper"
```

Expected: commit succeeds.

## Task 2: Detailed Paper Outline

**Files:**
- Create: `manuscript/outline.md`

**Interfaces:**
- Consumes: `manuscript/case-materials.md`
- Produces: paragraph-level structure used to draft `manuscript/draft.md`

- [ ] **Step 1: Create the detailed outline**

Create `manuscript/outline.md` with this content:

```markdown
# 高校工程类学科竞赛赛教融合育人模式构建与实践研究：详细大纲

## 题目

高校工程类学科竞赛赛教融合育人模式构建与实践研究

## 摘要写作要点

1. 背景：新工科建设和应用型本科高校工程人才培养需要强化实践能力。
2. 问题：竞赛与课程教学脱节、学生参与覆盖面有限、成果转化不足。
3. 方法：以 Robomaster、Robocon、电赛等竞赛为载体，构建“四位一体、双循环”模式。
4. 结论：该模式有助于提升学生综合工程能力、跨专业协同能力和课程教学质量。

## 关键词

工程类学科竞赛；赛教融合；新工科；实践教学；应用型本科高校；创新人才培养

## 一、引言

### 1. 新工科背景下工程实践能力培养的重要性

写作重点：说明地方应用型本科高校承担服务区域产业和培养应用型工程人才的任务，工程实践能力、创新能力和跨专业协同能力成为人才培养的重要目标。

### 2. 工程类学科竞赛的育人价值

写作重点：说明 Robomaster、Robocon、电赛等赛事具有真实任务、复杂系统、团队协作和持续迭代等特点，能弥补传统课堂实践不足。

### 3. 当前问题与本文切入点

写作重点：指出竞赛与课程脱节、覆盖面不足、教师协同不足、成果转化不足。提出本文构建“四位一体、双循环”赛教融合模式。

## 二、高校工程类学科竞赛赛教融合的现实价值

### 1. 促进理论知识向工程实践转化

写作重点：课程知识通过竞赛项目转化为硬件设计、控制调试、系统集成和工程验证。

### 2. 推动跨专业协同与项目化学习

写作重点：机器人和电子设计任务需要电子、机械、自动化、计算机等学生共同完成。

### 3. 提升创新意识和复杂工程问题解决能力

写作重点：竞赛任务具有不确定性和迭代性，可训练学生问题分析、方案设计、测试优化和团队沟通能力。

## 三、当前赛教融合存在的主要问题

### 1. 课程体系与竞赛任务衔接不够紧密

写作重点：课程内容偏知识传授，竞赛训练偏课外突击，二者缺少系统连接。

### 2. 竞赛训练覆盖面不足

写作重点：部分竞赛训练集中于少数骨干学生，普通学生难以分享竞赛育人成果。

### 3. 指导教师协同机制不完善

写作重点：跨专业赛事需要多学科教师协同，但实际中可能存在分工不清、持续指导不足等问题。

### 4. 竞赛成果转化不足

写作重点：竞赛中的项目、案例、技术问题没有充分进入课堂、实验、毕业设计和创新创业项目。

## 四、“四位一体、双循环”赛教融合育人模式构建

### 1. 课程教学：夯实竞赛能力基础

写作重点：建立课程群支撑机制，将电子、机械、控制、计算机相关课程与竞赛能力对应。

### 2. 项目训练：搭建课程与竞赛之间的桥梁

写作重点：将复杂竞赛任务拆解为模块化、层级化项目，使不同年级学生都能参与。

### 3. 竞赛实践：完成综合工程能力训练

写作重点：通过真实竞赛任务训练学生系统设计、工程调试、团队协作和现场应变能力。

### 4. 成果转化：推动竞赛反哺教学

写作重点：将竞赛案例转化为课程案例、实验项目、毕业设计题目、创新创业项目和教改成果。

### 5. 双循环机制：课程支撑竞赛，竞赛反哺课程

写作重点：解释“课程到竞赛”和“竞赛到课程”的闭环关系，体现持续改进。

## 五、地方应用型本科高校的实践路径

### 1. 构建跨专业竞赛团队

写作重点：形成电子信息、机械、自动化、计算机等学生共同参与的团队结构。

### 2. 建设“课程群+工作室+项目制”训练平台

写作重点：课程群提供知识基础，工作室提供训练空间，项目制提供实践任务。

### 3. 建立导师组协同指导机制

写作重点：由不同专业教师组成导师组，围绕机械结构、电控系统、软件算法、项目管理开展协同指导。

### 4. 将竞赛任务融入课程项目

写作重点：将底盘控制、传感器采集、目标识别、电源设计等任务转化为课程项目。

### 5. 形成学生梯队培养机制

写作重点：低年级参与基础训练，高年级承担核心模块，优秀学生反向带动新成员。

## 六、实践成效与反思

### 1. 学生工程实践能力提升

写作重点：从系统设计、动手调试、工程验证和问题解决能力展开。

### 2. 跨专业协同能力增强

写作重点：学生在真实项目中理解不同专业分工，提升沟通协作能力。

### 3. 课程教学资源更加真实

写作重点：竞赛问题进入课堂，课程案例更贴近真实工程场景。

### 4. 教师教学改革能力提升

写作重点：教师通过竞赛指导积累案例、项目和教改素材。

### 5. 反思与改进方向

写作重点：完善评价体系、资源保障、导师激励和长期运行机制。

## 七、结语

写作重点：总结“四位一体、双循环”模式的价值，强调竞赛不是目的，育人才是目标。
```

- [ ] **Step 2: Verify outline completeness**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\outline.md -Pattern '^## |^### '
```

Expected: title, abstract, keywords, and seven main sections appear.

- [ ] **Step 3: Commit**

Run:

```powershell
git add manuscript/outline.md
git commit -m "docs: add detailed outline for competition paper"
```

Expected: commit succeeds.

## Task 3: Draft Introduction and Problem Diagnosis

**Files:**
- Create or modify: `manuscript/draft.md`

**Interfaces:**
- Consumes: `manuscript/outline.md`
- Produces: complete sections “一、引言”, “二、现实价值”, and “三、主要问题”

- [ ] **Step 1: Create draft scaffold**

Create `manuscript/draft.md` with title, abstract placeholder replaced by the approved abstract text, keywords, and headings:

```markdown
# 高校工程类学科竞赛赛教融合育人模式构建与实践研究

## 摘要

新工科建设背景下，工程类学科竞赛已成为高校培养学生工程实践能力、创新能力和团队协作能力的重要载体。然而，在实际推进过程中，部分高校仍存在竞赛与课程教学脱节、学生参与覆盖面有限、竞赛成果转化不足等问题。文章以 Robomaster、Robocon、全国大学生电子设计竞赛等工程类学科竞赛为切入点，结合地方应用型本科高校跨专业人才培养实践，构建“课程教学、项目训练、竞赛实践、成果转化”四位一体的赛教融合育人模式，并提出“课程支撑竞赛、竞赛反哺课程”的双循环运行机制。实践表明，该模式有助于促进课程教学与真实工程任务衔接，提升学生综合工程能力和跨专业协同能力，为地方高校推进新工科实践教学改革提供参考。

## 关键词

工程类学科竞赛；赛教融合；新工科；实践教学；应用型本科高校；创新人才培养

## 一、引言

## 二、高校工程类学科竞赛赛教融合的现实价值

## 三、当前赛教融合存在的主要问题

## 四、“四位一体、双循环”赛教融合育人模式构建

## 五、地方应用型本科高校的实践路径

## 六、实践成效与反思

## 七、结语
```

- [ ] **Step 2: Draft the introduction**

Write 800-1000 Chinese characters under “一、引言”. Include these points in this order:

1. New engineering education and applied undergraduate universities need practical engineering talent.
2. Robomaster, Robocon, and electronic design competitions provide authentic engineering tasks.
3. Existing problems include disconnection from courses, limited student coverage, weak teacher collaboration, and insufficient achievement transformation.
4. The paper proposes the “four-in-one, dual-cycle” model.

- [ ] **Step 3: Draft the realistic value section**

Write 900-1100 Chinese characters under “二、高校工程类学科竞赛赛教融合的现实价值”. Use three subheadings:

```markdown
### （一）促进理论知识向工程实践转化

### （二）推动跨专业协同与项目化学习

### （三）提升创新意识和复杂工程问题解决能力
```

- [ ] **Step 4: Draft the problem diagnosis section**

Write 1000-1200 Chinese characters under “三、当前赛教融合存在的主要问题”. Use four subheadings:

```markdown
### （一）课程体系与竞赛任务衔接不够紧密

### （二）竞赛训练参与覆盖面不足

### （三）指导教师协同机制不完善

### （四）竞赛成果转化机制不健全
```

- [ ] **Step 5: Verify draft length and headings**

Run:

```powershell
(Get-Content -Raw -Encoding UTF8 manuscript\draft.md).Length
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '^## |^### '
```

Expected: character length is above 2500 and the first three main sections contain subheadings where required.

- [ ] **Step 6: Commit**

Run:

```powershell
git add manuscript/draft.md
git commit -m "docs: draft introduction value and problem sections"
```

Expected: commit succeeds.

## Task 4: Draft the Core Model Section

**Files:**
- Modify: `manuscript/draft.md`

**Interfaces:**
- Consumes: problem diagnosis from Task 3
- Produces: core model section that later practice-path section implements

- [ ] **Step 1: Insert the model-section subheadings**

Under “四、“四位一体、双循环”赛教融合育人模式构建”, add:

```markdown
### （一）课程教学：夯实竞赛能力基础

### （二）项目训练：搭建课程与竞赛之间的桥梁

### （三）竞赛实践：强化综合工程能力训练

### （四）成果转化：推动竞赛反哺教学

### （五）双循环机制：课程支撑竞赛，竞赛反哺课程
```

- [ ] **Step 2: Draft the four-in-one structure**

Write 1400-1700 Chinese characters across the first four subheadings. Use this conceptual chain exactly once:

```text
课程教学—项目训练—竞赛实践—成果转化
```

- [ ] **Step 3: Draft the dual-cycle mechanism**

Write 600-800 Chinese characters under “双循环机制”. Include both phrases:

```text
课程到竞赛的正向循环
竞赛到课程的反向循环
```

- [ ] **Step 4: Add a text-based figure description**

Add this figure note after the dual-cycle mechanism:

```markdown
图1 “四位一体、双循环”赛教融合育人模式：课程教学为项目训练提供知识基础，项目训练将课程知识转化为模块化工程任务，竞赛实践完成综合能力训练，成果转化推动竞赛经验进入课程案例、实验项目和教学改革；同时形成课程支撑竞赛、竞赛反哺课程的双向循环。
```

- [ ] **Step 5: Verify model terminology consistency**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '四位一体|双循环|课程教学|项目训练|竞赛实践|成果转化'
```

Expected: all six model terms appear in section four.

- [ ] **Step 6: Commit**

Run:

```powershell
git add manuscript/draft.md
git commit -m "docs: draft four-in-one dual-cycle model"
```

Expected: commit succeeds.

## Task 5: Draft Practice Path and Effectiveness Sections

**Files:**
- Modify: `manuscript/draft.md`

**Interfaces:**
- Consumes: `manuscript/case-materials.md` and model section from Task 4
- Produces: anonymous case practice section and experience-summary effectiveness section

- [ ] **Step 1: Insert practice-path subheadings**

Under “五、地方应用型本科高校的实践路径”, add:

```markdown
### （一）构建跨专业竞赛团队

### （二）建设“课程群+工作室+项目制”训练平台

### （三）建立导师组协同指导机制

### （四）将竞赛任务融入课程项目

### （五）形成学生梯队培养机制
```

- [ ] **Step 2: Draft the practice-path section**

Write 1500-1800 Chinese characters. Use the anonymous case wording:

```text
以某地方应用型本科高校工科学院为例
```

Use each competition name at least once:

```text
Robomaster
Robocon
全国大学生电子设计竞赛
```

- [ ] **Step 3: Insert effectiveness subheadings**

Under “六、实践成效与反思”, add:

```markdown
### （一）学生工程实践能力得到增强

### （二）跨专业协作能力得到提升

### （三）课程教学资源更加贴近真实工程场景

### （四）教师教学改革与科研转化能力得到促进

### （五）持续改进中仍需关注的问题
```

- [ ] **Step 4: Draft the effectiveness and reflection section**

Write 1200-1500 Chinese characters. Avoid specific numbers unless the user has provided verified data. Use careful expressions such as:

```text
在一定程度上
逐步形成
有效促进
仍需进一步完善
```

- [ ] **Step 5: Verify no invented numbers**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '\d+人|\d+项|\d+次|\d+年|\d+个'
```

Expected: any matches are reviewed manually. Keep only verified numbers or remove them.

- [ ] **Step 6: Commit**

Run:

```powershell
git add manuscript/draft.md
git commit -m "docs: draft practice path and effectiveness sections"
```

Expected: commit succeeds.

## Task 6: Draft Conclusion, Tighten Abstract, and Normalize Keywords

**Files:**
- Modify: `manuscript/draft.md`

**Interfaces:**
- Consumes: full body draft from Tasks 3-5
- Produces: complete readable first draft

- [ ] **Step 1: Draft the conclusion**

Write 400-600 Chinese characters under “七、结语”. Include:

1. A summary of why engineering competitions are valuable.
2. A statement that the goal is talent cultivation rather than awards.
3. A closing sentence about new engineering practice teaching reform.

- [ ] **Step 2: Revise the abstract**

Revise the abstract to 250-350 Chinese characters. It must contain:

```text
新工科
工程类学科竞赛
四位一体
双循环
赛教融合
```

- [ ] **Step 3: Normalize keywords**

Keep 5-6 keywords. Preferred final list:

```text
工程类学科竞赛；赛教融合；新工科；实践教学；应用型本科高校；创新人才培养
```

- [ ] **Step 4: Verify first draft completeness**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '^# |^## |^### '
(Get-Content -Raw -Encoding UTF8 manuscript\draft.md).Length
```

Expected: the draft has one title, abstract, keywords, seven main sections, and a character length appropriate for a province-journal manuscript.

- [ ] **Step 5: Commit**

Run:

```powershell
git add manuscript/draft.md
git commit -m "docs: complete first draft of competition paper"
```

Expected: commit succeeds.

## Task 7: Literature and Reference Support

**Files:**
- Create: `manuscript/references.md`
- Modify: `manuscript/draft.md`

**Interfaces:**
- Consumes: complete first draft from Task 6
- Produces: citation candidates and in-text citation placement notes

- [ ] **Step 1: Create reference retrieval plan**

Create `manuscript/references.md` with this content:

```markdown
# 参考文献与检索记录

## 1. 检索主题

1. 新工科与工程教育改革
2. 学科竞赛与创新人才培养
3. 赛教融合与实践教学改革
4. 项目式学习与跨专业协同育人
5. 地方应用型本科高校人才培养

## 2. 推荐检索式

| 数据库 | 检索式 |
| --- | --- |
| CNKI | “赛教融合” AND “学科竞赛” |
| CNKI | “工程类学科竞赛” AND “新工科” |
| CNKI | “Robomaster” OR “Robocon” |
| CNKI | “电子设计竞赛” AND “实践教学” |
| Wanfang | “应用型本科” AND “学科竞赛” AND “人才培养” |

## 3. 拟引用位置

| 论文位置 | 需要支撑的观点 | 文献类型 |
| --- | --- | --- |
| 引言 | 新工科强调工程实践能力培养 | 政策文件或工程教育改革论文 |
| 现实价值 | 学科竞赛促进创新能力培养 | 教改论文 |
| 问题分析 | 赛教融合存在课程脱节问题 | 教改论文 |
| 模式构建 | 项目式学习、OBE、实践教学体系 | 理论或教改论文 |
| 实践路径 | 地方应用型本科高校跨专业协同育人 | 教改论文 |

## 4. 文献使用原则

1. 优先使用近五年中文教育教学改革论文。
2. 至少保留 8 篇中文文献，条件允许时增加 1-2 篇政策或工程教育认证相关文献。
3. 不编造作者、题名、期刊、年份和页码。
4. 未核验的文献信息不进入正式参考文献表。
```

- [ ] **Step 2: Retrieve real references**

Use CNKI, Wanfang, official policy sources, or academic search tools to gather real bibliographic entries. Record each verified item under:

```markdown
## 5. 已核验参考文献
```

- [ ] **Step 3: Add citation markers in the draft**

Add citation markers such as `[1]`, `[2]`, `[3]` only after the referenced bibliographic entries are verified.

- [ ] **Step 4: Verify no fake references**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\references.md -Pattern '未核验|示例|作者甲|某某|XXXX'
```

Expected: no unverified sample references appear in the final reference list.

- [ ] **Step 5: Commit**

Run:

```powershell
git add manuscript/references.md manuscript/draft.md
git commit -m "docs: add verified references and citation notes"
```

Expected: commit succeeds.

## Task 8: Revision and Submission Package

**Files:**
- Create: `manuscript/revision-checklist.md`
- Modify: `manuscript/draft.md`
- Optional create: `manuscript/submission-version.docx`

**Interfaces:**
- Consumes: cited full draft from Task 7
- Produces: polished manuscript suitable for journal formatting

- [ ] **Step 1: Create revision checklist**

Create `manuscript/revision-checklist.md` with this content:

```markdown
# 论文修订检查表

## 1. 省刊适配

- [ ] 题目稳妥，能体现教学改革、实践研究和育人模式。
- [ ] 摘要包含背景、问题、方法和结论。
- [ ] 正文结构完整，问题、模式、路径、成效之间逻辑连贯。
- [ ] 语言不过度理论化，符合教育教学类省刊表达习惯。

## 2. 核心升级空间

- [ ] “四位一体、双循环”模型表达清楚。
- [ ] 引言和模式部分适度引入新工科、OBE 或项目式学习。
- [ ] 参考文献能支撑主要观点。
- [ ] 如有真实数据，已补入参赛人数、项目数量、获奖情况或课程改革成果。

## 3. 数据与伦理风险

- [ ] 未写入未经确认的具体数字。
- [ ] 案例单位保持匿名化。
- [ ] 没有夸大竞赛成果。
- [ ] 没有将论文写成赛事宣传稿。

## 4. 语言与格式

- [ ] 一级、二级标题层次统一。
- [ ] 关键词数量为 5-6 个。
- [ ] 参考文献格式统一。
- [ ] 摘要长度适合投稿要求。
```

- [ ] **Step 2: Perform coherence revision**

Read `manuscript/draft.md` from start to finish and revise transitions so every section follows this chain:

```text
背景需求 -> 现实价值 -> 存在问题 -> 模式构建 -> 实践路径 -> 成效反思 -> 结语
```

- [ ] **Step 3: Perform language polish**

Replace casual or overgeneral expressions with academic phrasing. Examples:

```text
很重要 -> 具有重要意义
学生能力变强 -> 学生工程实践能力得到提升
效果不错 -> 取得一定实践成效
比赛成绩 -> 竞赛育人成效
```

- [ ] **Step 4: Verify final manuscript**

Run:

```powershell
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '很|不错|大量|很多|非常|显著提高|极大促进'
Select-String -Encoding UTF8 -Path manuscript\draft.md -Pattern '^## |^### '
```

Expected: informal or exaggerated words are reviewed and either justified or replaced.

- [ ] **Step 5: Optional Word export**

If the user requests a `.docx`, convert the stable Markdown draft into `manuscript/submission-version.docx` and visually check title, abstract, headings, paragraphs, and references.

- [ ] **Step 6: Commit**

Run:

```powershell
git add manuscript/revision-checklist.md manuscript/draft.md
git commit -m "docs: polish competition paper for submission"
```

Expected: commit succeeds.

## Self-Review

### Spec Coverage

- Approved topic and title: covered by Task 2 and Task 3.
- Province-journal positioning with core-journal upgrade space: covered by Global Constraints and Task 8.
- Anonymous local applied undergraduate case: covered by Task 1 and Task 5.
- “Four-in-one, dual-cycle” model: covered by Task 4.
- Robomaster, Robocon, and electronic design contest examples: covered by Task 1 and Task 5.
- Experience-summary writing without invented data: covered by Task 1, Task 5, and Task 8.
- References and citation risk control: covered by Task 7.

### Placeholder Scan

The plan contains no “TBD”, “TODO”, “implement later”, fake references, or fabricated numeric results. Unknown data are handled through explicit non-invention rules and optional later supplementation.

### Consistency Check

All tasks use the same core model terms: “四位一体”, “双循环”, “课程教学”, “项目训练”, “竞赛实践”, and “成果转化”. File paths are consistent across tasks.
