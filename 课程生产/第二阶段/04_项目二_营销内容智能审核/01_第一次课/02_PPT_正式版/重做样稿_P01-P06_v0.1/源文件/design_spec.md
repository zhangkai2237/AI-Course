<!-- ppt-master-schema: design-spec/v1 -->
# Lesson 1 Rebuild Opening Sample - Design Spec

## I. Project Information

| Item | Value |
| --- | --- |
| Project Name | 第一次课重做样稿 开场与压缩后的第一节 |
| Canvas Format | ppt169 / PPT 16:9 / 1280×720 |
| Page Count | 6 |
| Primary Language | zh-CN |
| Target Audience | 已学过智能审核基础理论、正在学习AI产品项目落地的课程学员。 |
| Communication Intent | 让学员明确项目定位与七节学习顺序，再用四页说明企业原有审核方式、现有流程、人工核对动作和具体痛点。 |
| Desired Audience Outcome | 能按讲义七节定位课程，说明P1审核的角色流程、原句与依据差异、三类重复劳动和后续复查需要。 |
| Core Message / Ask / Action | 以P1常规营销内容为案例，先完成审核提效方案；七节按业务、目标、流程、检查要求、分工、产品与验证推进。 |
| Delivery Context | 现场投影授课；本次范围为开场两页与第一节四页，原大纲P03—P10压缩为P03—P06。完整课程180分钟含两次10分钟休息。 |
| Artifact Afterlife | 六页可编辑PPTX和逐页PNG预览，用于审阅讲义对应及第一节压缩；同步AI-Course，不改旧版整套PPT。 |
| Reading Mode | presentation |
| Content Strategy | 以完整版讲义开场、目录和第一节为依据；以企业原来怎样审核、流程怎样走、痛点在哪里为主线；保留价格例子解释核对动作，商品、赠品、配送等详细资料留在讲义与后续审核点章节。不额外增加独立章节导入和小结页。 |
| Design Style | 课堂主题封面与编号学习路径；自由设计。 |
| AI Image Acquisition Path | not applicable — 复用现有教学情境图，不新增生成 |
| Generation Mode | continuous |
| Spec Refinement | disabled |
| Speaker Notes | enabled — 延续用户授权自主决定，提供开场讲解与过渡 |
| Custom Animations | disabled — 主动策略，使用静态可编辑页面 |
| Narration Audio | disabled — 主动策略，教师现场讲解 |
| Created Date | 2026-10-18 |

User delegation remains “方案你来定，直接生成”; the latest steering requests first-section compression after the first two pages. Scope is six slides, including a four-slide first section. Latest user steering prioritizes the original enterprise review method, its process, and concrete pain points. Free design and communication contract are retained in chat; no UI receipts are synthesized. Three directions compared: (1) clear course title plus a numbered seven-step learning path, (2) full-scene product image and narrative case entrance, (3) dark engineering blueprint and module map. Direction 1 is selected because the user's priority is correspondence to the lecture; direction 2's single supporting scene is retained on the cover only. No template workspace is selected.

## II. Canvas Specification

| Property | Value |
| --- | --- |
| Format | PPT 16:9 |
| Dimensions | 1280 × 720 |
| viewBox | `0 0 1280 720` |
| Margins | 48–56px horizontal; safe footer region |
| Content Area | 全页开放排布，正文和页脚分开 |

## III. Visual Theme

### Theme Style

- **Mode**: custom
- **Mode Behavior**: 封面先明确课程主题与提效定位，目录使用一至七的连续阅读路径，章节名称和提问逐项对应讲义。
- **Visual style**: custom
- **Visual Style Behavior**: 白底、深蓝标题、蓝色重点和细线构成清楚的教学层级；封面主标题配单张商品情境图，目录以编号路径与留白承载结构，不用等权卡片堆叠。
- **Theme**: Lecture-aligned opening
- **Tone**: 清楚、克制、便于投影与逐节定位

### Color Scheme

| Role | HEX | Purpose |
| --- | --- | --- |
| Background | #F7F9FC | 页面背景 |
| Secondary background | #EAF0FB | 提效定位与流程背景 |
| Primary | #2457D6 | 编号、重点与流程 |
| Accent | #EAA45B | 稀疏强调 |
| Secondary accent | #D6E2FA | 辅助区域 |
| Body text | #172C4D | 主标题与正文 |
| Muted | #61718A | 注释与页脚 |
| Grid | #D8E0EC | 路径与分隔 |
| Surface | #FFFFFF | 图片衬底与开放页面 |

## IV. Typography System

### Font Plan

| Role | Character (Reference) | Primary | English if non-English | Fallback tail |
| --- | --- | --- | --- | --- |
| Title | 粗体无衬线 | Microsoft YaHei | Arial | PingFang SC, sans-serif |
| Body | 清晰无衬线 | Microsoft YaHei | Arial | PingFang SC, sans-serif |

- **Title stack**: Microsoft YaHei, Arial, PingFang SC, sans-serif
- **Body stack**: Microsoft YaHei, Arial, PingFang SC, sans-serif

### Font Size Hierarchy

| Purpose | Anchor Size (px) |
| --- | ---: |
| Body | 28 |
| Title | 48 |
| Subtitle | 32 |
| Annotation | 20 |
| Label | 24 |
| Footer | 14 |
| Display | 64 |

## V. Layout Principles

### Deck-wide Direction

- **Hierarchy direction**: 课程主题或目录标题先进入视线，再读定位与章节内容；辅助说明退后。
- **Composition tendency**: 封面采用主标题与情境图并置、下方三环节横向路径；目录采用连续编号的纵向阅读路径，每一项包含讲义节名与本节问题。
- **Cross-page continuity**: 同一蓝色编号、深蓝文字、细分隔线和页码位置；主题封面过渡到讲义七节导航。
- **Spacing posture**: 封面开放，目录适度紧凑，确保投影时标题与章节名优先可读。

## VI. Icon Usage Specification

- **Primary bundled library**: none

| Icon Path | Suitable Scenarios |
| --- | --- |

## VIII. Image Resource List

| Filename | Dimensions | Ratio | Purpose | Type | Layout pattern | Crop Policy | Acquire Via | Status | Reference | text_policy | page_role |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| case-home.png | 1536×1024 | 1.5 | 让P1养宠家庭营销文案有具体商品与使用情境 | Scene | 完整展示养宠家庭与扫拖机，放在封面主题侧边作为案例锚点 | no-crop | user | Existing | 复用本对话已有AI生成教学情境图，不代表真实产品或业务照片 | none | local |

## IX. Content Outline

### Part 1: 开场两页

#### Slide 01 - 电商营销内容审核提效项目

- **Audience move**: 从泛化的AI课程进入明确的常规营销审核项目，知道首版定位和覆盖环节。
- **Layout**: 大号课程主题与单张情境图并置；首版目标独立强调，下方横向呈现三个审核环节。
- **Title**: 电商营销内容审核提效项目
- **Core message**: 首版以审核提效为目标，覆盖自检、正式审核与巡检，正式决定和处置由人承担。
- **Content**: AI产品项目实战 / 第一次课；电商营销内容 / 审核提效项目；首版定位：审核提效；案例：净小伴P1标准版 · 常规营销文案审核；从业务背景，走到产品方案与落地验证；首版覆盖事前自检、事中审核、事后巡检，正式决定与处置由人承担；案例品牌、资料与配图为教学模拟。
- **Images**: case-home.png，保留完整原生比例，精确信息写在图外。
- **Cover impact**: 同一篇P1营销文案贯穿三个审核环节，主标题与横向流程共同建立项目范围。

#### Slide 02 - 本次课的七节内容

- **Audience move**: 从知道课程主题到能按讲义定位七节内容，并理解每节解决的问题。
- **Layout**: 七个连续编号节点构成纵向学习路径，左侧为讲义节名，旁侧为该节问题；整体按一至七顺序阅读。
- **Title**: 本次课的七节内容
- **Core message**: 严格沿完整版讲义的七节与项目推进顺序学习。
- **Content**: 01业务背景：原来的审核怎么做，哪些动作在重复？02项目目标：第一版要让谁少做什么，范围到哪里？03业务方案：自检、正式审核、巡检怎样接成闭环？04审核体系：一条业务要求，怎样定义成具体审核点？05方案选择：规则、模型、人工怎样分工与交接？06系统与产品方案：流程怎样落实到架构、页面、版本与状态？07落地与验证：先做什么，用哪些证据判断是否提效？180分钟，含两次各10分钟休息。章节名称与完整版讲义逐节对应。

### Part 2: 第一节 企业原有审核方式、流程与痛点 / 四页

#### Slide 03 - 企业原来如何审核
- **Audience move**: 明确既有人工审核的组织方式和事实确认责任。
- **Layout**: 先说明人工阅读与资料对照，再区分作者、审核员和业务资料负责人，下方保留内容使用责任。
- **Title**: 企业原来是怎样审核营销内容的？
- **Core message**: 日常审核以人工阅读、对照资料和责任人确认为基础。
- **Content**: 第一节·业务背景 / 1.1 原有审核方式。小周按日常计划撰写P1推广；作者负责写稿、自查、修改；小林负责读文案、找依据、核对、反馈与复核；产品负责人确认商品事实，活动负责人确认价格赠品，履约负责人确认配送；内容负责人维护实际使用版本和位置。业务事实不能由审核员凭常识补齐。

#### Slide 04 - 现有审核流程
- **Audience move**: 能沿同一篇文案说明提交到定稿使用的动作与修改回路。
- **Layout**: 七个有序动作形成主流程，下面标出需要修改的一轮及复核未解决时继续处理的关系。
- **Title**: 一篇文案，怎样走完现有审核流程？
- **Core message**: 反馈、修改与复核是正式审核的一部分，定稿后还要关联实际使用。
- **Content**: 第一节·业务背景 / 1.2 现有审核流程。作者撰写自查→提交正文与资料→小林对照核对→给出反馈意见→作者修改重提→逐项复核→定稿并按计划使用。图示需要修改的一轮；没有待处理事项时可按既定权限定稿，未解决的问题继续修改与复核。通过结论对应明确版本，内容负责人维护实际使用。

#### Slide 05 - 走一次人工核对
- **Audience move**: 用一个价格例子理解审核员的实际劳动，不把第一节变成完整审核点训练。
- **Layout**: 原句与活动依据并置，下方展开读原句、找资料、核对条件、写反馈四个动作，最后给可执行意见。
- **Title**: 跟着小林，走一次人工核对
- **Core message**: 审核员要把原句与当前适用条件对应起来，再解释清楚怎样改。
- **Content**: 第一节·业务背景 / 1.3 人工核对动作；教学模拟。原句“本次活动到手1799元，人人同价！”；F-PROMO-01 v1：本次活动价2399元，领取并使用400元活动券后1999元，新会员再领取并使用200元专享券后1799元，仅本次活动适用期内有效。读原句→找到当前活动资料→核对价格、资格及两类券→返回正文写反馈。需要补充新会员资格及两类券领取与使用条件，不能扩大为人人同价。这里不代表其余审核项或整篇已经完成。
- **Data class**: scenario — 金额与条件为教学给定。

#### Slide 06 - 从流程归纳痛点
- **Audience move**: 从可见工作中形成具体问题清单，再进入项目目标与首版范围。
- **Layout**: 四个实际片段对应四类劳动，保留价格往返对话和后续权益变化事件，页尾回收到下一节问题。
- **Title**: 审核提效需求，来自哪些具体痛点？
- **Core message**: 重复查找、解释往返、版本复核与后续内容维护都应计入人工工作量。
- **Content**: 第一节·业务背景 / 1.4 痛点归纳。查找资料：正文与多份资料反复切换，正常句也要核实。解释往返：“价格补条件”→作者改“领券1799”→仍缺新会员和两类券条件。版本复核：已修改不等于已解决，重连旧意见、新原句、当前依据，其他改动同样核对。后续维护：活动负责人确认赠品名额用完后，找相关活动、在用内容、维护人，停用与仅功能内容不能一律修改。需求来自具体动作，需记录双方总投入；下一节明确希望减少哪些劳动及首版范围。

## X. Speaker Notes Requirements

- **Generation**: enabled
- **Filename**: match each SVG filename under notes/
- **Content**: 逐页说明可见内容并自然衔接讲义。开场合计2分钟；P03原有方式3分钟；P04审核流程4分钟；P05人工核对5分钟；P06痛点6分钟含讨论。对比与答案均在材料之后揭示，资料细则可放备注，不将模拟事实说成实测。
- **Total duration**: 本次六页合计20分钟，为第一次课开场与第一节；完整课仍180分钟。
- **Notes style**: conversational
- **Presentation purpose**: instruct and orient
