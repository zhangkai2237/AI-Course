<!-- ppt-master-schema: design-spec/v1 -->
# Lesson 1 Section 2 - Design Spec

## I. Project Information

| Item | Value |
| --- | --- |
| Project Name | 第一次课 第二节 项目目标 |
| Canvas Format | ppt169 / PPT 16:9 / 1280×720 |
| Page Count | 6 |
| Primary Language | zh-CN |
| Target Audience | 已学过智能审核基础、正在学习AI产品项目落地的学员。 |
| Communication Intent | 承接第一节的人工审核痛点，明确首版希望改变哪些工作、覆盖什么、由谁负责及怎样判断提效。 |
| Desired Audience Outcome | 能说清三个目标，限定首版范围与人工责任，区分输出速度和整体业务价值，并形成首版项目定义。 |
| Core Message / Ask / Action | 首版以减少重复查找、沟通和复核为目标，人保留正式决定与内容处置；用质量、总人工投入和处理结果验证。 |
| Delivery Context | 现场投影授课；第二节15分钟；全课连续页码P07—P12，单独交付本节6页。 |
| Artifact Afterlife | 可编辑PPTX、逐页预览与教师备注；同步AI-Course课程项目。 |
| Reading Mode | presentation |
| Content Strategy | 依据完整版讲义第二节正文，不提前展开第三节操作流程；精简投影文字，解释留在教师备注。 |
| Design Style | 延续已确认的浅底、深蓝标题、蓝色编号与细分隔线；每节开头有清楚的章节标题页。 |
| AI Image Acquisition Path | not applicable — 本节用可编辑关系图与文字承载教学内容，不新增位图 |
| Generation Mode | continuous |
| Spec Refinement | disabled |
| Speaker Notes | enabled — 延续此前授权的逐页教师讲解 |
| Custom Animations | disabled — 现场静态授课 |
| Narration Audio | disabled — 由教师现场讲解 |
| Created Date | 2026-10-18 |

沿用用户已确认的视觉风格和“方案你来定，直接生成”的授权。新增明确要求：从本节起，后续每节均以章节标题页开场。比较的方向为痛点对应的工作目标、按功能模块罗列、以量化指标开场；选择工作目标方向，因为讲义要求先定义业务变化，而且没有实测基线。首版成果不是PRD。本节不修改第一节已确认的六页。

## II. Canvas Specification

| Property | Value |
| --- | --- |
| Format | PPT 16:9 |
| Dimensions | 1280 × 720 |
| viewBox | `0 0 1280 720` |
| Margins | 56px horizontal; 40px top; footer separated |
| Content Area | 主体区域与页脚分开，投影阅读优先 |

## III. Visual Theme

### Theme Style
- **Mode**: custom
- **Mode Behavior**: 章节标题先定位，随后按目标、范围责任、判断、讨论、项目定义推进。
- **Visual style**: custom
- **Visual Style Behavior**: 章节页使用蓝色整面形成停顿；正文延续浅底与深蓝标题，用编号、真实前后关系、范围边界与阅读路径组织内容。
- **Theme**: Lecture-aligned project goals
- **Tone**: 清晰、具体、便于课堂讲述

### Color Scheme
| Role | HEX | Purpose |
| --- | --- | --- |
| Background | #F7F9FC | 正文页底色 |
| Secondary background | #EAF0FB | 分组背景 |
| Primary | #2457D6 | 标题页、编号与重点 |
| Accent | #EAA45B | 讨论提示 |
| Secondary accent | #D6E2FA | 边界和辅助强调 |
| Body text | #172C4D | 正文与标题 |
| Muted | #61718A | 辅助说明 |
| Grid | #D8E0EC | 分隔与连接 |
| Surface | #FFFFFF | 深色底上的文字与必要面板 |

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
- **Hierarchy direction**: 一页一个主问题，先读结论或课堂任务，再看对应依据与条件。
- **Composition tendency**: 标题页形成章节停顿；目标页沿三条工作变化组织；范围页区分系统辅助与人工责任；验证页形成连续判断；讨论页保留留白；定义页合成一份项目说明。
- **Cross-page continuity**: 相同字体、蓝色编号、标题层级、连续页码；每节首部有标题页。
- **Spacing posture**: 章节与讨论页留白，范围和定义页适度紧凑。

## VI. Icon Usage Specification
- **Primary bundled library**: none

| Icon Path | Suitable Scenarios |
| --- | --- |

## VIII. Image Resource List

| Filename | Dimensions | Ratio | Purpose | Type | Layout pattern | Crop Policy | Acquire Via | Status | Reference | text_policy | page_role |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |

## IX. Content Outline

### Part 1: 第二节 项目目标 / P07—P12

#### Slide 07 - 第二节 项目目标
- **Audience move**: 从已知人工审核痛点，进入目标与首版范围的定义。
- **Layout**: 蓝色章节页，清楚显示第二节与主问题，底部用三段承接关系定位本节。
- **Title**: 项目目标
- **Core message**: 第一版到底要让谁少做什么？
- **Content**: 第一次课 / 第二节；02 项目目标；第一版，到底要让谁少做什么？本节明确三个目标、首版范围、验证方向。承接：业务背景→项目目标→业务方案；本节约15分钟。页码07。

#### Slide 08 - 把智能审核变成三个具体目标
- **Audience move**: 从抽象功能愿望，转成具体的工作改善。
- **Layout**: 三条并列的工作变化，每条先给目标，再用原有劳动和期望变化解释；结尾提醒需要验证。
- **Title**: 把“智能审核”变成三个具体目标
- **Core message**: 先明确谁少做什么，再考虑功能。
- **Content**: 第二节 / 2.1 工作目标。业务说“做个智能审核”，还要问谁的工作怎样改善。目标一：依据来到原句旁，减少适用资料的查找与定位；小林看1799元时，能看到本次活动的价格、资格和用券条件，仍需确认依据是否正确适用。目标二：意见具体，修改与复核接得上，减少作者和审核员的反复解释与重新拼接；补齐新会员与两类券条件，新稿能对应旧意见并看见其他改动。目标三：找对需要复查的在用内容，减少范围定位与找维护人的劳动；赠品条件变化后，按本次活动、使用状态与内容表达确认范围，交负责人处理。以上均是期望改善，尚需验证。

#### Slide 09 - 第一版做到哪里 谁来负责
- **Audience move**: 从目标转成可实施的试用边界和责任分工。
- **Layout**: 左侧首版边界，右侧人工责任；底部用一条贯通说明资料与关联前提，及首版内部建设顺序。
- **Title**: 第一版做到哪里，谁来负责？
- **Core message**: 小范围辅助、三环节覆盖，正式决定与处置由人承担。
- **Content**: 第二节 / 2.2 首版范围与责任。先选P1及少量资料齐全商品；每份文本绑定单SKU与活动，核对功能事实、价格赠品条件及配送依据；覆盖自检、正式审核、负责人选范围发起巡检。系统辅助定位问题、展示依据、衔接修改复核与处理。人工保留正式审核决定，以及内容负责人确认修改、暂停等处置；缺资料、执行失败、争议都交责任人确认。前提：资料经对应业务负责人确认，文案关联商品、活动、版本、使用状态；不齐先补或缩小范围。建设顺序：先做深正式审核，再接自检与巡检，三者都属首版。未发现已覆盖项问题，不等于完成全部审核。

#### Slide 10 - 怎样证明审核真的更省力
- **Audience move**: 从结果速度转到质量、总劳动与处理结果的业务判断。
- **Layout**: 三个连续检查视角，沿阅读路径展开；总人工投入以双方工作构成表达，等待单列；底部说明无基线时怎样开始。
- **Title**: 怎样证明审核真的更省力？
- **Core message**: 同时看问题质量、总人工投入与实际处理结果。
- **Content**: 第二节 / 2.3 提效判断。先看问题与依据：真实问题有没有漏、正常内容有没有误报、引用资料是否适用。再看总投入：作者自检、读反馈、修改与审核员确认、纠错、复核合起来算，等待业务回复另记。最后看处理：文案是否改对、复核是否确认、巡检内容是否实际更新。工具出结果快只是一个局部。先记录原流程，再用可比任务做小范围对照；没有业务基线，不承诺固定节省比例或通用准确率门槛。

#### Slide 11 - 课堂判断 这算提效了吗
- **Audience move**: 能识别局部变快和工作前移，主动提出缺失证据。
- **Layout**: 一段真实工作情境加一个大问题，下面保留三个讨论提示，不在主视觉直接给出结论；教师备注提供完整答案。
- **Title**: 这就能宣布“项目提效”了吗？
- **Core message**: 自检投入增加、退回减少、复核变快，仍需进一步判断。
- **Content**: 第二节 / 课堂讨论·约1分钟。接入自检后：小周提交前多花了时间；正式退回减少；小林复核更快。现在能直接宣布提效吗？还缺哪些证据？提示：两个人总共花了多少时间；等待是否变化；审核质量和任务可比性是否确认。以上是教学情境，没有量化实测数据。

#### Slide 12 - 收成一份首版项目定义
- **Audience move**: 能用完整项目定义衔接第三节具体业务方案。
- **Layout**: 一个项目说明主体分层呈现对象、范围、目标、责任与验证，不变成密集规格表；页尾明确下一节要回答的流程问题。
- **Title**: 把目标收成一份首版项目定义
- **Core message**: 在资料齐全的小范围中辅助审核，用真实工作效果验证价值。
- **Content**: 第二节 / 本节产物。对象：内容作者、审核员、相关业务负责人。场景：单SKU营销文案的自检、正式审核与历史内容巡检。目标：减少重复查找、沟通与复核。辅助方式：依据已确认且适用的商品、活动、配送资料，定位问题、展示依据、衔接修改复核与后续处理。责任：人作正式审核与内容处置决定。试用：少量资料齐全商品，用问题质量、总人工投入和处理效果验证。演进：后续另设覆盖、资料有效性、质量证据、权限与人工兜底条件，再评估限定场景自动审核。下一节：按此定义走通三环节，谁发起、带什么、看到什么、谁接手。

## X. Speaker Notes Requirements
- **Generation**: enabled
- **Filename**: match each SVG filename under notes/
- **Content**: 依据最终SVG生成完整中文口语讲解，补充讲义细节，讨论页保留完整参考答案；不声称已经实测；不把系统无提示当审核通过。
- **Total duration**: 15分钟；章节页约0.5分钟，目标3分钟，范围3分钟，验证3分钟，讨论2分钟，项目定义与过渡3.5分钟。
- **Notes style**: interactive
- **Presentation purpose**: instruct and define project
