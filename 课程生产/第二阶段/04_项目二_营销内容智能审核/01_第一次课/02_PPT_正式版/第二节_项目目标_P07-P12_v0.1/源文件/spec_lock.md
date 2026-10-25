<!-- ppt-master-schema: spec-lock/v1 -->
# Execution Lock

## canvas
- format: PPT 16:9
- viewBox: 0 0 1280 720

## communication
- primary_language: zh-CN
- audience: 已学过审核基础、正在学习AI产品项目落地的学员
- objective: 承接第一节痛点，明确三个目标、首版边界、人工责任和提效判断，形成下一节业务方案的项目定义。
- core_message: 首版减少重复查找、沟通和复核；人保留正式决定与处置，用质量、总人工投入和处理结果验证。
- consumption_mode: presentation

## mode
- mode: custom
- mode_behavior: 以章节标题页定位，按目标、范围责任、判断、讨论、项目定义推进。

## visual_style
- visual_style: custom
- visual_style_behavior: 蓝色章节页形成停顿；正文沿用浅底、深蓝标题、蓝色编号和细分隔线，按关系组织内容。

## colors
- background: #F7F9FC
- secondary_bg: #EAF0FB
- primary: #2457D6
- accent: #EAA45B
- secondary_accent: #D6E2FA
- body_text: #172C4D
- muted: #61718A
- grid: #D8E0EC
- surface: #FFFFFF

## typography
- font_family: Microsoft YaHei, Arial, PingFang SC, sans-serif
- title_family: Microsoft YaHei, Arial, PingFang SC, sans-serif
- body_family: Microsoft YaHei, Arial, PingFang SC, sans-serif
- body: 28
- title: 48
- subtitle: 32
- annotation: 20
- label: 24
- footer: 14
- display: 64

## icons
- library: none
- inventory: none

## page_rhythm
- P07: anchor
- P08: dense
- P09: dense
- P10: dense
- P11: breathing
- P12: dense

## pptx_structure
- mode: flat

## forbidden
- `mask`, `<style>`, `class`, external CSS, `<foreignObject>`, `textPath`, `@font-face`, `<animate*>`, `<set>`, `<script>` / event attributes, `<iframe>`
- HTML named entities in text; write typography as raw Unicode and escape XML reserved characters
- 缺少本节及后续各节开头的章节标题页
- 把教学期望改善说成实测效果，或无依据承诺节省比例
- 首版自动批准、自动发布，或把系统无提示等同整篇审核完成
- 提前生成第三节及之后的页面
