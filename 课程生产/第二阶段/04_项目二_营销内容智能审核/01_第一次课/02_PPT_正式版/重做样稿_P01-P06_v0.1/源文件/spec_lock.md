<!-- ppt-master-schema: spec-lock/v1 -->
# Execution Lock

## canvas
- format: PPT 16:9
- viewBox: 0 0 1280 720

## communication
- primary_language: zh-CN
- audience: 已学过审核基础理论、正在学习AI产品项目落地的学员
- objective: 用开场和四页第一节明确项目定位与七节路径，从P1日常业务中说明企业原有审核方式、流程及具体痛点。
- core_message: 同一篇P1常规营销文案贯穿自检、审核与巡检，按讲义七节推进到可验证的落地方案。
- consumption_mode: presentation

## mode
- mode: custom
- mode_behavior: 封面明确课程主题与首版定位，目录以连续编号列出七节原有名称及问题，建立可对照讲义的学习路径。

## visual_style
- visual_style: custom
- visual_style_behavior: 白底、深蓝标题、蓝色编号和细线形成清楚层级；封面主标题配单张情境图，目录用纵向编号路径和留白承载结构。

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

## images
- case-home: images/case-home.png | source=user | pattern=完整展示养宠家庭与扫拖机，放在封面主题侧边作为案例锚点 | crop=no-crop

## page_rhythm
- P01: anchor
- P02: dense
- P03: dense
- P04: dense
- P05: anchor
- P06: dense

## pptx_structure
- mode: flat

## forbidden
- `mask`, `<style>`, `class`, external CSS, `<foreignObject>`, `textPath`, `@font-face`, `<animate*>`, `<set>`, `<script>` / event attributes, `<iframe>`
- HTML named entities in text; write typography as raw Unicode and escape XML reserved characters
- 将情境图说成真实产品或实测案例
- 将三环节提效说成首版自动批准或自动发布
- 生成本次要求以外的后续页面
