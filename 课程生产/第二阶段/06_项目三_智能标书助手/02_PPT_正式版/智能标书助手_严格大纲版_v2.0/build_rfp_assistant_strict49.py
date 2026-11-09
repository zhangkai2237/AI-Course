#!/usr/bin/env python3
"""智能标书助手：严格按照已确认49页大纲生成的正式授课版。"""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
SKILL = Path('/Users/keivn/.codex/skills/slide-maker/scripts')
sys.path.insert(0, str(SKILL))

from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from deckkit import (
    blank_deck, add_slide, box, text, picture, arrow, speaker_notes,
    design_intent, tag_motif, lint_layout, declare_delivery,
    set_palette, set_geometry, set_ground, _flat,
)

OUT = ROOT / '智能标书助手_严格按照49页大纲_正式授课PPT_v2.0.pptx'
ASSETS = ROOT / 'assets'

BG = RGBColor(244, 246, 248)
PAPER = RGBColor(255, 255, 255)
INK = RGBColor(24, 36, 51)
GREY = RGBColor(91, 106, 121)
MUTE = RGBColor(137, 150, 162)
LINE = RGBColor(218, 225, 231)
DEEP = RGBColor(13, 27, 42)
DEEP2 = RGBColor(21, 43, 63)
TEAL = RGBColor(19, 138, 126)
TEAL_SOFT = RGBColor(227, 243, 240)
BLUE = RGBColor(62, 111, 176)
BLUE_SOFT = RGBColor(233, 240, 250)
AMBER = RGBColor(201, 133, 23)
AMBER_SOFT = RGBColor(255, 243, 217)
CORAL = RGBColor(210, 90, 74)
CORAL_SOFT = RGBColor(253, 233, 230)
GREEN = RGBColor(43, 144, 100)
FONT = 'Noto Sans CJK SC'
EA = 'Noto Sans CJK SC'

W, H = 10.0, 5.625
M = 0.60

SECTIONS = {
    1: '开场', 2: '开场',
    3: '01 为什么需要智能标书系统', 4: '01 为什么需要智能标书系统', 5: '01 为什么需要智能标书系统',
    6: '01 为什么需要智能标书系统', 7: '01 为什么需要智能标书系统', 8: '01 为什么需要智能标书系统',
    9: '01 为什么需要智能标书系统', 10: '01 为什么需要智能标书系统',
    11: '02 RFP需求拆解', 12: '02 RFP需求拆解', 13: '02 RFP需求拆解', 14: '02 RFP需求拆解',
    15: '02 RFP需求拆解', 16: '02 RFP需求拆解', 17: '02 RFP需求拆解', 18: '02 RFP需求拆解',
    19: '03 响应矩阵与参标判断', 20: '03 响应矩阵与参标判断', 21: '03 响应矩阵与参标判断',
    22: '03 响应矩阵与参标判断', 23: '03 响应矩阵与参标判断', 24: '03 响应矩阵与参标判断',
    25: '03 响应矩阵与参标判断', 26: '03 响应矩阵与参标判断',
    27: '04 受控长文档生成', 28: '04 受控长文档生成', 29: '04 受控长文档生成', 30: '04 受控长文档生成',
    31: '04 受控长文档生成', 32: '04 受控长文档生成', 33: '04 受控长文档生成', 34: '04 受控长文档生成',
    35: '05 一致性与风险检查', 36: '05 一致性与风险检查', 37: '05 一致性与风险检查', 38: '05 一致性与风险检查',
    39: '05 一致性与风险检查', 40: '05 一致性与风险检查', 41: '05 一致性与风险检查',
    42: '06 MVP产品设计', 43: '06 MVP产品设计', 44: '06 MVP产品设计', 45: '06 MVP产品设计', 46: '06 MVP产品设计',
    47: '07 场景迁移与作业', 48: '07 场景迁移与作业', 49: '07 场景迁移与作业',
}

TITLES = {
    2: '核心链路：从RFP到可提交标书',
    4: '案例背景',
    5: '招标文件关键要求',
    6: '启明数科已有企业材料与事实',
    7: 'AI生成的标书初稿',
    8: '课堂互动：找出这段初稿的问题',
    9: '参考答案：问题出在哪里',
    10: '通用AI写作工具 vs 智能标书系统',
    12: 'RFP的典型结构',
    13: '八类要求抽取框架',
    14: '摘要 vs 要求清单',
    15: '复合要求拆分',
    16: '要求卡片字段',
    17: '课堂练习：从原文中抽取要求',
    18: '参考答案：六项独立要求',
    20: '响应矩阵的作用',
    21: '五种响应状态',
    22: '资格与技术响应矩阵',
    23: '交付与服务响应矩阵',
    24: '参标初步判断',
    25: '课堂练习：判断响应状态',
    26: '参考答案',
    28: '一次生成整篇标书的问题',
    29: '两级标书树',
    30: '评分标准决定目录深度',
    31: '要求与章节映射',
    32: '章节生成任务包',
    33: '企业材料使用规则',
    34: '受控生成示例',
    36: '七类检查',
    37: '跨章节一致性',
    38: '过度承诺检查',
    39: '课堂练习：找出方案中的风险',
    40: '规则、AI与人工的分工',
    41: '问题处理流程',
    43: 'MVP产品流程',
    44: '页面1–2：项目主页与要求解析',
    45: '页面3–5：响应矩阵、章节编辑、风险检查与导出',
    46: '课堂MVP验收标准',
    48: '三个迁移方向',
    49: '个人项目V0.1作业与结课',
}

REQS = [
    ('R-001', '资格', '有效期内的信息安全认证', '证书已过期', '存在差异'),
    ('R-002', '资格', '近三年3项政务同类案例', '2项政务 + 1项零售', '部分满足'),
    ('R-003', '技术', '私有部署且并发≥5,000', '支持私有部署；实测10,000', '已满足'),
    ('R-004', '交付', '60天内完成初验', '标准周期90天', '存在差异'),
    ('R-005', '技术', '国产OS与数据库兼容', '已确认三类国产环境', '已满足'),
    ('R-006', '服务', '7×24服务可得3分', '标准服务5×8', '待确认'),
]


def rgb(value):
    return value if isinstance(value, RGBColor) else RGBColor(*value)


def tx(slide, x, y, w, h, value, size=15, color=INK, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, italic=False):
    runs = [[(line, size, rgb(color), bold, italic)] for line in str(value).split('\n')]
    return text(slide, x, y, w, h, runs, align=align, anchor=valign,
                space_after=2, line_spacing=1.0)


def rule(slide, x1, y1, x2, y2, color=LINE, width=1.0, dash=False):
    shp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = rgb(color)
    shp.line.width = Pt(width)
    if dash:
        shp.line.dash_style = 2
    return _flat(shp)


def circle(slide, x, y, d, fill=TEAL, line_c=None, line_w=0.8):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    if line_c:
        shp.line.color.rgb = rgb(line_c); shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    return _flat(shp)


def rounded(slide, x, y, w, h, fill=PAPER, line_c=LINE, radius=True, lw=1.0):
    return box(slide, x, y, w, h, fill=rgb(fill), line=rgb(line_c) if line_c else None, line_w=lw, round=radius)


def icon(slide, name, x, y, size=0.38):
    path = ASSETS / f'{name}.png'
    if path.exists():
        return picture(slide, str(path), x, y, size, size, fit='contain')


def note(slide, page, body):
    speaker_notes(slide, f'第{page}页讲解建议：\n{body}\n\n说明：案例中的公司、招标项目与数据均为课堂模拟。')


def interior_motif(slide, dark=False):
    c = RGBColor(69, 91, 110) if dark else LINE
    rule(slide, 8.50, 0.43, 9.38, 0.43, c, 1.1)
    for i, (label, fill) in enumerate([('R', TEAL), ('M', BLUE), ('C', AMBER)]):
        x = 8.48 + i * 0.42
        circle(slide, x, 0.33, 0.20, fill)
        tx(slide, x, 0.365, 0.20, 0.12, label, 5.7, PAPER, True, align=PP_ALIGN.CENTER)


def chrome(slide, page, title=None, dark=False, title_style=0, show_section=True):
    box(slide, 0, 0, W, H, fill=DEEP if dark else BG)
    interior_motif(slide, dark)
    base = RGBColor(174, 190, 202) if dark else MUTE
    if title_style % 3 == 0:
        tx(slide, M, 0.33, 4.8, 0.22, SECTIONS.get(page, ''), 9, base, True)
        if title:
            tx(slide, M, 0.70, 8.68, 0.75, title, 24.5, PAPER if dark else INK, True)
            rule(slide, M, 1.42, 1.42, 1.42, TEAL, 3.0)
    elif title_style % 3 == 1:
        rounded(slide, M, 0.31, 2.45, 0.30, TEAL if dark else TEAL_SOFT, None)
        tx(slide, M+0.12, 0.38, 2.2, 0.15, SECTIONS.get(page, ''), 7.8, PAPER if dark else TEAL, True)
        if title:
            tx(slide, M, 0.82, 8.70, 0.65, title, 24.5, PAPER if dark else INK, True)
    else:
        tx(slide, M, 0.45, 0.55, 0.50, f'{page:02d}', 24, TEAL, True)
        if title:
            title_size = 18.5 if page == 32 else 21
            tx(slide, 1.28, 0.54, 8.10, 0.75, title, title_size, PAPER if dark else INK, True)
        if show_section:
            tx(slide, 1.30, 1.24, 3.5, 0.18, SECTIONS.get(page, ''), 8.5, base, True)
    rule(slide, M, 5.18, 9.38, 5.18, RGBColor(53, 75, 94) if dark else LINE, 0.8)
    tx(slide, M, 5.28, 6.7, 0.16, '智能标书助手 · RFP需求拆解与受控长文档生成', 7.2, base)
    tx(slide, 9.0, 5.26, 0.38, 0.18, f'{page:02d}', 8, base, True, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body='', accent=TEAL, n=None, icon_name=None, dark=False):
    fill = DEEP2 if dark else PAPER
    lc = RGBColor(62, 88, 108) if dark else LINE
    rounded(slide, x, y, w, h, fill, lc)
    box(slide, x, y, 0.055, h, fill=accent, line=None)
    cursor = x + 0.23
    if n is not None:
        circle(slide, cursor, y+0.18, 0.34, accent)
        tx(slide, cursor, y+0.25, 0.34, 0.14, str(n), 8.5, PAPER, True, align=PP_ALIGN.CENTER)
        cursor += 0.48
    elif icon_name:
        icon(slide, icon_name, cursor, y+0.18, 0.36)
        cursor += 0.50
    tx(slide, cursor, y+0.18, w-(cursor-x)-0.20, 0.34, title, 13.5, PAPER if dark else INK, True)
    if body:
        tx(slide, x+0.23, y+0.67, w-0.46, h-0.82, body, 10.4, RGBColor(191,204,214) if dark else GREY)


def status_chip(slide, x, y, label, state):
    cfg = {
        '已满足': (TEAL_SOFT, TEAL), '部分满足': (BLUE_SOFT, BLUE),
        '存在差异': (CORAL_SOFT, CORAL), '待确认': (AMBER_SOFT, AMBER),
        '不适用': (RGBColor(235,238,241), GREY),
    }
    fill, ink = cfg[state]
    w = 0.74 if len(label) <= 4 else 0.92
    rounded(slide, x, y, w, 0.30, fill, None)
    tx(slide, x, y+0.065, w, 0.13, label, 7.5, ink, True, align=PP_ALIGN.CENTER)


def flow(slide, labels, y=2.20, x=0.72, w=8.56, subs=None, colors=None):
    colors = colors or [TEAL, BLUE, AMBER, TEAL, CORAL]
    gap = 0.22
    cw = (w-gap*(len(labels)-1))/len(labels)
    for i, label in enumerate(labels):
        xx = x+i*(cw+gap); c=colors[i % len(colors)]
        rounded(slide, xx, y, cw, 0.98, PAPER, c)
        tx(slide, xx+0.06, y+0.18, cw-0.12, 0.26, label, 11.4, c, True, align=PP_ALIGN.CENTER)
        if subs:
            tx(slide, xx+0.06, y+0.55, cw-0.12, 0.22, subs[i], 8.2, GREY, align=PP_ALIGN.CENTER)
        if i < len(labels)-1:
            arrow(slide, xx+cw+0.03, y+0.37, gap-0.06, 0.22, LINE)


def section_divider(prs, page, num, title, body):
    s = add_slide(prs); chrome(s, page, dark=True, title_style=2, show_section=False)
    tx(s, 0.72, 1.30, 1.25, 0.78, num, 45, TEAL, True)
    tx(s, 1.98, 1.48, 6.95, 0.72, title, 30, PAPER, True)
    tx(s, 1.99, 2.45, 6.50, 0.72, body, 13.5, RGBColor(177,194,206))
    # Signature document-control rail
    rule(s, 1.98, 3.66, 8.86, 3.66, RGBColor(64,91,111), 2.2)
    for i, (lab, c) in enumerate([('原文', TEAL), ('要求', BLUE), ('响应', AMBER), ('章节', TEAL), ('风险', CORAL)]):
        x=2.12+i*1.58; circle(s,x,3.50,0.32,c); tx(s,x-0.23,4.02,0.78,0.22,lab,9,RGBColor(188,202,212),True,align=PP_ALIGN.CENTER)
    note(s,page,body)
    return s


def cover(prs):
    s=add_slide(prs)
    box(s,0,0,W,H,fill=DEEP)
    box(s,0,0,0.12,H,fill=TEAL)
    tx(s,0.72,0.62,3.6,0.25,'AI 产品经理项目课 · 第二阶段',10,RGBColor(151,176,193),True)
    tx(s,0.72,1.25,6.7,0.78,'智能标书助手',36,PAPER,True)
    tx(s,0.74,2.17,6.8,0.52,'RFP需求拆解与受控长文档生成',20,RGBColor(191,208,219),True)
    tx(s,0.74,3.12,5.6,0.60,'不是让AI替你“写得更多”，\n而是让每一句话都知道从哪里来、由谁负责。',15,RGBColor(178,194,206))
    # right document device
    rounded(s,6.72,0.74,2.62,4.02,RGBColor(247,249,250),None)
    tx(s,7.02,1.05,1.6,0.24,'RFP / 186页',9,TEAL,True)
    for i,w in enumerate([1.70,1.92,1.55,1.82,1.38,1.92,1.65]):
        box(s,7.02,1.60+i*0.31,w,0.06,fill=RGBColor(207,216,223),line=None)
    for i,(lab,c) in enumerate([('要求',BLUE),('差异',CORAL),('确认',AMBER)]):
        rounded(s,7.02,3.94+i*0.30,0.70,0.20,rgb(c),None)
        tx(s,7.02,3.985+i*0.30,0.70,0.12,lab,6.3,PAPER,True,align=PP_ALIGN.CENTER)
    rule(s,6.42,2.76,9.55,2.76,TEAL,1.5)
    tag_motif(s.shapes[-1], loud=True)
    note(s,1,'开场先明确：这节课不是教写作技巧，而是用一个完整项目理解高约束长文档怎样被拆解、生成、检查和人工决策。')
    design_intent(s, reason='用被扫描和标注的招标文件作为全课视觉母题，强调可追溯与受控生成。')


def slide_pipeline(prs):
    s=add_slide(prs); chrome(s,2,TITLES[2],title_style=1)
    labels=['拆要求','对能力','搭目录','生成章节','查风险']
    subs=['从原文到卡片','从要求到状态','从评分到结构','从材料到草稿','从草稿到决策']
    flow(s,labels,2.05,0.60,8.82,subs,[TEAL,BLUE,AMBER,TEAL,CORAL])
    tx(s,1.14,3.65,7.72,0.46,'系统真正交付的不是一篇文档，而是一条可以复核的生产链。',17,INK,True,align=PP_ALIGN.CENTER)
    note(s,2,'先给出全课主线。后面每一节都会补齐这条链上的一个控制点，最终落到一个可演示的网页MVP。')


def build_slide(prs, page):
    s=add_slide(prs); chrome(s,page,TITLES.get(page,''),title_style=page%3)
    if page == 4:
        card(s,0.68,1.70,2.60,2.70,'招标项目','某市政务服务智能客服与知识库平台；要求私有化部署并兼容国产环境。',TEAL,icon_name='file-check-2')
        card(s,3.70,1.70,2.60,2.70,'投标企业','启明数科具备产品和交付基础，但资质、案例、周期与服务存在差异。',BLUE,icon_name='briefcase-business')
        card(s,6.72,1.70,2.60,2.70,'课堂任务','把RFP拆成要求，形成响应矩阵，受控生成章节，并在提交前检查风险。',CORAL,icon_name='workflow')
        tx(s,2.72,4.62,4.56,0.24,'课堂模拟案例 · 后续49页均沿用同一组事实',10,MUTE,True,align=PP_ALIGN.CENTER)
        note(s,page,'先交代同一案例的项目、企业和课堂任务。案例是教学模拟，后续所有练习与判断都基于这组固定事实。')
    elif page == 5:
        rows=[('资格','有效信息安全认证；3项政务案例',CORAL),('技术','私有部署；并发≥5,000；国产兼容',BLUE),('交付','60天内完成部署、联调与初验',AMBER),('服务','7×24服务可得3分',TEAL)]
        for i,(a,b,c) in enumerate(rows):
            y=1.60+i*0.73; rounded(s,0.72,y,8.56,0.56,PAPER,LINE); box(s,0.72,y,0.07,0.56,fill=c,line=None)
            tx(s,0.98,y+0.14,0.92,0.22,a,11.5,c,True); tx(s,2.00,y+0.14,6.88,0.23,b,11,INK)
        tx(s,1.05,4.70,7.90,0.25,'要求看起来都很清楚，但企业是否能响应，取决于另一份事实清单。',12.5,GREY,True,align=PP_ALIGN.CENTER)
        note(s,page,'不要评价能不能做，先只读招标方的要求。让学员记住：要求本身和企业能力是两张表，不能混在一起。')
    elif page == 6:
        facts=[('安全认证','已过有效期',CORAL),('行业案例','2项政务 + 1项零售',AMBER),('并发能力','实测10,000用户',TEAL),('私有部署','正式方案已确认',TEAL),('交付周期','标准90天',CORAL),('服务能力','标准5×8',AMBER),('国产兼容','确认3类环境',TEAL),('历史说法','“全面国产适配”无正式材料',CORAL)]
        for i,(a,b,c) in enumerate(facts):
            x=0.64+(i%4)*2.23; y=1.56+(i//4)*1.50
            rounded(s,x,y,2.05,1.23,PAPER,c)
            tx(s,x+0.16,y+0.18,1.70,0.27,a,11.5,c,True)
            tx(s,x+0.16,y+0.62,1.70,0.38,b,9.5,GREY)
        tx(s,0.98,4.61,8.04,0.32,'事实库的价值，不是让AI知道更多，而是让AI不敢乱写。',14,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'逐项介绍公司真实情况。特别强调历史标书不等于正式事实，旧文档里的漂亮说法可能正是风险来源。')
    elif page == 7:
        rounded(s,0.72,1.55,8.56,3.30,PAPER,LINE)
        tx(s,0.99,1.82,2.25,0.25,'AI 初稿 · 技术方案节选',10,TEAL,True)
        draft='启明数科拥有完善的信息安全管理体系和丰富的政务行业经验，已成功服务多个政府客户。平台全面兼容各类国产操作系统与数据库，可在60天内完成项目交付，并提供7×24小时技术支持。系统最大并发能力可达到10,000用户，充分满足本项目要求。'
        tx(s,1.00,2.32,7.90,1.28,draft,16,INK,False)
        for x,label,c in [(1.02,'无依据资质',CORAL),(2.35,'案例泛化',AMBER),(3.55,'全面兼容',CORAL),(4.77,'60天承诺',CORAL),(5.89,'7×24承诺',AMBER)]:
            status_chip(s,x,4.03,label,'存在差异' if c==CORAL else '待确认')
        note(s,page,'先完整读一遍，不要马上给答案。这段文字流畅、专业，也把系统最危险的地方暴露出来：它把要求自动改写成了承诺。')
    elif page == 8:
        tx(s,0.78,1.67,2.2,0.27,'请圈出至少 4 个问题',13,CORAL,True)
        rounded(s,0.78,2.07,5.80,2.40,PAPER,LINE)
        tx(s,1.05,2.38,5.25,1.55,'“拥有完善的信息安全体系”\n“丰富的政务行业经验”\n“全面兼容各类国产软硬件”\n“60天完成交付并提供7×24服务”',16,INK,False)
        card(s,6.88,2.07,2.34,2.40,'判断提示','它来自哪份材料？\n与现状是否冲突？\n谁有权做出承诺？\n是否需要保留条件？',AMBER,n='?')
        note(s,page,'给学员2—3分钟。要求他们不只找事实错误，还要找承诺、责任、范围和材料引用问题。')
    elif page == 9:
        issues=[('事实失真','把过期证书写成有效',CORAL),('范围泛化','把三类适配写成全面兼容',CORAL),('跨域替代','零售案例替代政务案例',AMBER),('承诺越权','自动承诺60天和7×24',CORAL),('来源缺失','没有材料编号与页码',BLUE),('不确定性消失','没有待确认标记',AMBER)]
        for i,(a,b,c) in enumerate(issues):
            x=0.68+(i%3)*2.98; y=1.58+(i//3)*1.46
            card(s,x,y,2.72,1.20,a,b,c,n=i+1)
        note(s,page,'公布答案。把“AI幻觉”拆成六种产品问题，帮助学员看到这不是只靠换一个提示词就能解决的。')
    elif page == 10:
        # comparison table
        tx(s,0.72,1.55,2.0,0.24,'比较维度',9,MUTE,True); tx(s,3.05,1.55,2.6,0.24,'通用AI写作工具',9,MUTE,True); tx(s,6.35,1.55,2.75,0.24,'智能标书系统',9,TEAL,True)
        rows=[('输入','一段指令 + 若干文件','结构化要求 + 企业材料 + 决策'),('过程','一次生成正文','拆解、映射、生成、检查'),('输出','看起来完整的文章','可追踪的章节与风险'),('责任','模型替用户补全','把不确定项留给人'),('成功标准','写得像','写得对、写得住、查得到')]
        for i,(a,b,c) in enumerate(rows):
            y=1.88+i*0.57; rounded(s,0.70,y,8.60,0.46,PAPER,LINE); tx(s,0.92,y+0.12,1.56,0.20,a,10.5,INK,True); tx(s,3.05,y+0.12,2.70,0.20,b,10,GREY); tx(s,6.35,y+0.12,2.65,0.20,c,10,TEAL,True)
        note(s,page,'收束第一节。真正的产品价值不是“更会写”，而是建立一套从原文到承诺的控制机制。')
    elif page == 12:
        blocks=[('商务与资格','主体、资质、案例、财务'),('技术需求','功能、性能、部署、兼容'),('实施交付','里程碑、周期、验收、培训'),('服务保障','响应时间、驻场、维护'),('评分办法','分值、得分条件、扣分点'),('格式程序','签章、装订、时间、提交方式')]
        for i,(a,b) in enumerate(blocks):
            x=0.68+(i%3)*2.98; y=1.60+(i//3)*1.38
            card(s,x,y,2.72,1.12,a,b,[TEAL,BLUE,AMBER,CORAL,TEAL,GREY][i],n=i+1)
        note(s,page,'讲RFP的典型组成。它不是按“产品经理好理解”的方式写，而是按采购、法务、技术和程序共同需要的方式组织。')
    elif page == 13:
        cats=[('主体资格','能不能投'),('证明材料','拿什么证明'),('业务目标','为什么建设'),('功能要求','系统做什么'),('非功能','性能与安全'),('交付要求','何时怎样交'),('服务要求','上线后谁负责'),('评分项','怎样拿分')]
        for i,(a,b) in enumerate(cats):
            x=0.58+(i%4)*2.27; y=1.55+(i//4)*1.42
            rounded(s,x,y,2.05,1.17,PAPER,[TEAL,BLUE,AMBER,CORAL][i%4]); tx(s,x+0.18,y+0.22,1.65,0.29,a,12.5,[TEAL,BLUE,AMBER,CORAL][i%4],True); tx(s,x+0.18,y+0.67,1.65,0.23,b,9.5,GREY)
        note(s,page,'八类不是唯一分类法，但足够支持课堂MVP。让学员理解：分类的目的，是后续使用不同的响应与检查规则。')
    elif page == 14:
        card(s,0.72,1.63,4.06,2.90,'摘要','“项目要求投标人具备相关资质和案例，系统需支持私有化部署、国产兼容，并按期交付。”\n\n适合：快速了解文件大意。',MUTE,n='A')
        card(s,5.18,1.63,4.10,2.90,'要求清单','R-001 有效认证｜资格｜必须\nR-002 3项政务案例｜资格｜必须\nR-003 私有部署｜技术｜必须\n……\n\n适合：逐条响应、检查与追踪。',TEAL,n='B')
        note(s,page,'比较两种产出。摘要会主动合并信息，而系统恰恰需要保留每个最小条件，因此摘要不能直接替代要求解析。')
    elif page == 15:
        rounded(s,0.72,1.58,8.56,0.78,PAPER,LINE)
        tx(s,0.96,1.82,8.05,0.28,'“系统应支持政务内网私有化部署，整体并发用户数不低于5,000，并兼容国产操作系统及数据库。”',13.2,INK,True)
        arrow(s,4.74,2.55,0.25,0.52,TEAL)
        for i,(a,b,c) in enumerate([('R-003A','支持私有化部署',TEAL),('R-003B','并发用户≥5,000',BLUE),('R-003C','兼容国产OS与数据库',AMBER)]):
            x=0.90+i*2.92; rounded(s,x,3.27,2.60,0.92,PAPER,c); tx(s,x+0.17,3.47,0.72,0.22,a,9,c,True); tx(s,x+0.17,3.82,2.20,0.22,b,10.5,INK,True)
        note(s,page,'讲“原子化”。判断标准是：如果三个条件可能出现不同响应状态，就必须拆成三项。')
    elif page == 16:
        rounded(s,0.70,1.50,8.60,3.20,PAPER,TEAL)
        tx(s,0.98,1.76,1.25,0.26,'R-004',15,TEAL,True); status_chip(s,7.92,1.72,'存在差异','存在差异')
        tx(s,0.98,2.25,7.90,0.35,'合同签订后 60 个自然日内完成部署、联调和初验',15,INK,True)
        fields=[('分类','交付'),('强制性','必须'),('来源','第42页 3.3.4'),('验收方式','初验通过'),('企业现状','标准周期90天'),('负责人','交付负责人'),('处理动作','压缩范围 / 增加资源 / 商务澄清'),('材料','项目计划、资源测算')]
        for i,(a,b) in enumerate(fields):
            x=0.98+(i%4)*2.02; y=3.02+(i//4)*0.64
            tx(s,x,y,0.62,0.18,a,7.5,MUTE,True); tx(s,x,y+0.23,1.73,0.23,b,9.4,INK,True)
        note(s,page,'逐字段解释要求卡片。字段不是越多越好，核心是支撑后续的状态判断、章节映射、检查和责任分配。')
    elif page == 17:
        rounded(s,0.75,1.58,8.50,1.45,PAPER,LINE)
        tx(s,1.00,1.87,8.00,0.88,'投标人须提供有效的信息安全认证及近三年不少于三项政务行业案例；系统应支持私有化部署，并发用户数不低于5,000，兼容国产操作系统及数据库，在60日内完成初验。',14.5,INK,True)
        tx(s,0.94,3.47,3.8,0.28,'任务：拆成最小可判定要求',14,CORAL,True)
        for i in range(6):
            x=0.95+(i%3)*2.75; y=3.92+(i//3)*0.47
            rounded(s,x,y,2.48,0.34,RGBColor(237,240,243),None); tx(s,x+0.12,y+0.09,0.35,0.13,f'{i+1}',7.5,MUTE,True)
        note(s,page,'给学员3分钟。要求每条只包含一个可判定条件，并标注大致分类。')
    elif page == 18:
        ans=[('1','有效信息安全认证','资格'),('2','3项政务行业案例','资格'),('3','支持私有化部署','技术'),('4','并发≥5,000','技术'),('5','国产环境兼容','技术'),('6','60日内完成初验','交付')]
        for i,(n,a,b) in enumerate(ans):
            x=0.68+(i%3)*2.98; y=1.56+(i//3)*1.38
            card(s,x,y,2.72,1.10,a,b,[CORAL,CORAL,TEAL,BLUE,AMBER,CORAL][i],n=n)
        tx(s,0.92,4.50,8.16,0.38,'拆分的边界：每一项都可以独立判断“满足 / 不满足 / 待确认”。',14,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'公布答案并解释为何是六项。提醒学员：原句里“国产操作系统及数据库”可在更细系统中继续拆分，本课MVP先保留为一项。')
    elif page == 20:
        # signature page: source → card → matrix
        rounded(s,0.55,1.52,2.52,3.10,RGBColor(251,250,246),LINE)
        tx(s,0.82,1.80,1.55,0.22,'招标原文 · P42',9,TEAL,True)
        tx(s,0.83,2.28,1.98,1.44,'系统应支持政务内网私有化部署，整体并发用户数不低于5,000。',13,INK,True)
        box(s,0.77,3.03,2.07,0.47,fill=TEAL_SOFT,line=None)
        arrow(s,3.20,2.78,0.50,0.24,TEAL)
        rounded(s,3.84,1.86,2.30,2.38,PAPER,TEAL)
        tx(s,4.10,2.13,0.78,0.22,'R-003',11,TEAL,True); status_chip(s,5.10,2.08,'已满足','已满足')
        tx(s,4.10,2.68,1.78,0.70,'私有部署\n并发 ≥ 5,000',13.5,INK,True)
        tx(s,4.10,3.65,1.78,0.22,'来源、分类、判定规则',8.5,MUTE)
        arrow(s,6.28,2.78,0.48,0.24,TEAL)
        rounded(s,6.90,1.62,2.56,3.00,PAPER,LINE)
        tx(s,7.14,1.92,1.65,0.22,'响应矩阵',10,TEAL,True)
        tx(s,7.14,2.40,0.65,0.18,'现状',8,MUTE,True); tx(s,7.14,2.68,1.88,0.32,'私有部署 + 10,000并发',10.2,INK,True)
        tx(s,7.14,3.22,0.65,0.18,'动作',8,MUTE,True); tx(s,7.14,3.50,1.88,0.50,'引用部署方案和性能测试报告',10.2,INK,True)
        rule(s,0.75,4.82,9.25,4.82,TEAL,1.5)
        note(s,page,'这是全课的签名页。强调一条要求要能从原文定位，进入要求卡片，再接到企业现状和处理动作，三段缺一不可。')
        design_intent(s, reason='原文、要求卡片与矩阵连成一条审计轨迹，后续所有页面沿用此结构。')
    elif page == 21:
        states=[('已满足','事实与要求一致，可直接响应',TEAL),('部分满足','只有部分条件成立，需要限定',BLUE),('存在差异','现状与强制要求冲突',CORAL),('待确认','需要商务、交付或管理层决策',AMBER),('不适用','与本项目或本章节无关',GREY)]
        for i,(a,b,c) in enumerate(states):
            x=0.56+i*1.82
            rounded(s,x,1.72,1.62,2.55,PAPER,c); circle(s,x+0.57,2.03,0.48,c); tx(s,x+0.18,2.75,1.26,0.28,a,12.5,c,True,align=PP_ALIGN.CENTER); tx(s,x+0.18,3.25,1.26,0.58,b,9.1,GREY,align=PP_ALIGN.CENTER)
        note(s,page,'五种状态的关键是互斥且可行动。尤其不要把“部分满足”当成模糊地带，它必须说明哪部分满足、哪部分不满足。')
    elif page in (22,23):
        selected=REQS[:3] if page==22 else REQS[3:]
        headers=['编号','要求','企业现状','状态']
        for x,w,h in [(0.70,0.85,'编号'),(1.60,3.02,'要求'),(4.70,3.00,'企业现状'),(7.88,1.12,'状态')]: tx(s,x,1.56,w,0.20,h,8.5,MUTE,True)
        for i,(rid,typ,req,fact,state) in enumerate(selected):
            y=1.90+i*0.86; rounded(s,0.66,y,8.68,0.68,PAPER,LINE); tx(s,0.88,y+0.20,0.62,0.20,rid,9.5,TEAL,True); tx(s,1.60,y+0.15,2.84,0.34,req,10.2,INK,True); tx(s,4.70,y+0.15,2.90,0.34,fact,9.7,GREY); status_chip(s,7.88,y+0.18,state,state)
        bottom='资质与技术响应必须有正式材料支撑，历史文案不能代替事实。' if page==22 else '交付与服务属于组织承诺，AI只能暴露差异，不能代替决策。'
        tx(s,1.00,4.72,8.00,0.26,bottom,13,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,bottom)
    elif page == 24:
        rounded(s,0.72,1.60,8.56,2.34,PAPER,AMBER)
        tx(s,1.02,1.91,3.20,0.24,'初步判断',10,AMBER,True)
        tx(s,1.02,2.36,4.35,0.48,'有条件参标',27,INK,True)
        tx(s,1.02,3.05,4.72,0.50,'核心技术能力满足；资质、案例、交付周期与服务承诺需要进一步决策。',12,GREY)
        circle(s,7.46,2.06,1.12,AMBER_SOFT,AMBER); tx(s,7.46,2.30,1.12,0.34,'4',25,AMBER,True,align=PP_ALIGN.CENTER); tx(s,7.46,2.80,1.12,0.20,'项需决策',8,AMBER,True,align=PP_ALIGN.CENTER)
        tx(s,1.05,4.42,7.90,0.36,'系统不是替公司决定要不要投，而是让决策建立在一张完整差异表上。',14,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'解释参标判断的边界。MVP只做初判和风险暴露，最终决策仍然属于销售、交付、法务和管理层。')
    elif page == 25:
        for i,(rid,typ,req,fact,state) in enumerate(REQS):
            x=0.62+(i%2)*4.55; y=1.50+(i//2)*1.10
            rounded(s,x,y,4.25,0.88,PAPER,LINE); tx(s,x+0.18,y+0.17,0.68,0.20,rid,9,TEAL,True); tx(s,x+0.92,y+0.14,2.52,0.35,req,9.6,INK,True); rounded(s,x+3.50,y+0.23,0.52,0.28,RGBColor(237,240,243),None); tx(s,x+3.50,y+0.30,0.52,0.12,'？',7.5,MUTE,True,align=PP_ALIGN.CENTER)
        note(s,page,'让学员基于前面已经给出的公司事实，为六条要求标状态。不要新增假设。')
    elif page == 26:
        for i,(rid,typ,req,fact,state) in enumerate(REQS):
            x=0.62+(i%2)*4.55; y=1.50+(i//2)*1.10
            rounded(s,x,y,4.25,0.88,PAPER,LINE); tx(s,x+0.18,y+0.17,0.68,0.20,rid,9,TEAL,True); tx(s,x+0.92,y+0.14,2.08,0.35,req,9.4,INK,True); status_chip(s,x+3.15,y+0.23,state,state)
        note(s,page,'公布答案。重点讨论R-002为什么是部分满足、R-006为什么不是存在差异：因为它是评分项，企业可以选择不拿这3分，也可以升级服务。')
    elif page == 28:
        probs=[('要求遗漏','长文档中某些条款没有进入正文'),('章节冲突','不同章节出现不同数字和承诺'),('材料误用','把旧标书和未确认说法当事实'),('责任越权','模型替管理层承诺成本与周期')]
        for i,(a,b) in enumerate(probs):
            x=0.72+i*2.20; card(s,x,1.72,1.98,2.45,a,b,[TEAL,BLUE,AMBER,CORAL][i],n=i+1)
        note(s,page,'说明为什么不能one-shot。文本越长，模型越容易在局部写得通顺、在全局失去一致性。')
    elif page == 29:
        tx(s,0.72,1.55,2.0,0.22,'一级目录',9,MUTE,True)
        roots=[('01 资格证明',CORAL),('02 技术方案',TEAL),('03 实施方案',BLUE),('04 服务方案',AMBER)]
        for i,(a,c) in enumerate(roots):
            y=1.90+i*0.69; rounded(s,0.72,y,2.20,0.50,PAPER,c); tx(s,0.90,y+0.14,1.83,0.20,a,10.5,c,True)
        rule(s,3.25,1.91,3.25,4.43,LINE,1.3)
        tx(s,3.72,1.55,2.0,0.22,'二级章节任务',9,MUTE,True)
        subs=[('2.1 总体架构','回应系统边界与部署模式'),('2.2 性能与可靠性','回应并发、可用性、容灾'),('2.3 国产兼容','回应已验证环境与适配边界'),('2.4 安全方案','回应认证、权限与审计')]
        for i,(a,b) in enumerate(subs):
            y=1.83+i*0.72; rounded(s,3.70,y,5.55,0.58,PAPER,LINE); tx(s,3.92,y+0.12,1.56,0.20,a,10.5,TEAL,True); tx(s,5.60,y+0.12,3.35,0.28,b,9.5,GREY)
        note(s,page,'解释两级目录树：一级解决整份文档结构，二级把每章变成可以独立生成和检查的任务。')
    elif page == 30:
        rows=[('评分 15 分','核心技术方案','深写：架构、流程、指标、材料',TEAL),('评分 8 分','实施与交付','中写：阶段、资源、验收',BLUE),('评分 3 分','7×24服务','谨慎写：决策后再承诺',AMBER),('符合性','格式与签章','规则校验，不靠生成',GREY)]
        for i,(a,b,c,col) in enumerate(rows):
            y=1.62+i*0.72; rounded(s,0.72,y,8.56,0.56,PAPER,LINE); tx(s,0.95,y+0.14,1.22,0.22,a,10,col,True); tx(s,2.35,y+0.14,2.10,0.22,b,10.5,INK,True); tx(s,4.75,y+0.14,4.15,0.22,c,10,GREY)
        note(s,page,'章节深度不应平均分配。高分项需要更具体的方案与材料，低分或高风险承诺项则要谨慎，程序性要求优先用规则检查。')
    elif page == 31:
        maps=[('R-003 私有部署','3.1 总体架构','说明部署边界与网络分区',TEAL),('R-003 并发≥5,000','3.3 性能与可靠性','引用10,000并发测试报告',BLUE),('R-005 国产兼容','3.2 部署与兼容','逐项列出已验证环境',AMBER),('R-004 60天交付','4.1 实施计划','标记差异，等待资源决策',CORAL)]
        for i,(a,b,c,col) in enumerate(maps):
            y=1.62+i*0.73
            rounded(s,0.67,y,2.40,0.56,PAPER,col); tx(s,0.86,y+0.15,2.05,0.22,a,9.7,col,True)
            arrow(s,3.18,y+0.17,0.55,0.22,LINE)
            rounded(s,3.88,y,2.20,0.56,PAPER,LINE); tx(s,4.05,y+0.15,1.86,0.22,b,9.7,INK,True)
            arrow(s,6.19,y+0.17,0.48,0.22,LINE)
            rounded(s,6.80,y,2.54,0.56,PAPER,LINE); tx(s,6.98,y+0.12,2.20,0.30,c,9.2,GREY,True)
        note(s,page,'映射表是生成和检查的共同基础。它告诉系统某条要求在哪里回答、回答到什么程度、引用哪份材料。')
    elif page == 32:
        rounded(s,0.70,1.52,8.60,3.22,PAPER,BLUE)
        tx(s,0.96,1.78,2.30,0.25,'章节任务包 · 3.2 部署与兼容性',11,BLUE,True)
        fields=[('目标','回应 R-003、R-005'),('必须包含','私有部署；三类已验证环境'),('可用材料','部署方案 v3.2；兼容清单'),('禁止写入','有效认证；全面兼容；60天承诺'),('表达要求','按环境逐项列示，保留适配条件'),('待确认','其他国产版本的适配范围')]
        for i,(a,b) in enumerate(fields):
            x=0.98+(i%2)*4.08; y=2.33+(i//2)*0.69
            tx(s,x,y,0.75,0.18,a,8,MUTE,True); tx(s,x+0.88,y-0.02,3.02,0.34,b,10.2,INK,True)
        note(s,page,'任务包相当于生成接口。它把“写一章”变成有输入、有边界、有验收条件的任务。')
    elif page == 33:
        rules=[('确认等级','正式材料 / 待确认 / 禁用'),('适用范围','哪个产品、版本、地区、时间'),('有效期限','材料是否过期或需要复核'),('引用方式','材料编号、页码、负责人'),('冲突优先级','新材料优先，正式材料优先'),('使用动作','可直接写 / 只能改写 / 必须人工确认')]
        for i,(a,b) in enumerate(rules):
            x=0.68+(i%3)*2.98; y=1.58+(i//3)*1.42; card(s,x,y,2.72,1.16,a,b,[TEAL,BLUE,AMBER,CORAL,TEAL,GREY][i],n=i+1)
        note(s,page,'知识库解决“找得到”，使用规则解决“能不能用”。对标书场景，材料治理往往比向量检索本身更重要。')
    elif page == 34:
        rounded(s,0.68,1.50,5.85,3.22,PAPER,LINE)
        tx(s,0.96,1.78,2.15,0.23,'生成结果 · 3.2 部署与兼容性',10,TEAL,True)
        tx(s,0.96,2.23,5.23,1.48,'启明数科平台支持在客户政务内网环境中进行私有化部署。当前已完成对麒麟 V10、达梦数据库 8、人大金仓数据库的适配验证，并可提供兼容性清单与部署说明。\n\n其他国产软硬件环境需结合实际版本进行适配评估，本方案不作未经验证的兼容承诺。',12.2,INK)
        status_chip(s,1.00,4.28,'已引用2份材料','已满足')
        card(s,6.83,1.50,2.45,1.40,'写进去','私有部署\n三类已验证环境',TEAL,n='✓')
        card(s,6.83,3.18,2.45,1.54,'留给人','其他环境适配范围\n是否追加承诺',AMBER,n='!')
        note(s,page,'逐句对照任务包。好的受控生成不是没有不确定性，而是不会偷偷把不确定性改写成确定承诺。')
    elif page == 36:
        checks=[('遗漏','要求是否都有回答'),('矛盾','跨章节数字是否一致'),('失真','事实是否有材料支撑'),('越权','是否出现未经批准承诺'),('范围','适用版本与条件是否明确'),('引用','材料编号与页码是否齐全'),('格式','签章、目录、命名是否符合')]
        for i,(a,b) in enumerate(checks):
            x=0.48+i*1.34; rounded(s,x,1.72,1.18,2.52,PAPER,[TEAL,BLUE,AMBER,CORAL,TEAL,BLUE,GREY][i]); circle(s,x+0.39,2.02,0.40,[TEAL,BLUE,AMBER,CORAL,TEAL,BLUE,GREY][i]); tx(s,x+0.13,2.65,0.92,0.25,a,12,[TEAL,BLUE,AMBER,CORAL,TEAL,BLUE,GREY][i],True,align=PP_ALIGN.CENTER); tx(s,x+0.13,3.15,0.92,0.58,b,8.8,GREY,align=PP_ALIGN.CENTER)
        note(s,page,'七类检查覆盖内容完整性、事实可靠性和程序合规性。课堂MVP不必全部做深，但至少要覆盖遗漏、矛盾和越权。')
    elif page == 37:
        tx(s,0.72,1.52,1.75,0.22,'同一个承诺，出现在四处',10,MUTE,True)
        items=[('技术方案','支持10,000并发',TEAL),('实施方案','资源按5,000并发配置',CORAL),('报价表','只含单节点部署',CORAL),('服务方案','承诺7×24响应',AMBER)]
        for i,(a,b,c) in enumerate(items):
            x=0.70+i*2.20; rounded(s,x,2.02,1.98,1.46,PAPER,c); tx(s,x+0.18,2.27,1.62,0.26,a,11.5,c,True,align=PP_ALIGN.CENTER); tx(s,x+0.18,2.82,1.62,0.38,b,9.5,INK,True,align=PP_ALIGN.CENTER)
            if i<3: rule(s,x+1.98,2.75,x+2.18,2.75,LINE,1.2)
        tx(s,1.05,4.12,7.90,0.42,'单章都“写得通”，整份标书仍然可能互相打架。',17,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'用并发、资源和服务例子讲跨章节一致性。检查不能只对每章单独打分，需要建立全局事实与承诺表。')
    elif page == 38:
        left='材料事实\n“标准交付周期为90天”'; right='AI补写\n“为体现项目重视，可在60天内完成”'
        card(s,0.78,1.76,3.42,1.70,'事实',left,TEAL,n='1')
        arrow(s,4.35,2.37,1.06,0.26,CORAL)
        card(s,5.65,1.76,3.57,1.70,'听起来更专业',right,CORAL,n='2')
        tx(s,1.00,4.02,8.00,0.55,'过度承诺并不总像“胡编乱造”，它常常只是模型为了让文案更完整，主动补了一句组织从未决定的话。',14,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'强调过度承诺的隐蔽性。模型优化的是文本完整度，而企业需要控制的是责任边界，这两者天然不同。')
    elif page == 39:
        rounded(s,0.68,1.48,3.02,3.28,PAPER,LINE)
        tx(s,0.92,1.75,1.40,0.22,'背景事实',10,TEAL,True)
        tx(s,0.92,2.20,2.52,1.70,'• 标准交付：90天\n• 标准服务：5×8\n• 国产兼容：已确认3类环境\n• 安全认证：已过期\n• 政务案例：2项',11,INK)
        rounded(s,3.98,1.48,5.34,3.28,PAPER,LINE)
        tx(s,4.24,1.75,1.40,0.22,'待检查草稿',10,CORAL,True)
        tx(s,4.24,2.20,4.75,1.72,'我司具备完善资质与丰富政务经验，可全面适配国产软硬件环境，并在60天内完成交付。项目上线后提供7×24小时服务，确保系统长期稳定运行。',13.2,INK,True)
        tx(s,4.24,4.24,4.70,0.24,'任务：找出问题，并判断该由规则、AI还是人工处理。',10.5,CORAL,True)
        note(s,page,'这次练习重复必要背景，避免学生只能凭感觉找问题。答案包括资质失真、案例泛化、全面兼容、60天承诺、7×24承诺。')
    elif page == 40:
        cols=[('规则','确定性检查','数字一致、必填、格式、签章',TEAL),('AI','语义性检查','遗漏、矛盾、无依据表述、范围泛化',BLUE),('人工','责任性决策','参不参标、能否承诺、是否增加资源',AMBER)]
        for i,(a,b,c,col) in enumerate(cols):
            x=0.70+i*3.02; rounded(s,x,1.64,2.72,2.83,PAPER,col); circle(s,x+1.05,1.98,0.62,col); tx(s,x+0.24,2.83,2.24,0.27,a,16,col,True,align=PP_ALIGN.CENTER); tx(s,x+0.24,3.30,2.24,0.22,b,10,INK,True,align=PP_ALIGN.CENTER); tx(s,x+0.28,3.70,2.16,0.45,c,9.1,GREY,align=PP_ALIGN.CENTER)
        note(s,page,'给出职责分工。能写成明确if-then的优先用规则；需要理解上下文的交给AI；涉及成本、责任和策略的必须由人决定。')
    elif page == 41:
        labels=['发现问题','定位来源','判断类型','回到负责人','修改事实/要求/草稿','重新检查']
        subs=['检查结果','原文/材料/决策','事实/规则/承诺','业务/交付/管理层','改根因，不只改句子','确认问题消失']
        flow(s,labels,1.95,0.50,9.02,subs,[CORAL,BLUE,AMBER,TEAL,BLUE,TEAL])
        tx(s,1.03,3.82,7.94,0.50,'风险处理不是“把红字改绿”，而是回到产生问题的那一层。',17,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'讲问题闭环。比如60天冲突不能只删掉数字，要么改变资源计划，要么改变响应策略，要么放弃参标。')
    elif page == 43:
        labels=['上传RFP','拆解要求','形成矩阵','生成章节','检查风险','人工导出']
        subs=['PDF/Word','卡片+来源','状态+动作','任务包驱动','七类检查','保留决策记录']
        flow(s,labels,1.96,0.48,9.04,subs,[TEAL,BLUE,AMBER,TEAL,CORAL,GREY])
        card(s,1.08,3.69,7.84,0.96,'MVP边界','只支持一个项目、六条要求、一个技术章节和一次风险检查；但全链路必须真实可演示。',TEAL,icon_name='workflow')
        note(s,page,'定义MVP。功能少没有关系，关键是从文件到导出真正走通，中间每个结果都能回到来源。')
    elif page == 44:
        # two UI mockups
        for x,title in [(0.58,'页面1 · 项目主页'),(5.18,'页面2 · 要求解析')]:
            rounded(s,x,1.52,4.24,3.15,PAPER,LINE); box(s,x,1.52,4.24,0.40,fill=DEEP,line=None); tx(s,x+0.20,1.65,2.30,0.15,title,8,PAPER,True)
        rounded(s,0.88,2.18,3.64,0.78,DEEP2,None); tx(s,1.10,2.40,2.50,0.18,'政务智能客服平台',11,PAPER,True); rounded(s,1.00,3.20,0.92,0.72,TEAL_SOFT,None); rounded(s,2.12,3.20,0.92,0.72,BLUE_SOFT,None); rounded(s,3.24,3.20,0.92,0.72,CORAL_SOFT,None)
        rounded(s,5.47,2.10,1.65,2.20,RGBColor(251,250,246),LINE); for_text='招标原文\n\n有效认证……\n3项案例……\n私有部署……'; tx(s,5.69,2.35,1.20,1.45,for_text,8.6,INK)
        for i,c in enumerate([CORAL,AMBER,TEAL]):
            rounded(s,7.38,2.12+i*0.68,1.72,0.52,PAPER,c); tx(s,7.55,2.26+i*0.68,1.38,0.20,f'R-00{i+1} 要求卡片',8.2,c,True)
        note(s,page,'页面1强调项目状态和下一步；页面2强调左侧原文、右侧要求卡片，点击卡片可定位原文。')
    elif page == 45:
        labels=['页面3\n响应矩阵','页面4\n章节编辑','页面5\n风险检查与导出']
        cols=[TEAL,BLUE,CORAL]
        for i,(lab,c) in enumerate(zip(labels,cols)):
            x=0.55+i*3.12; rounded(s,x,1.50,2.82,3.28,PAPER,LINE); box(s,x,1.50,2.82,0.38,fill=DEEP,line=None); tx(s,x+0.16,1.61,2.22,0.16,lab.replace('\n',' · '),7.5,PAPER,True)
        # matrix
        for r in range(4):
            y=2.16+r*0.45; rounded(s,0.78,y,2.36,0.30,RGBColor(247,249,250),None); status_chip(s,2.35,y+0.01,['已满足','部分满足','存在差异','待确认'][r],['已满足','部分满足','存在差异','待确认'][r])
        # editor
        box(s,3.88,2.05,0.70,2.20,fill=RGBColor(239,243,246),line=None)
        tx(s,4.75,2.18,1.55,0.22,'3.2 部署与兼容',8.5,INK,True)
        for i in range(5):
            box(s,4.75,2.63+i*0.28,1.20+(i%2)*0.50,0.05,fill=LINE,line=None)
        # risk and export
        for i,c in enumerate([CORAL,CORAL,AMBER,AMBER]):
            rounded(s,7.06,2.10+i*0.52,2.12,0.38,PAPER,c); circle(s,7.20,2.21+i*0.52,0.16,c); tx(s,7.48,2.20+i*0.52,1.45,0.14,'风险项与处理动作',7.2,INK,True)
        rounded(s,7.35,4.28,1.55,0.28,DEEP,None); tx(s,7.35,4.35,1.55,0.12,'确认后导出标书',7.2,PAPER,True,align=PP_ALIGN.CENTER)
        note(s,page,'快速走过后三个页面。响应矩阵用于决策，编辑器用于受控生成；风险处理完成后，才允许导出可提交版本。')
    elif page == 46:
        tests=[('能走完','文件→要求→矩阵→章节→检查'),('能追溯','每条要求和句子能回到来源'),('能拦住','差异、待确认和越权承诺不会静默通过'),('能修改','人可以更改状态、材料与草稿'),('能演示','5分钟内完成一条端到端流程')]
        for i,(a,b) in enumerate(tests):
            y=1.57+i*0.65; circle(s,0.82,y+0.05,0.33,TEAL); tx(s,0.82,y+0.13,0.33,0.12,'✓',7,PAPER,True,align=PP_ALIGN.CENTER); tx(s,1.35,y,1.18,0.28,a,13,TEAL,True); tx(s,2.72,y+0.01,6.10,0.28,b,11,INK)
        tx(s,1.08,4.80,7.84,0.24,'验收的核心不是模型多聪明，而是产品有没有把风险变得可见、可控、可追踪。',12.5,GREY,True,align=PP_ALIGN.CENTER)
        note(s,page,'用五条验收标准收束产品设计。学员做作业时，优先检查链路而不是追求复杂模型或精美界面。')
    elif page == 48:
        scenes=[('政府项目申报','政策条件 → 申报资格 → 材料章节 → 合规检查','把政策条款拆成申报清单',TEAL),('法律合同审核','合同条款 → 公司底线 → 修改意见 → 风险确认','对应知乎第5课案例的迁移讲法',BLUE),('供应商安全问卷','问题项 → 安全事实 → 回答草稿 → 缺口跟踪','大量重复问卷的事实复用',AMBER)]
        for i,(a,b,c,col) in enumerate(scenes):
            x=0.67+i*3.00; rounded(s,x,1.57,2.73,2.92,PAPER,col); tx(s,x+0.22,1.88,2.28,0.35,a,15,col,True); rule(s,x+0.22,2.40,x+2.48,2.40,LINE,1); tx(s,x+0.22,2.67,2.28,0.83,b,10.2,INK,True); tx(s,x+0.22,3.72,2.28,0.48,c,9.3,GREY)
        note(s,page,'迁移时不要只换业务名词，要寻找相同结构：外部高约束文本、内部事实材料、逐项响应、正式承诺、提交前检查。合同审核案例可重点展开。')
    elif page == 49:
        steps=[('01','准备一份20—50页RFP或模拟文件'),('02','拆出至少20条要求并形成卡片'),('03','为每条要求标响应状态与处理动作'),('04','生成一个受控章节并保留材料引用'),('05','设计一次风险检查并完成5分钟演示')]
        for i,(n,a) in enumerate(steps):
            x=0.50+i*1.83; rounded(s,x,1.58,1.63,2.75,PAPER,[TEAL,BLUE,AMBER,CORAL,TEAL][i]); tx(s,x+0.18,1.85,0.44,0.27,n,13,[TEAL,BLUE,AMBER,CORAL,TEAL][i],True); tx(s,x+0.18,2.47,1.28,1.05,a,11.2,INK,True)
        tx(s,1.00,4.62,8.00,0.35,'完成标准：不是生成了一篇标书，而是能解释每一步为什么可信。',15,INK,True,align=PP_ALIGN.CENTER)
        note(s,page,'布置作业并收尾。展示时要求学员从一条具体要求开始，完整走过来源、状态、章节和风险，而不是只展示最终页面。')
    else:
        raise ValueError(page)
    return s


def main():
    set_palette(deep=DEEP, blue=BLUE, teal=TEAL, magenta=CORAL, slate=GREY,
                mute=MUTE, mono=INK, font=FONT, display=FONT, eafont=EA,
                eadisplay=EA, accents=[TEAL,BLUE,AMBER,CORAL])
    set_geometry(radius=0.11, rule_w=1.0)
    set_ground(BG)
    prs=blank_deck(W,H)
    prs.core_properties.title='智能标书助手：严格按照49页大纲生成'
    prs.core_properties.subject='AI产品经理项目课 · 严格大纲版 v2.0'
    prs.core_properties.author='AI Course'
    cover(prs)
    slide_pipeline(prs)
    dividers={
        3:('第一节','为什么需要智能标书系统','先从一个“写得很像、却不能交”的AI草稿开始。'),
        11:('第二节','RFP需求拆解','把长文档拆成可判定、可追踪、可分配的最小要求。'),
        19:('第三节','响应矩阵与参标判断','让每条要求都与企业真实能力、责任和动作发生连接。'),
        27:('第四节','受控长文档生成','先控制结构、材料与边界，再让AI生成章节。'),
        35:('第五节','一致性与风险检查','单章写得通不够，整份文档还要经得住交叉复核。'),
        42:('第六节','MVP产品设计','把方法收敛成一条可以在课堂上真实演示的产品链路。'),
        47:('第七节','场景迁移与课后作业','从标书迁移到合同、申报与安全问卷，再完成自己的V0.1。'),
    }
    for page in range(3,50):
        if page in dividers:
            section_divider(prs,page,*dividers[page])
        else:
            build_slide(prs,page)
    lint_layout(prs,strict=True)
    prs.save(str(OUT))
    declare_delivery(str(OUT),'presented')
    print(f'saved -> {OUT} | slides: {len(prs.slides)}')


if __name__=='__main__':
    main()
