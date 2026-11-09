#!/usr/bin/env python3
"""智能标书助手补充页：从招标文件生成要求清单（5页）。"""
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
    blank_deck, add_slide, box, text, arrow, speaker_notes,
    design_intent, tag_motif, lint_layout, declare_delivery,
    set_palette, set_geometry, set_ground, _flat,
)


OUT = ROOT / '智能标书助手_补充页_从招标文件生成要求清单_5页_v1.0.pptx'

BG = RGBColor(244, 246, 248)
PAPER = RGBColor(255, 255, 255)
INK = RGBColor(24, 36, 51)
GREY = RGBColor(91, 106, 121)
MUTE = RGBColor(91, 106, 121)
LINE = RGBColor(218, 225, 231)
DEEP = RGBColor(13, 27, 42)
TEAL = RGBColor(19, 138, 126)
TEAL_SOFT = RGBColor(227, 243, 240)
BLUE = RGBColor(62, 111, 176)
BLUE_SOFT = RGBColor(233, 240, 250)
AMBER = RGBColor(201, 133, 23)
AMBER_SOFT = RGBColor(255, 243, 217)
AMBER_TEXT = RGBColor(142, 88, 10)
CORAL = RGBColor(210, 90, 74)
CORAL_SOFT = RGBColor(253, 233, 230)

FONT = 'Noto Sans CJK SC'
EA = 'Noto Sans CJK SC'
W, H = 10.0, 5.625
M = 0.60


def rgb(value):
    return value if isinstance(value, RGBColor) else RGBColor(*value)


def tx(slide, x, y, w, h, value, size=15, color=INK, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, italic=False, wrap=True):
    runs = [[(line, size, rgb(color), bold, italic, EA)] for line in str(value).split('\n')]
    return text(slide, x, y, w, h, runs, align=align, anchor=valign,
                space_after=2, line_spacing=1.08, wrap=wrap)


def rounded(slide, x, y, w, h, fill=PAPER, line_c=LINE, lw=1.0):
    return box(slide, x, y, w, h, fill=rgb(fill),
               line=rgb(line_c) if line_c else None, line_w=lw, round=True)


def rule(slide, x1, y1, x2, y2, color=LINE, width=1.0):
    shp = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1), Inches(y1), Inches(x2), Inches(y2),
    )
    shp.line.color.rgb = rgb(color)
    shp.line.width = Pt(width)
    return _flat(shp)


def circle(slide, x, y, d, fill=TEAL, line_c=None, line_w=0.8):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d)
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    if line_c:
        shp.line.color.rgb = rgb(line_c)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    return _flat(shp)


def note(slide, page, body):
    speaker_notes(
        slide,
        f'补充页{page}讲解建议：\n{body}\n\n'
        '说明：页面中的示例字段、页码与招标要求均为课堂模拟。'
    )


def chrome(slide, page, title, subtitle=''):
    box(slide, 0, 0, W, H, fill=BG)
    tx(slide, M, 0.34, 3.8, 0.20, '02 RFP需求拆解 · 补充页', 9, MUTE, True)
    tx(slide, M, 0.72, 8.55, 0.60, title, 24.5, INK, True)
    if subtitle:
        tx(slide, M, 1.38, 8.55, 0.24, subtitle, 11.5, GREY)
    # 与主课件一致的 R/M/C 控制点母题
    rule(slide, 8.50, 0.43, 9.38, 0.43, LINE, 1.1)
    for x, (label, fill) in zip(
        [8.48, 8.90, 9.32], [('R', TEAL), ('M', BLUE), ('C', AMBER)]
    ):
        circle(slide, x, 0.33, 0.20, fill)
        tx(slide, x, 0.365, 0.20, 0.12, label, 5.7, PAPER, True,
           align=PP_ALIGN.CENTER, wrap=False)
    rule(slide, M, 5.18, 9.38, 5.18, LINE, 0.8)
    tx(slide, M, 5.28, 6.95, 0.16,
       '智能标书助手 · 从招标文件生成要求清单', 7.2, MUTE)
    tx(slide, 8.75, 5.26, 0.63, 0.18, f'补充{page:02d}', 8, MUTE, True,
       align=PP_ALIGN.RIGHT)


def step_box(slide, x, y, w, h, n, title, sub, color):
    rounded(slide, x, y, w, h, PAPER, color)
    cx = x + w / 2 - 0.16
    circle(slide, cx, y + 0.13, 0.32, color)
    tx(slide, cx, y + 0.21, 0.32, 0.12, str(n), 8.5, PAPER, True,
       align=PP_ALIGN.CENTER, wrap=False)
    tx(slide, x + 0.10, y + 0.57, w - 0.20, 0.22, title, 10.5, color, True,
       align=PP_ALIGN.CENTER)
    tx(slide, x + 0.10, y + 0.84, w - 0.20, 0.28, sub, 9.5, GREY,
       align=PP_ALIGN.CENTER)


def slide_1(prs):
    """role=overview | form=process | static: standalone supplement | takeaway='整份上传，不等于整份直接交给AI'"""
    s = add_slide(prs)
    chrome(s, 1, '整份上传，不等于整份直接交给AI',
           '用户只做一次上传，系统内部要把长文档拆成一组可控任务。')
    labels = [
        ('上传文件', 'PDF / Word等格式', TEAL),
        ('文档解析', '正文、表格、页码', BLUE),
        ('结构切分', '标题、章节、段落', AMBER),
        ('AI抽取', '识别并拆分要求', TEAL),
        ('合并校验', '去重、关联、核对', CORAL),
        ('要求清单', '卡片化、人工确认', BLUE),
    ]
    gap, w = 0.18, 1.34
    for i, (x, (title, sub, color)) in enumerate(zip(
        [0.56, 2.08, 3.60, 5.12, 6.64, 8.16], labels
    )):
        step_box(s, x, 2.06, w, 1.30, i + 1, title, sub, color)
        if i < len(labels) - 1:
            arrow(s, x + w + 0.035, 2.53, gap - 0.07, 0.18, LINE)
    rounded(s, 1.18, 3.90, 7.64, 0.72, DEEP, None)
    tx(s, 1.48, 4.09, 7.04, 0.30,
       '文件是整份上传的，但AI不是一次读完、一次作答。',
       16, PAPER, True, align=PP_ALIGN.CENTER)
    note(s, 1,
         '先回答学员最直接的问题：是整份上传，但不是整份原样塞进一次模型调用。'
         '产品内部要经过解析、切分、抽取、校验和人工确认，最后才得到正式要求清单。')
    design_intent(s, reason='六步链路是本组补充页的总图，先给出用户动作与系统动作之间的差异。')


def slide_2(prs):
    """role=parsing | form=input-transform-output | static: standalone supplement | takeaway='AI之前先恢复文档结构与来源位置'"""
    s = add_slide(prs)
    chrome(s, 2, '先把文件变成“有结构的内容”',
           '这一阶段的目标不是生成，而是把原始文件恢复成可定位、可切分的文档对象。')

    rounded(s, 0.62, 1.86, 2.12, 2.74, PAPER, LINE)
    tx(s, 0.88, 2.10, 1.60, 0.26, '输入文件', 13.5, RGBColor(13, 112, 101), True,
       align=PP_ALIGN.CENTER)
    for i, (y, (lab, col)) in enumerate(zip(
        [2.55, 3.10, 3.65], [('PDF', TEAL), ('Word', BLUE), ('扫描件', AMBER)]
    )):
        rounded(s, 0.95, y, 1.46, 0.44, [TEAL_SOFT, BLUE_SOFT, AMBER_SOFT][i], None)
        tx(s, 0.95, y + 0.07, 1.46, 0.25, lab, 10.5,
           AMBER_TEXT if lab == '扫描件' else col, True,
           align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    arrow(s, 2.92, 2.85, 0.44, 0.22, LINE)
    rounded(s, 3.48, 1.86, 2.86, 2.74, PAPER, BLUE)
    tx(s, 3.78, 2.10, 2.26, 0.26, '文档解析层', 13.5, BLUE, True,
       align=PP_ALIGN.CENTER)
    items = ['提取正文文字', '扫描件先做OCR', '识别标题与目录层级', '提取表格与单元格', '保留页码和段落位置']
    for y, item in zip([2.48, 2.85, 3.22, 3.59, 3.96], items):
        circle(s, 3.82, y + 0.01, 0.16, BLUE)
        tx(s, 4.13, y - 0.01, 1.78, 0.22, item, 10.5, INK, True)

    arrow(s, 6.52, 2.85, 0.44, 0.22, LINE)
    rounded(s, 7.08, 1.86, 2.30, 2.74, RGBColor(251, 250, 246), AMBER)
    tx(s, 7.34, 2.10, 1.78, 0.26, '结构化片段', 13.5, AMBER_TEXT, True,
       align=PP_ALIGN.CENTER)
    tx(s, 7.36, 2.52, 1.72, 1.24,
       '第42页\n第三章 项目需求\n3.2 系统性能要求\n系统应支持不少于\n5,000个用户同时在线',
       10.8, INK, True)
    rounded(s, 7.35, 4.04, 1.76, 0.34, AMBER_SOFT, None)
    tx(s, 7.35, 4.10, 1.76, 0.20, '原文位置被保留下来', 9.0, AMBER_TEXT, True,
       align=PP_ALIGN.CENTER)
    note(s, 2,
         '强调AI调用之前还有一层文档工程。系统不能只得到一大段纯文本，必须保留标题层级、表格、页码和原文位置。'
         '否则后面虽然能生成一句要求，却无法证明它来自哪里。')


def slide_3(prs):
    """role=extraction | form=chapter-map-and-split | static: standalone supplement | takeaway='按章节分段，AI逐段识别并拆开复合要求'"""
    s = add_slide(prs)
    chrome(s, 3, '按照章节分段，再让AI逐段抽取',
           '切分不是简单“每1000字一段”，而是尽量尊重文档原有结构。')

    sections = [
        ('资格要求', '资格条件', TEAL),
        ('技术需求', '功能与性能', BLUE),
        ('实施要求', '周期与交付', AMBER),
        ('服务要求', '售后与响应', CORAL),
        ('评分标准', '加分条件', TEAL),
    ]
    for x, (title, sub, color) in zip(
        [0.58, 2.40, 4.22, 6.04, 7.86], sections
    ):
        rounded(s, x, 1.88, 1.60, 0.98, PAPER, color)
        tx(s, x + 0.10, 2.08, 1.40, 0.22, title, 10.5,
           AMBER_TEXT if color == AMBER else color, True,
           align=PP_ALIGN.CENTER)
        tx(s, x + 0.10, 2.45, 1.40, 0.18, sub, 9.5, GREY,
           align=PP_ALIGN.CENTER)
    tx(s, 0.72, 3.07, 1.62, 0.22, '复合原文', 10, MUTE, True)
    rounded(s, 0.72, 3.39, 8.56, 0.58, PAPER, LINE)
    tx(s, 0.96, 3.55, 8.08, 0.25,
       '系统须支持私有化部署，兼容国产操作系统和数据库，并在60日内完成交付。',
       12.5, INK, True, align=PP_ALIGN.CENTER)
    arrow(s, 4.80, 4.05, 0.22, 0.34, TEAL, direction='down')
    outputs = [
        ('R-003A', '支持私有化部署', TEAL),
        ('R-003B', '兼容国产OS与数据库', BLUE),
        ('R-004', '60日内完成交付', CORAL),
    ]
    for i, (x, (rid, body, color)) in enumerate(zip(
        [0.94, 3.96, 6.98], outputs
    )):
        rounded(s, x, 4.40, 2.68, 0.55, [TEAL_SOFT, BLUE_SOFT, CORAL_SOFT][i], color)
        tx(s, x + 0.14, 4.55, 0.68, 0.20, rid, 8.5, color, True)
        tx(s, x + 0.86, 4.53, 1.66, 0.22, body, 9.4, INK, True)
    note(s, 3,
         '先按资格、技术、实施、服务和评分等章节切分，再让AI处理一个相对完整的语义片段。'
         '提示词要求它抽取必须满足、建议满足和影响评分的独立要求，并把可能出现不同状态的条件拆开。')


def slide_4(prs):
    """role=schema | form=requirement-card | static: standalone supplement | takeaway='抽取结果应成为可追溯的要求卡片'"""
    s = add_slide(prs)
    chrome(s, 4, '抽取结果不是摘要，而是一张“要求卡片”',
           '要求卡片既能被人阅读，也能进入响应矩阵、章节映射和风险检查。')

    rounded(s, 0.66, 1.84, 6.40, 2.97, PAPER, TEAL)
    tx(s, 0.94, 2.08, 1.12, 0.24, 'R-003B', 14, RGBColor(13, 112, 101), True)
    rounded(s, 5.80, 2.06, 0.94, 0.34, TEAL_SOFT, None)
    tx(s, 5.80, 2.15, 0.94, 0.15, '必须满足', 8.8, RGBColor(13, 112, 101), True,
       align=PP_ALIGN.CENTER)
    tx(s, 0.94, 2.53, 5.65, 0.38, '系统支持不少于5,000个并发用户', 17, INK, True)

    fields = [
        ('分类', '技术要求'), ('指标', '≥5,000并发'),
        ('来源章节', '3.2 系统性能要求'), ('来源页码', '第42页'),
        ('原始条款', '系统应支持不少于5,000个用户同时在线'), ('状态', '待企业材料确认'),
    ]
    field_positions = [
        (0.96, 3.16), (3.96, 3.16),
        (0.96, 3.70), (3.96, 3.70),
        (0.96, 4.24), (3.96, 4.24),
    ]
    for (x, y), (label, value) in zip(field_positions, fields):
        tx(s, x, y, 0.68, 0.16, label, 8.8, MUTE, True)
        tx(s, x + 0.78, y - 0.02, 2.02, 0.25, value, 10.5, INK, True)

    rounded(s, 7.40, 1.84, 1.94, 2.97, DEEP, None)
    tx(s, 7.63, 2.13, 1.48, 0.24, '它不只是摘要', 12.5, PAPER, True,
       align=PP_ALIGN.CENTER)
    tx(s, 7.67, 2.68, 1.40, 1.40,
       '摘要｜快速理解\n要求卡｜逐项判定\n来源字段｜回到原文\n结构字段｜进入后续',
       9.8, PAPER, True, align=PP_ALIGN.CENTER)
    note(s, 4,
         '展示产品最终中间产物。要求卡片至少要保存编号、要求、分类、强制性、指标、来源章节、页码、原文和当前状态。'
         '摘要是为了读懂，要求卡片是为了控制和追踪。')
    design_intent(s, weight='left', reason='大要求卡片是页面主角，右侧深色窄栏只负责解释它与摘要的区别。')


def slide_5(prs):
    """role=governance | form=three-way-division | static: standalone supplement | takeaway='AI、规则和人工共同保证不漏、不错、可追溯'"""
    s = add_slide(prs)
    chrome(s, 5, '为什么一次AI调用不够？',
           '正式产品把语义理解、确定性核对和责任确认分给不同机制。')

    lanes = [
        ('AI', '理解与抽取', '识别要求\n拆分复合条件\n关联相似条款', BLUE),
        ('规则', '确定性校验', '数字与单位\n日期与金额\n“必须/不得”扫描', TEAL),
        ('人工', '最终确认', '判断是否遗漏\n处理冲突与例外\n确认后进入矩阵', AMBER),
    ]
    for x, (name, role, body, color) in zip([0.68, 3.70, 6.72], lanes):
        rounded(s, x, 1.88, 2.72, 2.50, PAPER, color)
        circle(s, x + 1.08, 2.11, 0.56, color)
        tx(s, x + 1.08, 2.25, 0.56, 0.18, name, 9.5, PAPER, True,
           align=PP_ALIGN.CENTER, wrap=False)
        tx(s, x + 0.25, 2.94, 2.22, 0.25, role, 13.5,
           AMBER_TEXT if color == AMBER else (RGBColor(13, 112, 101) if color == TEAL else color), True,
           align=PP_ALIGN.CENTER)
        tx(s, x + 0.34, 3.35, 2.04, 0.72, body, 11.2, INK, True,
           align=PP_ALIGN.CENTER)

    rounded(s, 1.18, 4.62, 7.64, 0.40, DEEP, None)
    tx(s, 1.18, 4.69, 7.64, 0.25,
       '产品目标：不漏要求 · 不改数字 · 每条都能回到原文',
       10.5, PAPER, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    note(s, 5,
         '最后用三方分工收束。AI擅长理解和抽取，规则擅长数字、日期、单位和关键词等确定性校验，人工负责冲突、例外和最终责任。'
         '课堂上可以用“不漏、不错、找得到原文”概括这套机制。')


SLIDES = [slide_1, slide_2, slide_3, slide_4, slide_5]


def main():
    set_palette(
        deep=DEEP, blue=BLUE, teal=TEAL, magenta=CORAL, slate=GREY,
        mute=MUTE, mono=INK, font=FONT, display=FONT, eafont=EA,
        eadisplay=EA, accents=[TEAL, BLUE, AMBER, CORAL],
    )
    set_geometry(radius=0.11, rule_w=1.0)
    set_ground(BG)
    prs = blank_deck(W, H)
    prs.core_properties.title = '从招标文件生成要求清单：5页补充课件'
    prs.core_properties.subject = '智能标书助手 · RFP需求拆解补充页'
    prs.core_properties.author = 'AI Course'
    for builder in SLIDES:
        builder(prs)
    lint_layout(prs, strict=True)
    prs.save(str(OUT))
    declare_delivery(str(OUT), 'presented')
    print(f'saved -> {OUT} | slides: {len(prs.slides)}')


if __name__ == '__main__':
    main()
