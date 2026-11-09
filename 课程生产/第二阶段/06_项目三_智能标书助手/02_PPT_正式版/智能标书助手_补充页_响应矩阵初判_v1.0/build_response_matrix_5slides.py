#!/usr/bin/env python3
"""智能标书助手补充页：AI生成响应矩阵初步判断（5页）。"""
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
    blank_deck, add_slide, slide_background, box, text, arrow, table,
    flow_chain, speaker_notes, design_intent, lint_layout, declare_delivery,
    set_palette, set_geometry, set_ground, _flat,
)


OUT = ROOT / '智能标书助手_补充页_AI生成响应矩阵初判_5页_v1.0.pptx'

BG = RGBColor(244, 246, 248)
PAPER = RGBColor(255, 255, 255)
INK = RGBColor(24, 36, 51)
GREY = RGBColor(91, 106, 121)
MUTE = RGBColor(91, 106, 121)
LINE = RGBColor(218, 225, 231)
DEEP = RGBColor(13, 27, 42)
TEAL = RGBColor(19, 138, 126)
TEAL_DARK = RGBColor(13, 112, 101)
TEAL_SOFT = RGBColor(227, 243, 240)
BLUE = RGBColor(62, 111, 176)
BLUE_SOFT = RGBColor(233, 240, 250)
AMBER = RGBColor(201, 133, 23)
AMBER_DARK = RGBColor(142, 88, 10)
AMBER_SOFT = RGBColor(255, 243, 217)
CORAL = RGBColor(210, 90, 74)
CORAL_DARK = RGBColor(174, 58, 47)
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
        '说明：页面中的企业资料、要求编号与判断结果均为课堂模拟。'
    )


def chrome(slide, page, title, subtitle=''):
    slide_background(slide, BG)
    tx(slide, M, 0.34, 4.4, 0.20, '03 响应矩阵 · AI初步判断', 9, MUTE, True)
    tx(slide, M, 0.72, 8.55, 0.60, title, 24.5, INK, True)
    if subtitle:
        tx(slide, M, 1.38, 8.55, 0.24, subtitle, 11.5, GREY)
    rule(slide, 8.50, 0.43, 9.38, 0.43, LINE, 1.1)
    for x, (label, fill) in zip(
        [8.48, 8.90, 9.32], [('R', TEAL), ('M', BLUE), ('C', AMBER)]
    ):
        circle(slide, x, 0.33, 0.20, fill)
        tx(slide, x, 0.365, 0.20, 0.12, label, 5.7, PAPER, True,
           align=PP_ALIGN.CENTER, wrap=False)
    rule(slide, M, 5.18, 9.38, 5.18, LINE, 0.8)
    tx(slide, M, 5.28, 7.30, 0.16,
       '智能标书助手 · 响应矩阵初步判断', 7.2, MUTE)
    tx(slide, 8.75, 5.26, 0.63, 0.18, f'补充{page:02d}', 8, MUTE, True,
       align=PP_ALIGN.RIGHT)


def source_card(slide, x, y, title, body, color, fill):
    rounded(slide, x, y, 2.56, 0.84, fill, color)
    circle(slide, x + 0.18, y + 0.19, 0.44, color)
    tx(slide, x + 0.18, y + 0.31, 0.44, 0.14, title[0], 8.5, PAPER, True,
       align=PP_ALIGN.CENTER, wrap=False)
    tx(slide, x + 0.78, y + 0.16, 1.60, 0.22, title, 10.8, color, True)
    tx(slide, x + 0.78, y + 0.46, 1.60, 0.18, body, 9.2, GREY, True)


def status_card(slide, x, title, definition, trigger, color, fill):
    rounded(slide, x, 1.94, 2.03, 2.62, PAPER, color)
    circle(slide, x + 0.75, 2.16, 0.52, color)
    mark = {'可以响应': '✓', '部分响应': '◐', '暂不响应': '×', '待确认': '?'}[title]
    tx(slide, x + 0.75, 2.29, 0.52, 0.18, mark, 12, PAPER, True,
       align=PP_ALIGN.CENTER, wrap=False)
    tx(slide, x + 0.18, 2.90, 1.67, 0.26, title, 13.2,
       AMBER_DARK if color == AMBER else color, True, align=PP_ALIGN.CENTER)
    rounded(slide, x + 0.18, 3.34, 1.67, 0.56, fill, None)
    tx(slide, x + 0.31, 3.46, 1.41, 0.30, definition, 9.4, INK, True,
       align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    tx(slide, x + 0.25, 4.10, 1.53, 0.28, trigger, 8.8, GREY, True,
       align=PP_ALIGN.CENTER)


def slide_1(prs):
    """role=overview | form=input-match-output | static: standalone supplement | takeaway='AI把客户要求和企业能力逐条比对，先生成建议而不是最终承诺'"""
    s = add_slide(prs)
    chrome(s, 1, '响应矩阵的初步判断，AI怎么介入？',
           'AI的工作是逐条匹配、给出理由和暴露缺口；最终承诺仍由人确认。')

    rounded(s, 0.66, 1.94, 2.48, 2.50, PAPER, TEAL)
    tx(s, 0.92, 2.18, 1.96, 0.24, '客户要求', 13.5, TEAL_DARK, True,
       align=PP_ALIGN.CENTER)
    tx(s, 0.98, 2.70, 1.84, 1.08,
       'R-001  私有化部署\nR-002  国产数据库\nR-003  60天内交付',
       11.0, INK, True, align=PP_ALIGN.CENTER)
    tx(s, 0.98, 4.03, 1.84, 0.20, '来自要求清单', 8.8, MUTE, True,
       align=PP_ALIGN.CENTER)

    arrow(s, 3.28, 2.91, 0.42, 0.20, LINE)
    rounded(s, 3.82, 2.30, 2.38, 1.70, DEEP, None)
    circle(s, 4.73, 2.55, 0.56, BLUE)
    tx(s, 4.73, 2.69, 0.56, 0.18, 'AI', 9.5, PAPER, True,
       align=PP_ALIGN.CENTER, wrap=False)
    tx(s, 4.18, 3.34, 1.66, 0.26, '检索 · 匹配 · 判断', 12.3, PAPER, True,
       align=PP_ALIGN.CENTER)
    tx(s, 4.18, 3.67, 1.66, 0.16, '同时返回依据与缺口', 8.8, RGBColor(206, 218, 229), True,
       align=PP_ALIGN.CENTER)
    arrow(s, 6.32, 2.91, 0.42, 0.20, LINE)

    rounded(s, 6.86, 1.94, 2.48, 2.50, PAPER, BLUE)
    tx(s, 7.12, 2.18, 1.96, 0.24, '初步响应矩阵', 13.5, BLUE, True,
       align=PP_ALIGN.CENTER)
    tx(s, 7.18, 2.68, 1.84, 1.08,
       '可以响应\n部分响应\n暂不响应\n待确认',
       11.0, INK, True, align=PP_ALIGN.CENTER)
    tx(s, 7.18, 4.03, 1.84, 0.20, '等待责任人确认', 8.8, MUTE, True,
       align=PP_ALIGN.CENTER)

    rounded(s, 3.82, 4.28, 2.38, 0.44, BLUE_SOFT, None)
    tx(s, 3.82, 4.38, 2.38, 0.18, '企业产品、技术、案例与交付资料', 8.7, BLUE, True,
       align=PP_ALIGN.CENTER)
    note(s, 1,
         '这一页先给出整体逻辑。AI不是只看招标要求，也不是直接替企业承诺。它要同时看到客户要求和企业资料，逐条匹配后生成初步判断、引用材料与待确认问题。')
    design_intent(s, reason='两端对照，中间深色AI节点承担“匹配”这一核心动作。')


def slide_2(prs):
    """role=retrieval | form=flow-chain-plus-source-library | static: standalone supplement | takeaway='每条要求都要先检索企业材料，再让AI基于相关片段判断'"""
    s = add_slide(prs)
    chrome(s, 2, '先检索企业资料，再生成判断',
           '企业材料不应一次全部交给模型；系统要围绕每条要求寻找最相关的材料。')

    flow_chain(
        s, 0.72, 1.92, 8.56, 1.12,
        ['要求卡片', '检索企业资料', '提取相关片段', 'AI生成初判'],
        subs=['R-002 国产数据库', '产品 / 技术 / 案例', '保留原文与位置', '状态 + 理由 + 缺口'],
        accent=BLUE, hub_idx=3, hub_accent=TEAL_DARK,
    )

    source_card(s, 0.72, 3.46, '产品说明', '功能与部署方式', TEAL, TEAL_SOFT)
    source_card(s, 3.72, 3.46, '技术文档', '兼容性与架构说明', BLUE, BLUE_SOFT)
    source_card(s, 6.72, 3.46, '历史案例', '真实交付与客户场景', AMBER_DARK, AMBER_SOFT)
    rounded(s, 2.22, 4.50, 5.56, 0.38, DEEP, None)
    tx(s, 2.22, 4.58, 5.56, 0.20,
       '判断必须附带：引用材料 · 原文位置 · 不确定点',
       9.6, PAPER, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    note(s, 2,
         '这里强调检索增强的产品过程。针对每条招标要求，系统先在产品说明、技术文档、历史案例、资质和实施材料中检索，再把相关片段交给AI判断。这样判断才能回到企业材料，而不是靠模型常识猜。')


def slide_3(prs):
    """role=classification | form=four-status-board | static: standalone supplement | takeaway='AI只生成四类初判，并说明为什么这样判断'"""
    s = add_slide(prs)
    chrome(s, 3, 'AI生成四类“初步判断”',
           '状态不是结论本身；真正有价值的是每个状态背后的匹配依据与待办。')
    status_card(s, 0.62, '可以响应', '材料明确说明满足要求', '输出：对应材料与引用位置', TEAL_DARK, TEAL_SOFT)
    status_card(s, 2.87, '部分响应', '部分满足，或附带条件', '输出：已满足项与能力缺口', BLUE, BLUE_SOFT)
    status_card(s, 5.12, '暂不响应', '当前能力与要求有冲突', '输出：冲突点与影响范围', CORAL_DARK, CORAL_SOFT)
    status_card(s, 7.37, '待确认', '企业材料不足，无法可靠判断', '输出：待确认问题与责任人', AMBER_DARK, AMBER_SOFT)
    note(s, 3,
         '课堂上要提醒学员，不要把AI输出设计成简单的“是或否”。部分响应和待确认尤其重要，因为它们把模糊风险变成可以继续处理的问题。每个状态都要有依据、缺口和下一步。')
    design_intent(s, rhyme=1, reason='四类状态采用同一语法，颜色仅辅助，文字和符号共同编码。')


def slide_4(prs):
    """role=artifact | form=editable-table | static: standalone supplement | takeaway='初步响应矩阵把判断、依据、来源和待办放在同一行'"""
    s = add_slide(prs)
    chrome(s, 4, '最终形成一张可继续处理的响应矩阵',
           '每一行既回答“能不能响应”，也告诉团队“依据在哪里、下一步找谁”。')

    rows = [
        ['编号', '招标要求', '初步判断', '判断依据', '材料来源', '待处理事项'],
        ['R-001', '支持私有化部署', '可以响应', '支持本地服务器部署', '架构说明P12', '确认资源要求'],
        ['R-002', '兼容国产数据库', '部分响应', '已适配达梦；其他待核实', '兼容清单P6', '技术确认范围'],
        ['R-003', '60天完成交付', '待确认', '历史平均周期75天', '项目案例库', '评估人员投入'],
    ]
    table(
        s, 0.60, 1.94, 8.80, rows,
        col_w=[0.62, 1.72, 1.00, 2.06, 1.45, 1.95],
        header=True, highlight=1, size=8.8, row_h=0.66,
        head_c=DEEP, body_c=INK, rule_c=LINE,
        hi_fill=BLUE_SOFT, hi_c=BLUE, font=EA,
    )
    rounded(s, 1.02, 4.65, 7.96, 0.38, DEEP, None)
    tx(s, 1.02, 4.73, 7.96, 0.20,
       'AI不是只填“满足/不满足”，而是同时生成理由、引用与待办。',
       9.7, PAPER, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    note(s, 4,
         '这一页展示真正的产品产物。响应矩阵至少包含要求编号、要求内容、初步判断、判断依据、材料来源和待处理事项。这样它才能继续进入责任人确认、标书撰写和风险检查。')
    design_intent(s, reason='矩阵本身是页面主角，底部只保留一句阅读结论。')


def slide_5(prs):
    """role=governance | form=flow-chain-and-contrast | static: standalone supplement | takeaway='AI判断企业可能做到什么，人决定企业是否能够正式承诺'"""
    s = add_slide(prs)
    chrome(s, 5, '“能够做到”不等于“可以对外承诺”',
           '技术能力只是一个条件；正式响应还要考虑成本、资源、合同责任和交付风险。')

    flow_chain(
        s, 0.86, 1.94, 8.28, 1.34,
        ['AI初步判断', '责任人确认', '形成正式响应'],
        subs=['匹配材料，暴露缺口', '技术 / 商务 / 交付 / 法务', '写入标书并承担承诺'],
        accent=BLUE, hub_idx=1, hub_accent=AMBER,
    )

    rounded(s, 0.92, 3.58, 3.72, 1.02, BLUE_SOFT, BLUE)
    tx(s, 1.20, 3.79, 3.16, 0.22, 'AI回答：现有材料能否支持这条要求？',
       11.2, BLUE, True, align=PP_ALIGN.CENTER)
    tx(s, 1.20, 4.17, 3.16, 0.20, '输出建议、出处、缺口和待确认问题',
       9.2, GREY, True, align=PP_ALIGN.CENTER)

    rounded(s, 5.36, 3.58, 3.72, 1.02, AMBER_SOFT, AMBER)
    tx(s, 5.64, 3.79, 3.16, 0.22, '人回答：我们是否愿意承担这个承诺？',
       11.2, AMBER_DARK, True, align=PP_ALIGN.CENTER)
    tx(s, 5.64, 4.17, 3.16, 0.20, '综合人员、成本、周期与合同风险',
       9.2, GREY, True, align=PP_ALIGN.CENTER)

    note(s, 5,
         '最后用责任边界收束。AI可以判断材料是否支持、哪些地方有缺口，但正式承诺会形成合同责任，必须由技术、商务、项目和法务等责任人确认。技术上可能做得到，也不等于当前项目愿意以这个成本和周期承诺。')
    design_intent(s, reason='上方给出治理流程，下方用AI与人的两个问题形成清晰对照。')


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
    prs.core_properties.title = 'AI生成响应矩阵初步判断：5页补充课件'
    prs.core_properties.subject = '智能标书助手 · 响应矩阵补充页'
    prs.core_properties.author = 'AI Course'
    for builder in SLIDES:
        builder(prs)
    lint_layout(prs, strict=True)
    prs.save(str(OUT))
    declare_delivery(str(OUT), 'presented')
    print(f'saved -> {OUT} | slides: {len(prs.slides)}')


if __name__ == '__main__':
    main()
