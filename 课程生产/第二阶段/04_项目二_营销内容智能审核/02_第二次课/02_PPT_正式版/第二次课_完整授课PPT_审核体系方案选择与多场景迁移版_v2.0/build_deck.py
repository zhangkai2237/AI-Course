#!/usr/bin/env python3
"""Build the corrected 48-slide second-lesson deck.

Build shape: solo — subagent dispatch not authorized in this run.
Visual lock: follows lesson one (warm off-white, deep teal, restrained coral/gold).
Motion: static throughout.
"""
from pathlib import Path
import sys

SKILL = Path("/Users/keivn/.codex/skills/slide-maker")
sys.path.insert(0, str(SKILL / "scripts"))

import deckkit as dk  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402

OUTDIR = Path(__file__).resolve().parent
OUT = OUTDIR / "第二次课_完整授课PPT_审核体系方案选择与多场景迁移版_v2.0.pptx"

BG = RGBColor(0xF7, 0xF8, 0xF4)
PALE = RGBColor(0xE8, 0xF2, 0xEF)
TEAL = RGBColor(0x12, 0x6B, 0x63)
TEAL2 = RGBColor(0x18, 0x7A, 0x70)
CORAL = RGBColor(0xD8, 0x5C, 0x3D)
GOLD = RGBColor(0xDB, 0xAB, 0x4A)
INK = RGBColor(0x17, 0x30, 0x2F)
MUTE = RGBColor(0x61, 0x76, 0x73)
GRID = RGBColor(0xD3, 0xDF, 0xDA)
DARK = RGBColor(0x10, 0x2F, 0x30)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT_CORAL = RGBColor(0xF8, 0xE8, 0xE2)
SOFT_GOLD = RGBColor(0xF6, 0xEF, 0xDA)

dk.set_palette(deep=DARK, blue=TEAL, teal=TEAL2, magenta=CORAL, slate=INK,
               mute=MUTE, mono="Arial", font="Arial", display="Arial",
               eafont="Noto Sans CJK SC", eadisplay="Noto Sans CJK SC",
               accents=[TEAL, CORAL, GOLD, MUTE])
dk.set_ground(BG)

W, H = 10.0, 5.625
SECTIONS = ["4  审核体系", "5  方案选择", "7  场景迁移"]


def tx(s, x, y, w, h, value, size=17, color=INK, bold=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Noto Sans CJK SC"):
    return dk.text(s, x, y, w, h, [[(value, size, color, bold, False, font)]],
                   align=align, anchor=anchor, space_after=0, line_spacing=1.0)


def paras(s, x, y, w, h, values, size=12, color=INK, gap=5, bullet=False):
    rows = []
    for value in values:
        prefix = "•  " if bullet else ""
        rows.append([(prefix + value, size, color, False, False, "Noto Sans CJK SC")])
    return dk.text(s, x, y, w, h, rows, space_after=gap, line_spacing=1.0)


def nav(s, page, sec=1):
    dk.box(s, 0, 5.27, W, 0.01, fill=GRID, line=GRID)
    for i, (label, x) in enumerate(zip(SECTIONS, [0.38, 2.05, 3.72]), start=1):
        color = TEAL if i == sec else MUTE
        tx(s, x, 5.31, 1.55, 0.18, label, 7.8, color, i == sec, anchor=MSO_ANCHOR.MIDDLE)
    tx(s, 7.2, 5.31, 2.0, 0.18, "第二次课 · 项目二", 7.2, MUTE,
       align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    tx(s, 9.48, 5.31, 0.18, 0.18, str(page), 8, MUTE,
       align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, font="Arial")


def base(prs, title, kicker, page, sec, notes, accent=TEAL):
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=BG, line=BG)
    tx(s, 0.4, 0.27, 8.6, 0.22, kicker, 9.5, accent)
    tx(s, 0.4, 0.56, 9.1, 0.55, title, 24, INK, True)
    dk.register_mark(s, "arcs", corner="tr", color=PALE, size=0.75)
    nav(s, page, sec)
    dk.speaker_notes(s, notes)
    return s


def card(s, x, y, w, h, title, lines=(), fill=WHITE, title_c=TEAL, hub=False, num=None):
    bg = TEAL2 if hub else fill
    ink = WHITE if hub else INK
    dk.box(s, x, y, w, h, fill=bg, line=bg if hub else GRID, round=True)
    title_x = x + 0.18
    title_w = w - 0.36
    if num is not None:
        dk.box(s, x + 0.16, y + 0.15, 0.34, 0.34,
               fill=GOLD if hub else PALE, line=GOLD if hub else PALE, round=True)
        tx(s, x + 0.16, y + 0.145, 0.34, 0.34, str(num), 9.5, DARK, True,
           PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, "Arial")
        title_x = x + 0.60
        title_w = w - 0.76
    if lines:
        tx(s, title_x, y + 0.17, title_w, 0.34, title, 14, WHITE if hub else title_c, True)
        paras(s, x + 0.18, y + 0.64, w - 0.36, h - 0.78, list(lines), 10.8, ink, 4)
    else:
        tx(s, title_x, y + 0.06, title_w, h - 0.12, title, 13,
           WHITE if hub else title_c, True, anchor=MSO_ANCHOR.MIDDLE)


def takeaway(s, text, danger=False):
    fill = SOFT_CORAL if danger else PALE
    color = CORAL if danger else TEAL
    dk.box(s, 0.4, 4.58, 9.15, 0.48, fill=fill, line=fill, round=True)
    tx(s, 0.58, 4.68, 8.8, 0.25, text, 12.2, color, True, anchor=MSO_ANCHOR.MIDDLE)


def divider(prs, num, title, question, agenda, page, sec, notes):
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=DARK, line=DARK)
    tx(s, 0.42, 0.38, 4.8, 0.24, f"第二次课 / 第{num}节", 10, GOLD)
    tx(s, 0.42, 0.86, 4.5, 0.62, title, 34, WHITE, True)
    tx(s, 0.42, 1.76, 3.65, 1.1, question, 23, WHITE, True)
    dk.box(s, 4.45, 1.68, 0.015, 2.00, fill=MUTE, line=MUTE)
    for item, y in zip(agenda, [1.83, 2.52, 3.21]):
        tx(s, 4.73, y, 4.55, 0.42, item, 15, WHITE)
    dk.box(s, 0.42, 4.28, 9.12, 0.68, fill=TEAL2, line=TEAL2, round=True)
    tx(s, 0.64, 4.43, 8.68, 0.32, "本节重点  ·  " + agenda[-1], 13.3, WHITE,
       anchor=MSO_ANCHOR.MIDDLE)
    tx(s, 0.42, 5.25, 3.0, 0.18, "定义任务 → 选择能力 → 场景迁移", 7.5, GRID)
    tx(s, 9.45, 5.25, 0.2, 0.18, str(page), 8, GRID,
       align=PP_ALIGN.RIGHT, font="Arial")
    dk.speaker_notes(s, notes)
    return s


def flow(s, labels, subs, y=1.75, hub=-1):
    n = len(labels)
    gap = 0.25
    width = (9.05 - gap * (n - 1)) / n
    for i, (label, sub) in enumerate(zip(labels, subs)):
        x = 0.48 + i * (width + gap)
        card(s, x, y, width, 1.55, label, [sub], fill=PALE if i % 2 == 0 else WHITE,
             hub=i == hub, num=i + 1)
        if i < n - 1:
            dk.arrow(s, x + width + 0.04, y + 0.58, gap - 0.08, 0.22, color=TEAL)


def three_cards(prs, title, kicker, page, sec, items, note, bottom):
    s = base(prs, title, kicker, page, sec, note)
    fills = [PALE, WHITE, PALE]
    for i, ((head, lines), x) in enumerate(zip(items, [0.48, 3.57, 6.66])):
        card(s, x, 1.35, 2.86, 2.9, head, lines, fill=fills[i], num=i + 1)
    takeaway(s, bottom)
    return s


def case_slide(prs, title, page, label, scenario, rows, note, bottom):
    s = base(prs, title, f"第七节 · {label}", page, 3, note, accent=CORAL)
    dk.box(s, 0.48, 1.24, 3.05, 3.08, fill=DARK, line=DARK, round=True)
    tx(s, 0.68, 1.48, 2.65, 0.28, "业务情境", 14, GOLD, True)
    tx(s, 0.68, 1.94, 2.65, 2.05, scenario, 13.2, WHITE, True)
    y = 1.24
    row_pitch = 3.12 / len(rows)
    for i, (head, value) in enumerate(rows):
        fill = PALE if i % 2 == 0 else WHITE
        dk.box(s, 3.78, y, 5.73, row_pitch - 0.09, fill=fill, line=GRID)
        tx(s, 3.96, y + 0.08, 0.95, 0.28, head, 10.5, TEAL, True,
           anchor=MSO_ANCHOR.MIDDLE)
        tx(s, 4.98, y + 0.08, 4.32, 0.30, value, 10.3, INK,
           anchor=MSO_ANCHOR.MIDDLE)
        y += row_pitch
    takeaway(s, bottom)
    return s


def build(prs):
    # 1 Cover
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=BG, line=BG)
    dk.box(s, 0, 0, 0.16, 5.08, fill=TEAL, line=TEAL)
    tx(s, 0.64, 0.54, 5.0, 0.24, "第二阶段 · 项目二 · 第二次课", 10, TEAL)
    tx(s, 0.64, 1.20, 8.6, 1.2, "审核体系、方案选择\n与多场景迁移", 35, DARK, True)
    tx(s, 0.66, 2.88, 7.8, 0.42, "第四节 + 第五节 + 第七节", 18, MUTE)
    for i, (lab, sub, x) in enumerate([
        ("定义", "把要求拆成审核点", 0.75),
        ("分工", "规则·模型·人工", 3.72),
        ("迁移", "五个业务案例", 6.69),
    ], start=1):
        card(s, x, 3.66, 2.55, 1.18, lab, [sub], fill=PALE if i != 2 else WHITE, num=i)
    tx(s, 0.66, 5.16, 4.0, 0.18, "180分钟 · 项目深化课", 8.5, MUTE)
    dk.speaker_notes(s, "本课按第四节、第五节和第七节重新组合。前两节保持营销审核项目主线，第七节通过五个完整案例讲方法迁移，不安排分组练习。")

    # 2 Map
    s = base(prs, "今天只解决一条进阶路线", "全课地图", 2, 1,
             "先把审核要求定义清楚，再选择执行能力，最后到其他业务中重新走一遍产品判断。")
    flow(s, ["第四节", "第五节", "第七节"],
         ["审核点：究竟查什么", "方案选择：怎样查", "多场景：如何重新迁移"], y=1.65, hub=1)
    takeaway(s, "先定义任务，再选择能力；先理解业务，再迁移方法。")

    # 3 Hook
    s = base(prs, "“做一个AI审核”至少缺少四个答案", "开场判断", 3, 1,
             "让学员先说需要追问什么，再归纳为对象、依据、结果和失败。不要在开场讨论模型型号。")
    for i, (head, sub, x) in enumerate([
        ("审核什么", "对象与适用条件", 0.48), ("依据什么", "当前有效材料", 2.79),
        ("结果给谁", "角色与下一步", 5.10), ("失败怎么办", "重试与人工承接", 7.41)
    ], start=1):
        card(s, x, 1.55, 2.02, 2.5, head, [sub], fill=PALE if i % 2 else WHITE, num=i)
    takeaway(s, "笼统需求不能直接进入模型，也不能直接进入研发。", danger=True)

    # 4 Divider
    divider(prs, "四", "审核体系", "业务要求怎样变成\n可执行的检查？",
            ["审核点的六个要素", "三张核心审核点卡", "结果分类与处理动作"], 4, 1,
            "第四节只解决查什么。暂时不决定使用规则还是模型。")

    # 5-16 Section 4
    s = base(prs, "“准确、完整、合规”为什么不能直接实现", "第四节 · 问题起点", 5, 1,
             "用商品功能、价格、赠品三个追问说明：人能凭经验理解的要求，系统需要更明确的任务定义。")
    card(s, 0.52, 1.35, 2.55, 2.82, "商品功能准确", ["数字描述什么对象？", "全文语境是否改变含义？", "材料冲突时听谁的？"], fill=PALE)
    card(s, 3.44, 1.35, 2.55, 2.82, "价格条件完整", ["哪条优惠路径？", "适用于哪些人？", "条件是否必须在同一句？"], fill=WHITE)
    card(s, 6.36, 1.35, 2.55, 2.82, "赠品权益有效", ["初始规则是什么？", "当前还有名额吗？", "哪些内容仍在使用？"], fill=PALE)
    takeaway(s, "经验判断必须拆成团队能够共同确认的检查任务。")

    s = base(prs, "一条审核点由六个要素共同定义", "第四节 · 核心结构", 6, 1,
             "依次解释检查对象、适用条件、业务依据、判断方式、例外分支、处理动作。六项缺一都会产生不同理解。")
    items = ["检查对象", "适用条件", "业务依据", "判断方式", "例外分支", "处理动作"]
    subs = ["查哪个主张", "何时才执行", "使用哪份材料", "怎样形成结论", "缺项冲突如何处理", "结果交给谁"]
    for i, (lab, sub) in enumerate(zip(items, subs)):
        x = 0.48 + (i % 3) * 3.06
        y = 1.30 + (i // 3) * 1.42
        card(s, x, y, 2.80, 1.15, lab, [sub], fill=PALE if i % 2 == 0 else WHITE, num=i + 1)
    takeaway(s, "审核点是业务、产品、研发、测试共同使用的工作单元。")

    s = base(prs, "先用分类看范围，再向下拆成审核点", "第四节 · 范围地图", 7, 1,
             "分类帮助看到首版范围，但分类不是可执行规则。强调材料中有某项内容，不等于每篇文案都要写。")
    rows = [
        ["类别", "本案要回答的问题", "主要负责人"],
        ["商品事实", "数值描述哪个功能对象", "产品负责人"],
        ["价格条件", "金额怎样成立、适用于谁", "活动负责人"],
        ["赠品权益", "条件与当前状态是否有效", "活动负责人"],
        ["承诺范围", "表达是否超出材料支持", "产品 / 履约"],
        ["材料完整", "缺项时哪些判断不能继续", "提交人 / 负责人"],
    ]
    dk.table(s, 0.48, 1.22, 9.04, rows, col_w=[1.55, 5.25, 2.24], size=10.8, row_h=0.55,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "首版只做选定任务，不把所有审核要求一次装进系统。")

    s = base(prs, "R01 · 温度与功能对象一致性", "第四节 · 审核点卡", 8, 1,
             "原句是65℃热水直接拖地。逐项说明本审核点如何定义，不讨论技术实现。")
    labels = [
        ("检查", "温度描述哪个功能"), ("适用", "出现温度与功能表达"),
        ("材料", "文案、SKU 与商品资料"), ("判断", "理解对象后再比较"),
        ("例外", "无温度则不适用；对象不清待确认"), ("处理", "作者改、产品确认、审核复核")
    ]
    for i, (a, b) in enumerate(labels):
        x = 0.48 + (i % 3) * 3.06
        y = 1.27 + (i // 3) * 1.46
        card(s, x, y, 2.80, 1.18, a, [b], fill=PALE if i % 2 == 0 else WHITE, num=i + 1)
    takeaway(s, "不是检查数字是否出现，而是检查数字与功能对象是否一致。")

    s = base(prs, "四类样例共同限定R01的边界", "第四节 · 样例设计", 9, 1,
             "逐行比较正常、问题、待确认和不适用。只有错误句会让团队把规则理解得过窄。")
    rows = [
        ["表达", "结果状态", "为什么"],
        ["65℃热水直接拖地", "与材料不一致", "温度用于错误功能对象"],
        ["基站支持65℃热水洗拖布", "未发现该项问题", "与当前商品资料一致"],
        ["65℃热水深度清洁", "待确认", "对象不清，需要结合全文"],
        ["基站支持热水洗拖布", "本项不适用", "没有温度数值表达"],
    ]
    dk.table(s, 0.48, 1.30, 9.04, rows, col_w=[3.05, 2.10, 3.89], size=11.0, row_h=0.62,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "一条审核点至少准备正常、问题、不适用和待确认样例。")

    s = base(prs, "R02 · 价格数字成立，不代表文案成立", "第四节 · 审核点卡", 10, 1,
             "2399减400再减200等于1799 元，只完成金额路径。还要核对人群、两类券和表达范围。")
    tx(s, 0.58, 1.40, 8.9, 0.55, "2399  −  400  −  200  ＝  1799", 29, TEAL, True, PP_ALIGN.CENTER)
    for i, (lab, sub, x) in enumerate([
        ("金额路径", "程序可以核算", 0.72),
        ("资格条件", "新会员是否说明", 3.83),
        ("表达范围", "是否扩大为人人同价", 6.94),
    ], start=1):
        card(s, x, 2.48, 2.36, 1.48, lab, [sub], fill=PALE if i != 2 else WHITE, num=i)
    takeaway(s, "价格审核同时包含计算问题和语义条件问题。")

    s = base(prs, "三种价格表达，对应三种完整度", "第四节 · 条件关系", 11, 1,
             "比较人人同价、领券1799元、完整限定条件三种表达。说明完整句式只是课堂示例，不是唯一模板。")
    rows = [
        ["表达", "主要问题"],
        ["1799元，人人同价", "缺资格与用券条件，并扩大适用范围"],
        ["领券到手1799元", "仍缺新会员资格与两类券条件"],
        ["新会员使用两类指定券后1799元", "主要条件已说明，仍需结合全文与有效活动"],
    ]
    dk.table(s, 0.60, 1.46, 8.80, rows, col_w=[3.55, 5.25], size=11.5, row_h=0.72,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "系统核对语义和条件关系，不强制作者复制固定句式。")

    s = base(prs, "R03 · 赠品审核多了一项“当前状态”", "第四节 · 审核点卡", 12, 1,
             "初始前200名规则不能证明当前仍有名额。巡检时还需要权益变化、受影响范围和内容使用状态。")
    flow(s, ["初始规则", "当前状态", "影响范围", "处理动作"],
         ["前200名·每单一份", "名额是否已用完", "哪些内容仍在使用", "更新·说明·暂停"], y=1.55, hub=1)
    takeaway(s, "规则没有失效，但它可能已经不能支持“现在仍然赠送”。")

    s = base(prs, "同一审核点，在三个环节使用不同输入", "第四节 · 复用与变化", 13, 1,
             "审核要求可以复用，但自检、正式审核和巡检的输入不同。巡检多出变化事件和使用状态。")
    rows = [
        ["环节", "共同输入", "新增输入", "主要动作"],
        ["事前自检", "文案 + 当前材料", "提交上下文", "作者先修改"],
        ["事中审核", "文案 + 当前材料", "审核任务与历史意见", "审核员确认"],
        ["事后巡检", "审核点 + 业务材料", "变化事件 + 使用状态", "找到并处理在用内容"],
    ]
    dk.table(s, 0.48, 1.42, 9.04, rows, col_w=[1.45, 2.85, 2.55, 2.19], size=11.0, row_h=0.68,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "复用的是判断要求，不是把每次输入原样复制。")

    s = base(prs, "审核结果至少分成六种状态", "第四节 · 结果设计", 14, 1,
             "逐项说明状态和下一步，尤其区分业务未发现问题与技术执行未完成。")
    results = [
        ("未发现该项问题", "保留范围，不扩大结论"), ("与材料不一致", "作者修改或确认新材料"),
        ("表达缺少条件", "补条件后复核"), ("信息不足", "标明缺项与责任人"),
        ("本项不适用", "不要求补无关内容"), ("执行未完成", "重试或转人工")
    ]
    for i, (a, b) in enumerate(results):
        x = 0.48 + (i % 3) * 3.06
        y = 1.28 + (i // 3) * 1.43
        card(s, x, y, 2.80, 1.15, a, [b], fill=SOFT_CORAL if i == 5 else (PALE if i % 2 == 0 else WHITE),
             title_c=CORAL if i == 5 else TEAL, num=i + 1)
    takeaway(s, "系统没做完检查，不等于没有发现问题。", danger=True)

    s = base(prs, "审核点还要能够转成测试样本", "第四节 · 从定义到验证", 15, 1,
             "解释每种状态都需要对应样本，测试的不只是模型回答，也包括材料版本、适用性和后续处理。")
    flow(s, ["审核点定义", "样本组合", "预期状态", "处理动作"],
         ["六个要素", "正常·问题·边界", "六类结果", "修改·确认·重试"], y=1.60, hub=2)
    takeaway(s, "能写成预期结果的审核点，才真正进入可实现状态。")

    three_cards(prs, "第四节最终留下三张可执行的审核点卡", "第四节 · 小结", 16, 1,
                [("R01 商品功能", ["数字与功能对象一致"]),
                 ("R02 价格条件", ["金额、资格与范围一致"]),
                 ("R03 赠品权益", ["规则与当前状态一致"])],
                "用三张卡回看六个要素和六类结果。下一节不再讨论查什么，而讨论怎样查。",
                "先把任务定义清楚，再比较实现办法。")

    # 17 break
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=DARK, line=DARK)
    tx(s, 0.55, 1.35, 8.9, 0.55, "休息 10 分钟", 34, WHITE, True, PP_ALIGN.CENTER)
    tx(s, 0.55, 2.32, 8.9, 0.42, "回来以后：规则、模型与人工怎样分工", 18, GOLD, True, PP_ALIGN.CENTER)
    tx(s, 0.55, 3.20, 8.9, 0.35, "带着一个问题回来：哪一项工作一定不该交给模型？", 15, GRID, False, PP_ALIGN.CENTER)
    tx(s, 9.45, 5.25, 0.2, 0.18, "17", 8, GRID, align=PP_ALIGN.RIGHT, font="Arial")
    dk.speaker_notes(s, "休息十分钟。提醒学员回来后先回答屏幕上的问题。")

    # 18 Divider
    divider(prs, "五", "方案选择", "规则、模型与人工\n怎样分工？",
            ["四类候选方案", "三个审核点的能力分配", "交接、冲突与失败"], 18, 2,
            "第五节从审核要求进入执行方式。组合方案是候选，不是预设答案。")

    # 19-31 Section 5
    s = base(prs, "审核要求与执行方式，是两层不同设计", "第五节 · 两层结构", 19, 2,
             "第四节定义查什么，第五节定义怎样查。一个审核点可以拆给多种能力共同完成。")
    card(s, 0.72, 1.48, 3.55, 2.55, "业务层：查什么", ["检查对象", "适用条件", "材料与判断", "例外与处理"], fill=PALE, num="A")
    tx(s, 4.50, 2.30, 1.0, 0.5, "→", 32, TEAL, True, PP_ALIGN.CENTER)
    card(s, 5.72, 1.48, 3.55, 2.55, "能力层：怎样查", ["程序执行明确动作", "模型理解自然语言", "人工确认事实与责任"], fill=WHITE, num="B")
    takeaway(s, "一条审核点不必整条交给同一种能力。")

    s = base(prs, "混合审核不是三个方框，而是一套任务组合", "第五节 · 核心定义", 20, 2,
             "分别讲程序、模型、业务材料和人工。业务材料不是工具，却决定判断是否有依据。")
    for i, (head, sub, x) in enumerate([
        ("程序", "取材料·算金额·校验", 0.48),
        ("模型", "理解对象·主张·条件", 2.79),
        ("业务材料", "提供当前事实与版本", 5.10),
        ("人工", "确认争议·正式处置", 7.41)
    ], start=1):
        card(s, x, 1.52, 2.02, 2.52, head, [sub], fill=PALE if i % 2 else WHITE, num=i)
    takeaway(s, "组合价值来自清晰分工，不来自能力数量。")

    s = base(prs, "四类候选方案，没有天然赢家", "第五节 · 方案比较", 21, 2,
             "按本案比较材料清单、程序为主、模型为主和组合方案。强调简单办法若足够，就不必堆叠能力。")
    rows = [
        ["方案", "主要改善", "主要限制"],
        ["材料清单 + 人工", "减少找材料和漏项", "逐句理解仍由人完成"],
        ["程序为主", "稳定执行明确条件与计算", "自由表达维护成本高"],
        ["模型为主", "理解多种表达与上下文", "可能漏条件、误读或错引"],
        ["组合方案", "分担校验、理解与确认", "交接复杂，需验证整体投入"],
    ]
    dk.table(s, 0.48, 1.28, 9.04, rows, col_w=[1.70, 3.35, 3.99], size=10.8, row_h=0.61,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "AI项目也应允许“简单办法已经够用”这个答案。")

    s = base(prs, "方案选择看五个维度，不看技术名词多少", "第五节 · 选择标准", 22, 2,
             "用输入结构、稳定性、维护成本、纠错成本和责任要求判断每项任务适合什么能力。")
    labels = [("输入结构", "明确字段还是自由表达"), ("稳定性", "允许多大波动"),
              ("维护成本", "规则是否频繁变化"), ("纠错成本", "错一次要付出什么"),
              ("责任要求", "最终谁必须确认")]
    for i, ((a, b), x) in enumerate(zip(labels, [0.42, 2.33, 4.24, 6.15, 8.06])):
        card(s, x, 1.48, 1.68, 2.62, a, [b], fill=TEAL2 if i == 4 else (PALE if i % 2 == 0 else WHITE), hub=i == 4, num=i + 1)
    takeaway(s, "能力分工来自任务性质，而不是来自工具偏好。")

    s = base(prs, "R01分工：程序取材料，模型理解对象，人工确认争议", "第五节 · 商品功能", 23, 2,
             "沿65℃案例走完整条能力链。指出只搜索数字会同时命中文案与材料，却不能判断对象一致。")
    flow(s, ["绑定SKU", "程序取材料", "模型理解对象", "人工确认"],
         ["确认本次商品", "校验版本与引用", "拖地还是洗拖布", "处理对象不清与材料争议"], y=1.55, hub=2)
    takeaway(s, "语义理解之后仍需定位原句并校验材料。")

    s = base(prs, "R02分工：金额计算与条件理解要分开", "第五节 · 价格条件", 24, 2,
             "程序计算优惠路径，模型理解资格与条件关系，人工处理活动解释。两路结果合并，而不是互相覆盖。")
    card(s, 0.55, 1.40, 2.55, 2.82, "程序", ["取得有效活动", "计算优惠路径", "检查活动时间"], fill=PALE, num=1)
    card(s, 3.46, 1.40, 2.55, 2.82, "模型", ["识别宣传金额", "理解人群与条件", "结合全文判断约束"], fill=WHITE, num=2)
    card(s, 6.37, 1.40, 2.55, 2.82, "人工", ["确认业务解释", "处理材料冲突", "完成正式复核"], fill=PALE, num=3)
    takeaway(s, "算对金额，不等于说明了价格成立的条件。")

    s = base(prs, "R03分工：先选对范围，再识别承诺，最后确认处置", "第五节 · 赠品巡检", 25, 2,
             "沿名额用完事件讲分工。首版不让模型猜剩余名额，也不凭关键词自动下架。")
    flow(s, ["确认变化", "程序筛选范围", "模型识别承诺", "人工处置"],
         ["名额用完与生效条件", "商品·活动·使用状态", "是否仍在承诺赠送", "更新·说明·暂停"], y=1.55, hub=1)
    takeaway(s, "范围不完整时，不能宣称已经查全。", danger=True)

    s = base(prs, "三个审核点的首版能力矩阵", "第五节 · 分工总表", 26, 2,
             "横向比较三张审核点卡，帮助学员看到同一能力在不同任务中的位置。")
    rows = [
        ["审核点", "程序", "模型", "人工"],
        ["R01 商品功能", "取商品资料、校验版本", "理解温度对应对象", "确认商品争议"],
        ["R02 价格条件", "算金额、查时间", "理解资格与条件关系", "确认活动解释"],
        ["R03 赠品权益", "按关联与状态选范围", "识别当前赠送承诺", "确认范围与处置"],
    ]
    dk.table(s, 0.48, 1.40, 9.04, rows, col_w=[1.75, 2.45, 2.55, 2.29], size=10.6, row_h=0.69,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "分工的目标是减少总投入，同时保留必要责任。")

    s = base(prs, "混合审核真正难在交接，不在方框数量", "第五节 · 执行链", 27, 2,
             "每一步都要明确输入、输出、校验和失败去向。用五段链说明结果怎样进入人工处理。")
    flow(s, ["确认任务", "取得材料", "执行判断", "校验合并", "人工处理"],
         ["对象与适用条件", "版本与范围", "程序 + 模型", "定位·引用·冲突", "修改·确认·处置"], y=1.55, hub=3)
    takeaway(s, "交接设计决定组合方案能否真正省力。")

    three_cards(prs, "三类不能靠模型常识补齐的缺口", "第五节 · 事实边界", 28, 2,
                [("材料缺失", ["没有有效活动方案"]),
                 ("材料冲突", ["两份适用版本结论不同"]),
                 ("业务状态未知", ["库存、权益、使用状态不清"])],
                "强调未知就是未知。系统应当暴露缺口并交给明确角色，而不是选择更肯定的说法。",
                "事实缺失时，正确动作是停下并找负责人。")

    s = base(prs, "多路结果要合并，不是简单投票", "第五节 · 冲突处理", 29, 2,
             "金额成立与条件缺失可以同时成立；只有对同一事实的判断不同才是真冲突。")
    card(s, 0.60, 1.45, 3.55, 2.62, "并列问题", ["金额计算成立", "资格条件缺失", "合并后保留条件问题"], fill=PALE, num=1)
    card(s, 5.83, 1.45, 3.55, 2.62, "真实冲突", ["两路判断同一事实", "保留原句、材料与路径", "交人工确认"], fill=SOFT_CORAL, title_c=CORAL, num=2)
    tx(s, 4.32, 2.36, 1.36, 0.55, "≠", 34, TEAL, True, PP_ALIGN.CENTER)
    takeaway(s, "先判断结果是否在回答同一个问题，再决定是否冲突。")

    s = base(prs, "执行失败必须单独呈现", "第五节 · 失败交接", 30, 2,
             "超时、文件读取失败、引用校验失败和结果不可用都不是未发现问题。展示状态分流。")
    flow(s, ["调用开始", "局部失败", "保留结果", "重试 / 转人工"],
         ["记录任务与版本", "超时·读取失败·不可用", "不覆盖其他检查", "明确负责人和下一步"], y=1.55, hub=1)
    takeaway(s, "失败项不能被静默计入通过。", danger=True)

    s = base(prs, "模型选择要围绕任务表现，不围绕文风", "第五节 · 能力验证", 31, 2,
             "列出本项目应验证的六项表现。最后回到人工总投入：模型更强不代表流程更省力。")
    checks = ["功能对象是否识别正确", "资格与条件是否漏掉", "原句定位是否准确",
              "引用材料是否属于本任务", "不确定时能否交人工", "时延、成本与人工处理量"]
    for i, text in enumerate(checks):
        x = 0.48 + (i % 2) * 4.58
        y = 1.28 + (i // 2) * 0.92
        card(s, x, y, 4.25, 0.70, text, (), fill=PALE if i % 2 == 0 else WHITE, num=i + 1)
    takeaway(s, "最终比较的是有效问题与人工总投入。")

    s = base(prs, "第五节留下的不是固定架构，而是一套选择逻辑", "第五节 · 小结", 32, 2,
             "用四句话收束：规则确定动作，模型理解表达，材料提供事实，人工承担确认和处置。")
    for i, (head, sub, x) in enumerate([
        ("规则", "确定动作", 0.48), ("模型", "理解表达", 2.79),
        ("材料", "提供事实", 5.10), ("人工", "承担责任", 7.41)
    ], start=1):
        card(s, x, 1.48, 2.02, 2.60, head, [sub], fill=PALE if i % 2 else WHITE, num=i)
    takeaway(s, "输入、输出、校验、失败和承接，必须一起设计。")

    # 33 break
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=DARK, line=DARK)
    tx(s, 0.55, 1.32, 8.9, 0.55, "休息 10 分钟", 34, WHITE, True, PP_ALIGN.CENTER)
    tx(s, 0.55, 2.30, 8.9, 0.42, "下一部分：把同一方法迁移到五个业务案例", 18, GOLD, True, PP_ALIGN.CENTER)
    tx(s, 0.55, 3.18, 8.9, 0.35, "案例不是练习题，而是五次完整的产品推演", 15, GRID, False, PP_ALIGN.CENTER)
    tx(s, 9.45, 5.25, 0.2, 0.18, "33", 8, GRID, align=PP_ALIGN.RIGHT, font="Arial")
    dk.speaker_notes(s, "休息十分钟。第三部分全部由教师讲解，不组织分组练习。")

    # 34 Divider
    divider(prs, "七", "多场景迁移", "离开营销审核，\n方法还能成立吗？",
            ["合同审核母案例", "十一项迁移框架", "客服·招聘·报销·金融"], 34, 3,
            "第三部分用合同作为母案例，再连续展开四个新业务。强调重新完成业务分析，而不是替换页面名。")

    # 35-48 Section 7
    s = base(prs, "合同审核先从“AI法务”收窄到一个小场景", "第七节 · 合同母案例", 35, 3,
             "项目边界是常规业务合同的初步检查、定位和修改建议；复杂争议和最终决定仍由专业人员负责。", accent=CORAL)
    card(s, 0.62, 1.45, 3.55, 2.62, "不做", ["诉讼判断", "完整法律咨询", "自动作出最终决定"], fill=SOFT_CORAL, title_c=CORAL)
    card(s, 5.83, 1.45, 3.55, 2.62, "首版做", ["常规合同初步检查", "问题条款定位", "修改建议与待确认项"], fill=PALE)
    takeaway(s, "先选择能够验证的小场景，再讨论完整系统。")

    s = base(prs, "先理解法务原来检查什么", "第七节 · 人工工序", 36, 3,
             "这一页不谈模型。先还原主体、条款、引用、履约和周边风险五类人工任务。", accent=CORAL)
    labels = [("合同主体", "签约对象与异常"), ("合同条款", "付款·交付·验收·违约"),
              ("规范引用", "是否准确且仍适用"), ("履约方式", "是否真的能够执行"),
              ("周边风险", "履约历史与关系")]
    for i, ((a, b), x) in enumerate(zip(labels, [0.42, 2.33, 4.24, 6.15, 8.06])):
        card(s, x, 1.48, 1.68, 2.62, a, [b], fill=TEAL2 if i == 3 else (PALE if i % 2 == 0 else WHITE), hub=i == 3, num=i + 1)
    takeaway(s, "先把人类工作说清楚，再决定如何自动化。")

    s = base(prs, "从最终输出反推每一项输入", "第七节 · 输出与输入", 37, 3,
             "业务需要问题条款、类型、原因、影响、建议和待确认项，因此输入不能只有合同正文。", accent=CORAL)
    card(s, 0.58, 1.35, 3.70, 2.90, "最终输出", ["问题所在条款", "问题类型与原因", "业务影响", "修改建议", "人工待确认项"], fill=PALE, num=1)
    tx(s, 4.42, 2.43, 1.0, 0.5, "←", 32, TEAL, True, PP_ALIGN.CENTER)
    card(s, 5.72, 1.35, 3.70, 2.90, "所需输入", ["合同正文", "本次审核重点", "企业合同规则", "对方企业信息", "交易背景"], fill=WHITE, num=2)
    takeaway(s, "输出越具体，输入缺口越容易被发现。")

    s = base(prs, "先用脱敏合同验证基础能力", "第七节 · 快速验证", 38, 3,
             "验证重点不是文字流畅，而是定位、明显矛盾、建议可理解、不过度提示、不编造以及知道何时停下。", accent=CORAL)
    checks = ["定位真实条款", "发现明显矛盾", "建议可理解", "不过度提示", "不编造内容", "不确定时交人工"]
    for i, text in enumerate(checks):
        x = 0.48 + (i % 3) * 3.06
        y = 1.30 + (i // 3) * 1.42
        card(s, x, y, 2.80, 1.15, text, (), fill=PALE if i % 2 == 0 else WHITE, num=i + 1)
    takeaway(s, "模型可以像认真但缺少商业经验的法务实习生。")

    s = base(prs, "合同审核拆成三条并行路线", "第七节 · 合同工作流", 39, 3,
             "并行不是为了复杂，而是让每条路线目标、输入和结果都可单独检查。最后由法务确认。", accent=CORAL)
    flow(s, ["文件解析", "三路并行", "合并分级", "法务确认"],
         ["类型与可读性", "条款 / 引用 / 企业风险", "去重·冲突·待确认", "修改·审批·归档"], y=1.55, hub=1)
    takeaway(s, "多路能力最终仍要汇入一个人工负责的业务决定。")

    s = base(prs, "通用合同能力，要接入企业真实履约条件", "第七节 · 企业连接", 40, 3,
             "30天交付在条款上可能正常，结合产能、物料和历史履约后可能暴露真实风险。", accent=CORAL)
    card(s, 0.48, 1.35, 2.65, 2.95, "通用判断", ["条款结构", "常见问题", "规范引用", "文本修改建议"], fill=WHITE)
    card(s, 3.55, 1.35, 2.90, 2.95, "企业内部AI", ["订单", "库存与产能", "资金计划", "历史履约", "供应商关系"], hub=True)
    card(s, 6.87, 1.35, 2.65, 2.95, "新的业务结果", ["实际能否交付", "当前能否付款", "供应商是否可接受", "谁需要确认"], fill=PALE)
    takeaway(s, "企业产品的特殊价值，来自对这家公司当前业务的理解。")

    s = base(prs, "十一项迁移框架：完整迁移一次产品判断", "第七节 · 核心框架", 41, 3,
             "这是全课峰值页。按定义、构建、落地三组讲十一项，不要求学员当场练习。", accent=CORAL)
    groups = [
        ("A 定义", ["业务问题", "首版小场景", "人工工序", "最终输出"], 1.26, [1.95, 3.90, 5.85, 7.80], 1.70),
        ("B 构建", ["输入来源", "快速验证", "并行拆分", "边界分支"], 2.34, [1.95, 3.90, 5.85, 7.80], 1.70),
        ("C 落地", ["结果合并", "企业连接", "运行验证"], 3.42, [1.95, 4.55, 7.15], 2.20),
    ]
    idx = 1
    for label, arr, y, xs, width in groups:
        dk.box(s, 0.48, y, 1.18, 0.82, fill=TEAL2, line=TEAL2, round=True)
        tx(s, 0.48, y, 1.18, 0.80, label, 12, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        for item, x in zip(arr, xs):
            card(s, x, y, width, 0.82, item, (), fill=PALE if idx % 2 else WHITE, num=idx)
            idx += 1
    takeaway(s, "先定义决定，再设计能力，最后接入真实组织。")

    case_slide(prs, "案例一：售后退款会话质检", 42, "客服质检",
               "客服承诺“今天到账”，但前序口径是三个工作日；随后出现推诿。系统需要结合订单状态与服务政策判断。",
               [("输入", "会话、角色、订单退款状态、服务政策、质检规则"),
                ("并行", "流程规则 / 多轮语义 / 订单与政策核对"),
                ("输出", "问题位置、类型、等级、整改建议、待确认项"),
                ("边界", "转写失败、角色混淆、接口无返回、申诉"),
                ("人工", "质检主管复核；客服确认或申诉"),
                ("连接", "CRM 与订单、工单、知识库、培训系统")],
               "完整讲解客服质检。强调产品完成点是整改闭环，不是给会话打一个分。",
               "质检产品的终点是确认、申诉、整改与复训。")

    case_slide(prs, "案例二：岗位发布前的招聘信息审核", 43, "招聘审核",
               "岗位文案同时出现高薪承诺、性别年龄限制、本地户口、零经验与付费培训。系统要区分文本问题与企业事实。",
               [("输入", "招聘文案、企业、岗位编制、薪酬范围、渠道要求"),
                ("工序", "查企业、岗位、薪资、必填项、限制条件"),
                ("并行", "文案必填项 / 企业账号 / 岗位资料一致性"),
                ("输出", "缺项、问题位置、类型、修改要求、转人工"),
                ("边界", "岗位无法确认、资料冲突、异常引流"),
                ("连接", "招聘系统、岗位编制、薪酬与投诉记录")],
               "完整讲解招聘信息审核。强调岗位是否真实、薪酬是否批准必须来自企业内部事实。",
               "模型能看表达，企业系统才能确认岗位与承诺是否真实。")

    case_slide(prs, "案例三：员工费用报销初步审核", 44, "发票与报销",
               "周末高额休闲服务发票备注客户招待；金额在权限内，但没有OA申请，部门预算接近用完。多个单点正常不等于整体成立。",
               [("输入", "发票、报销单、员工部门、预算、合同、OA 与制度"),
                ("工序", "验票、核科目、核金额、查人员预算与审批"),
                ("并行", "票据验真 / 政策 / 预算 / 合同与审批"),
                ("输出", "可继续、待补材料、退回、等待预算、转人工"),
                ("边界", "接口失败、重复报销、业务矛盾、多路冲突"),
                ("连接", "财务、OA 与预算、合同和付款系统")],
               "完整讲解报销审核。强调发票真实不等于报销成立，首版不自动付款。",
               "多项判断需要合并，但最终处理仍由财务承接。")

    case_slide(prs, "案例四：理财产品营销材料审核", 45, "金融营销",
               "文案宣称“年化收益 6% / 本金安全 / 适合所有投资者”。产品要结合正式材料、风险提示、目标人群和渠道版本进行判断。",
               [("输入", "宣传材料、产品文件、风险提示、人群、渠道、版本"),
                ("工序", "查表达、产品事实、风险披露、适用人群、版本"),
                ("并行", "明确表达 / 产品事实 / 风险披露 / 渠道适配"),
                ("输出", "问题位置、缺失信息、等级、修改与人工结论"),
                ("边界", "材料冲突、人群不清、版本失效、结论冲突"),
                ("连接", "产品中心、合规、客户分层、渠道与版本管理")],
               "完整讲解金融营销材料审核。本案例只讲产品设计，不给出具体专业结论。",
               "高风险场景可由AI辅助定位与比对，最终责任必须明确到人。")

    s = base(prs, "五个案例放在一起，方法相同，业务边界不同", "第七节 · 跨场景比较", 46, 3,
             "横向比较被判断对象、关键材料、人工角色和企业连接价值。强调这些不是练习选项。", accent=CORAL)
    rows = [
        ["场景", "被判断对象", "关键材料", "最终人工角色", "企业连接价值"],
        ["合同", "条款", "模板·交易背景", "法务", "实际履约风险"],
        ["客服", "多轮会话", "政策·订单状态", "质检主管", "申诉与整改闭环"],
        ["招聘", "岗位文案", "编制·薪酬", "HR负责人", "岗位与承诺真实"],
        ["报销", "票据·报销单", "预算·OA·合同", "财务", "组合判断合理性"],
        ["金融营销", "宣传材料", "产品·风险·人群", "产品与合规", "版本与发布管理"],
    ]
    dk.table(s, 0.35, 1.20, 9.30, rows, col_w=[1.05, 1.55, 2.15, 1.65, 2.90], size=9.7, row_h=0.55,
             head_c=DARK, body_c=INK, rule_c=GRID, hi_fill=PALE, hi_c=TEAL)
    takeaway(s, "迁移不是换行业名称，而是重新完成一次完整产品判断。")

    s = base(prs, "真正可迁移的是八步主线", "第七节 · 方法收束", 47, 3,
             "用八步路线收束五个案例。指出变化的是对象、事实、角色与边界，不变的是产品方法。", accent=CORAL)
    flow(s, ["业务问题", "人工工序", "输出输入", "能力验证"],
         ["为何现在要解决", "人原来怎么做", "需要什么结果与材料", "基础能力是否可用"], y=1.35, hub=2)
    flow(s, ["并行流程", "异常承接", "企业连接", "运行验证"],
         ["任务怎样拆", "失败冲突交给谁", "接入真实系统", "决定继续或收缩"], y=3.18, hub=1)

    # 48 close
    s = dk.add_slide(prs)
    dk.box(s, 0, 0, W, H, fill=DARK, line=DARK)
    tx(s, 0.46, 0.35, 8.8, 0.25, "全课收束", 10, GOLD)
    tx(s, 0.46, 0.78, 8.9, 0.62, "专业性来自判断，不来自生成速度", 31, WHITE, True)
    cols = [
        ("第四节 · 定义", ["对象", "条件", "材料", "边界"]),
        ("第五节 · 分工", ["程序", "模型", "人工", "失败交接"]),
        ("第七节 · 迁移", ["重新理解业务", "重新设计流程", "重新确认责任"]),
    ]
    for i, ((head, items), x) in enumerate(zip(cols, [0.48, 3.60, 6.72])):
        fill = TEAL2 if i == 1 else RGBColor(0x16, 0x42, 0x42)
        dk.box(s, x, 1.72, 2.84, 2.55, fill=fill, line=fill, round=True)
        tx(s, x + 0.18, 1.94, 2.48, 0.32, head, 16, WHITE, True)
        paras(s, x + 0.18, 2.48, 2.48, 1.45, items, 12.5, WHITE, 7)
    tx(s, 0.48, 4.66, 9.0, 0.42,
       "把任务定义清楚 · 把能力边界分配合理 · 把失败与责任接住", 16.5, GOLD, True, PP_ALIGN.CENTER)
    tx(s, 0.48, 5.28, 2.5, 0.18, "第二次课 · 完", 8, GRID)
    tx(s, 9.46, 5.28, 0.2, 0.18, "48", 8, GRID, align=PP_ALIGN.RIGHT, font="Arial")
    dk.speaker_notes(s, "最后回扣三节关系。AI可以快速生成页面和方案草稿，但产品经理的专业性来自任务定义、能力边界、异常承接和跨业务重新判断。")


def main():
    OUTDIR.mkdir(parents=True, exist_ok=True)
    prs = dk.blank_deck()
    build(prs)
    dk.lint_layout(prs, strict=True)
    prs.save(OUT)
    dk.declare_delivery(str(OUT), "presented")
    print(f"saved -> {OUT} | slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
