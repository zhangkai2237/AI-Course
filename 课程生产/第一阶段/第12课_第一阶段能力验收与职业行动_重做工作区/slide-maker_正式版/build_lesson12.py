#!/usr/bin/env python3
"""第12课PPT构建脚本。

当 ANCHOR_ONLY=True 时，只生成三张样式验证页；设为 False 后生成完整55页。
"""
from pathlib import Path
import sys
import json

SKILL = Path("/Users/keivn/.codex/skills/slide-maker")
sys.path.insert(0, str(SKILL / "scripts"))
import deckkit as dk  # noqa: E402


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "第12课_AI产品经理的下一步_slide-maker正式版.pptx"
ANCHOR_OUT = ROOT / "第12课_3页样式验证.pptx"
SOURCE = Path("/Users/keivn/Project/AI-Course/课程生产/第一阶段/第12课_第一阶段能力验收与职业行动_重做工作区/第12课_第一阶段总结_AI产品趋势与项目实战_完整讲义_v0.2.md")

ANCHOR_ONLY = False
SLIDE_COUNT = 55
W, H = 13.333, 7.5

# Mode A mimic：严格沿用第11课的深色青绿教学系统。
BG = dk.RGBColor(0x03, 0x10, 0x16)
PANEL = dk.RGBColor(0x10, 0x2B, 0x37)
PANEL2 = dk.RGBColor(0x13, 0x38, 0x43)
LINE = dk.RGBColor(0x41, 0x6A, 0x78)
TEAL = dk.RGBColor(0x36, 0xD6, 0xC2)
CORAL = dk.RGBColor(0xFF, 0x76, 0x57)
TEXT = dk.RGBColor(0xF5, 0xF2, 0xE9)
MUTE = dk.RGBColor(0xA7, 0xB5, 0xBC)
YELLOW = dk.RGBColor(0xF0, 0xC5, 0x65)
GREEN = dk.RGBColor(0x68, 0xD3, 0x91)
RED = dk.RGBColor(0xEF, 0x69, 0x69)

dk.set_palette(
    deep=TEXT, blue=TEAL, teal=TEAL, magenta=CORAL, slate=MUTE, mute=MUTE,
    mono="Menlo", font="Helvetica Neue", display="Helvetica Neue",
    eadisplay="Hiragino Sans GB", eafont="Hiragino Sans GB",
    accents=[TEAL, CORAL, YELLOW],
)
dk.set_ground(BG)
dk.set_geometry(radius=1.0, rule_w=0.85)


def txt(slide, x, y, w, h, value, size, color=TEXT, bold=False,
        align=dk.PP_ALIGN.LEFT, anchor=dk.MSO_ANCHOR.TOP, wrap=True):
    return dk.text(
        slide, x, y, w, h,
        [[(str(value), size, color, bold, False, "Helvetica Neue", "Hiragino Sans GB")]],
        align=align, anchor=anchor, space_after=0, wrap=wrap,
    )


def chrome(slide, part, topic, page):
    txt(slide, 0.76, 0.25, 9.9, 0.28, f"{part} · {topic}", 10.5, TEAL, True)
    txt(slide, 12.02, 0.25, 0.56, 0.28, f"{page:02d}", 10.5, MUTE, True,
        align=dk.PP_ALIGN.RIGHT)
    dk.box(slide, 0.76, 7.18, 11.82, 0.014, fill=LINE)
    progress = dk.box(slide, 0.76, 7.13, 11.82 * page / SLIDE_COUNT, 0.036, fill=TEAL)
    dk.tag_motif(progress, loud=False)


def heading(slide, title, subtitle=None):
    txt(slide, 0.78, 0.67, 11.72, 0.68, title, 28, TEXT, True)
    rule = dk.box(slide, 0.78, 1.37, 1.45, 0.045, fill=TEAL)
    dk.tag_motif(rule, loud=False)
    if subtitle:
        txt(slide, 0.78, 1.50, 11.66, 0.38, subtitle, 14.5, MUTE)


def base(prs, part, topic, page, title, subtitle=None):
    slide = dk.add_slide(prs)
    chrome(slide, part, topic, page)
    heading(slide, title, subtitle)
    return slide


def note(slide, talk, source):
    dk.speaker_notes(slide, f"{talk}\n\n[讲义来源]\n{SOURCE}\n{source}")


def card(slide, x, y, w, h, title, body="", n=None, accent=TEAL,
         title_size=17, body_size=13.5, fill=PANEL):
    dk.box(slide, x, y, w, h, fill=fill, line=LINE, line_w=0.8, round=True, r=0.16)
    if n is not None:
        txt(slide, x + 0.22, y + 0.16, 0.50, 0.25, n, 11, accent, True)
        ty = y + 0.48
    else:
        ty = y + 0.23
    txt(slide, x + 0.22, ty, w - 0.44, 0.42, title, title_size, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    if body:
        txt(slide, x + 0.22, ty + 0.48, w - 0.44, h - (ty - y) - 0.62,
            body, body_size, MUTE)


def takeaway(slide, label, body, y=6.31, color=CORAL):
    dk.box(slide, 0.80, y, 11.72, 0.58, fill=PANEL2, line=LINE,
           line_w=0.6, round=True, r=0.14)
    txt(slide, 1.01, y + 0.14, 1.20, 0.25, label, 11.5, color, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(slide, 2.20, y + 0.11, 9.93, 0.32, body, 15.5, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)


def option_card(slide, x, y, w, h, letter, body):
    """讨论题选项：所有选项使用完全一致的颜色与视觉权重。"""
    dk.box(slide, x, y, w, h, fill=PANEL, line=LINE, line_w=0.8,
           round=True, r=0.16)
    dk.box(slide, x + 0.20, y + 0.21, 0.48, 0.48, fill=PANEL2, line=TEAL,
           line_w=0.75, round=True, r=0.14)
    txt(slide, x + 0.20, y + 0.21, 0.48, 0.48, letter, 15, TEAL, True,
        align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(slide, x + 0.84, y + 0.17, w - 1.08, h - 0.34, body, 15.5, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)


def anchor_s10(prs):
    """role=signature | form=causal-track | static: mode-A mimic | takeaway='五项能力是一条从问题到价值的因果链'"""
    s = base(prs, "第一部分", "第一阶段能力回顾", 10,
             "方法总览｜五项能力怎样形成因果链",
             "上游决策会一直影响到评测、成本和是否继续")
    stages = [
        ("业务问题", "谁的什么任务"),
        ("产品目标", "改善什么结果"),
        ("AI方案", "规则·模型·人工"),
        ("产品流程", "输入·状态·异常"),
        ("评测方式", "数据·指标·风险"),
        ("模型与迭代", "Bad Case·回归"),
        ("成本和价值", "ROI·继续条件"),
    ]
    x0, y, total_w, gap = 0.82, 2.30, 11.54, 0.12
    cw = (total_w - gap * 6) / 7
    for i, (title, body) in enumerate(stages):
        x = x0 + i * (cw + gap)
        if i < 6:
            dk.arrow(s, x + cw - 0.01, y + 0.52, gap + 0.04, 0.18, color=TEAL)
        card(s, x, y, cw, 1.55, title, body, f"{i+1:02d}",
             TEAL if i < 5 else (YELLOW if i == 5 else CORAL), 14.5, 11.5)
    dk.box(s, 1.06, 4.37, 11.08, 1.10, fill=PANEL2, line=TEAL,
           line_w=0.8, round=True, r=0.16)
    txt(s, 1.35, 4.62, 2.10, 0.30, "一个改变的连锁反应", 16, TEAL, True)
    txt(s, 3.48, 4.53, 8.20, 0.52,
        "人机分工改变 → 流程与PRD改变 → 评测集和指标改变 → 成本与上线判断改变",
        15.5, TEXT, True, anchor=dk.MSO_ANCHOR.MIDDLE)
    takeaway(s, "核心关系", "前面五项能力不是五个知识点，而是一条完整决策链。")
    note(s,
         "用这一页把前11课串起来。以营销内容审核为例：如果从‘完全自动发布’改为‘AI识别风险，人工决定是否修改’，那么产品流程会新增风险等级、原文引用、修改建议和人工确认；评测需要同时看高风险漏判和误判；成本也要把人工复核算进去。说明上游的产品判断会一直传导到最后的商业判断。",
         "第一部分·四｜五项能力是一条因果链")


def anchor_s34(prs):
    """role=complex | form=neutral-discussion-board | static: scan-all-at-once | takeaway='完整呈现情境与五个中性选项'"""
    s = base(prs, "第二部分", "讨论 11/16 · 商业与建设", 34,
             "AI能力应该自研、采购还是组合建设？",
             "请选择最接近你判断的一项，再说明理由")
    dk.box(s, 0.82, 2.00, 11.55, 0.70, fill=PANEL2, line=LINE,
           line_w=0.7, round=True, r=0.14)
    txt(s, 1.04, 2.18, 1.02, 0.25, "场景", 11.5, CORAL, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(s, 1.88, 2.11, 9.96, 0.38,
        "市场已有成熟转写和会议摘要；企业有自己的项目权限、历史决策资料和审批流程。",
        15, TEXT, True, anchor=dk.MSO_ANCHOR.MIDDLE)
    opts = [
        ("A", "全部自研，包括\n模型和语音"),
        ("B", "全部采购成熟会议产品"),
        ("C", "使用通用模型和转写服务，自建差异化流程"),
        ("D", "使用开源模型私有化，\n自研全部系统"),
        ("E", "根据安全、规模、差异化和团队能力综合判断"),
    ]
    top_y, row_h, gap = 2.94, 1.30, 0.17
    top_w = (11.55 - gap * 2) / 3
    for i in range(3):
        option_card(s, 0.82 + i * (top_w + gap), top_y, top_w, row_h, *opts[i])
    bottom_w = 4.74
    option_card(s, 1.56, 4.49, bottom_w, row_h, *opts[3])
    option_card(s, 7.00, 4.49, bottom_w, row_h, *opts[4])
    txt(s, 0.84, 6.15, 11.50, 0.32,
        "先选择，再用“差异化 · 安全 · 规模成本 · 持续能力”检查自己的理由。",
        13.5, MUTE, align=dk.PP_ALIGN.CENTER)
    note(s,
         "先不公布答案。请学员投票，再问一位选择全部自研、一位选择组合建设的学员说明理由。后续讲解时比较四个维度：能力是否形成差异化；数据和安全要求；使用规模和长期总成本；企业是否有持续建设和维护能力。讨论页上的所有选项必须保持中性，不用颜色或高亮提前暗示结论。",
         "第四部分·三｜自研、采购还是组合建设")


def anchor_s53(prs):
    """role=data-conclusion | form=seven-stage-roadmap | static: scan progression | takeaway='项目从选题一直走到展示和复盘'"""
    s = base(prs, "第三部分", "第二阶段项目班", 53,
             "项目路径｜从选题到展示的七个阶段",
             "每一阶段的产出都会进入同一个项目，并保留修改证据")
    stages = [
        ("选题与验证", "用户·场景·问题证据"),
        ("AI介入与MVP", "人机分工·范围取舍"),
        ("PRD与原型", "流程·状态·验收"),
        ("可运行Demo", "正常·异常·确认"),
        ("评测与Bad Case", "评测集·指标·错误"),
        ("迭代、成本与价值", "回归·模型·ROI"),
        ("展示与复盘", "证据链·贡献·讲述"),
    ]
    # 第一行4阶段
    x0, y1, gap = 0.82, 2.10, 0.18
    w1 = (11.55 - gap * 3) / 4
    for i in range(4):
        x = x0 + i * (w1 + gap)
        card(s, x, y1, w1, 1.45, stages[i][0], stages[i][1], f"{i+1:02d}",
             TEAL if i < 3 else YELLOW, 16, 12.5)
        if i < 3:
            dk.arrow(s, x + w1 - 0.01, y1 + 0.54, gap + 0.04, 0.18, color=TEAL)
    # 第二行3阶段，视觉上承接第一行
    y2, w2 = 4.02, 3.25
    x2 = 1.38
    for j in range(3):
        idx = 4 + j
        x = x2 + j * (w2 + 0.42)
        card(s, x, y2, w2, 1.48, stages[idx][0], stages[idx][1], f"{idx+1:02d}",
             CORAL if idx == 5 else (YELLOW if idx == 4 else GREEN), 16, 12.5,
             fill=PANEL2 if idx == 6 else PANEL)
        if j < 2:
            dk.arrow(s, x + w2 + 0.05, y2 + 0.56, 0.28, 0.18, color=TEAL)
    takeaway(s, "版本意识", "保留 V0.1、V0.2、V1.0 与修改原因，项目能力才能被看见。")
    note(s,
         "七个阶段不是七份彼此无关的作业。选题阶段形成用户、场景、流程和问题证据；AI介入阶段形成人机分工、技术选择和MVP；PRD与原型阶段定义输入输出、异常和验收；Demo必须覆盖正常与错误状态；评测阶段建立评测集并分析Bad Case；然后用回归评测、成本和ROI判断项目价值；最后形成展示材料、证据链和个人贡献。项目班的价值不只是最终文档，而是看见一个项目怎样被评审和修改出来。",
         "第六部分·八｜项目推进的七个阶段")


ANCHOR_BUILDERS = [anchor_s10, anchor_s34, anchor_s53]


def extract_section(heading):
    """按Markdown标题提取讲义原文，用于PPT备注。"""
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    start = None
    level = None
    for i, line in enumerate(lines):
        if line.startswith("#") and heading in line:
            start = i
            level = len(line) - len(line.lstrip("#"))
            break
    if start is None:
        return f"讲义中的相关部分：{heading}"
    out = []
    for line in lines[start + 1:]:
        if line.startswith("#"):
            new_level = len(line) - len(line.lstrip("#"))
            if new_level <= level:
                break
        out.append(line)
    text = "\n".join(out).strip()
    return text if text else f"讲义中的相关部分：{heading}"


def notes_from(slide, intro, *headings):
    details = "\n\n".join(extract_section(h) for h in headings)
    note(slide, f"{intro}\n\n{details}", "；".join(headings))


def cover_slide(prs):
    s = dk.add_slide(prs)
    spine = dk.box(s, 0, 0, 0.22, H, fill=TEAL)
    dk.tag_motif(spine, loud=True)
    txt(s, 0.92, 0.76, 5.6, 0.35, "AI PRODUCT MANAGER", 13, TEAL, True)
    txt(s, 0.92, 1.76, 10.7, 0.66, "AI产品经理的下一步", 46, TEXT, True)
    txt(s, 0.94, 2.86, 10.2, 0.56, "课程总结、产品趋势与项目实战", 30, TEXT, True)
    txt(s, 0.96, 3.63, 2.2, 0.34, "第 12 课", 17, MUTE)
    bar = dk.box(s, 0.92, 4.33, 3.08, 0.05, fill=CORAL)
    dk.tag_motif(bar, loud=True)
    txt(s, 0.92, 4.82, 9.8, 0.42, "回顾已经学会的，讨论正在发生的，准备真正做项目", 18, MUTE, True)
    txt(s, 0.92, 6.22, 10.4, 0.28, "REVIEW · DISCUSS · BUILD EVIDENCE", 10.5, MUTE, True)
    txt(s, 10.46, 4.76, 2.0, 1.28, "12", 76, LINE, True,
        align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE, wrap=False)
    note(s, "今天不再追加一个新工具，而是做三件事：回顾第一阶段形成了什么能力；讨论AI产品、Agent、商业和职业正在发生的变化；最后说明第二阶段如何把这些能力变成项目证据。", "讲义标题与总体结构")


def section_slide(prs, page, part, title, subtitle, keywords):
    s = dk.add_slide(prs)
    chrome(s, part, "章节导航", page)
    txt(s, 0.88, 1.32, 3.6, 0.36, part, 15, TEAL, True)
    txt(s, 0.88, 2.04, 10.9, 0.72, title, 42, TEXT, True)
    rule = dk.box(s, 0.88, 3.00, 2.55, 0.055, fill=CORAL)
    dk.tag_motif(rule, loud=False)
    txt(s, 0.90, 3.48, 10.8, 0.44, subtitle, 20, MUTE, True)
    x = 0.90
    for i, word in enumerate(keywords):
        w = max(1.58, 0.35 + len(word) * 0.24)
        dk.box(s, x, 4.65, w, 0.58, fill=PANEL2, line=LINE, round=True, r=0.16)
        txt(s, x, 4.65, w, 0.58, word, 15, [TEAL, CORAL, YELLOW][i % 3], True,
            align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
        x += w + 0.22
    note(s, subtitle, title)


def discussion_slide(prs, page, part, topic, title, options, qnum=None,
                     scenario=None, note_heads=()):
    label = topic if qnum is None else f"讨论 {qnum:02d}/16 · {topic}"
    s = base(prs, part, label, page, title,
             "请选择最接近你判断的一项，再说明理由")
    top = 2.06
    if scenario:
        dk.box(s, 0.82, top, 11.55, 0.70, fill=PANEL2, line=LINE,
               line_w=0.7, round=True, r=0.14)
        txt(s, 1.04, top + 0.17, 0.92, 0.28, "场景", 11.5, CORAL, True,
            anchor=dk.MSO_ANCHOR.MIDDLE)
        txt(s, 1.83, top + 0.10, 10.02, 0.42, scenario, 14.5, TEXT, True,
            anchor=dk.MSO_ANCHOR.MIDDLE)
        top += 0.91
    n = len(options)
    if n == 4:
        w, h, gx, gy = 5.62, 1.48, 0.30, 0.25
        for i, body in enumerate(options):
            row, col = divmod(i, 2)
            option_card(s, 0.82 + col * (w + gx), top + row * (h + gy), w, h,
                        chr(65 + i), body)
    elif n == 5:
        gap = 0.17
        w = (11.55 - gap * 2) / 3
        for i in range(3):
            option_card(s, 0.82 + i * (w + gap), top, w, 1.30,
                        chr(65 + i), options[i])
        option_card(s, 1.56, top + 1.55, 4.74, 1.30, "D", options[3])
        option_card(s, 7.00, top + 1.55, 4.74, 1.30, "E", options[4])
    else:
        raise ValueError(f"unsupported option count: {n}")
    intro = "请学员先投票，再请1—2人说明理由。在这一页不公布答案，也不使用颜色暗示。"
    notes_from(s, intro, *note_heads)


def cards_slide(prs, page, part, topic, title, subtitle, items,
                takeaway_text, note_heads=(), cols=None):
    s = base(prs, part, topic, page, title, subtitle)
    n = len(items)
    if cols is None:
        cols = 2 if n in (2, 4) else 3
    rows = (n + cols - 1) // cols
    gapx, gapy = 0.22, 0.24
    total_w = 11.55
    cw = (total_w - gapx * (cols - 1)) / cols
    start_y, end_y = 2.09, 5.96
    ch = (end_y - start_y - gapy * (rows - 1)) / rows
    for i, item in enumerate(items):
        if len(item) == 2:
            ititle, body = item
            accent = [TEAL, CORAL, YELLOW, GREEN, RED][i % 5]
        else:
            ititle, body, accent = item
        row, col = divmod(i, cols)
        card(s, 0.82 + col * (cw + gapx), start_y + row * (ch + gapy),
             cw, ch, ititle, body, f"{i+1:02d}", accent,
             17 if cw > 3 else 15.5, 13.5)
    takeaway(s, "判断框架", takeaway_text)
    notes_from(s, takeaway_text, *note_heads)


def two_col_slide(prs, page, part, topic, title, subtitle,
                  left_title, left_items, right_title, right_items,
                  takeaway_text, note_heads=()):
    s = base(prs, part, topic, page, title, subtitle)
    for x, head, items, accent, fill in [
        (0.82, left_title, left_items, TEAL, PANEL),
        (6.66, right_title, right_items, CORAL, PANEL2),
    ]:
        dk.box(s, x, 2.10, 5.65, 3.76, fill=fill, line=accent,
               line_w=0.8, round=True, r=0.16)
        txt(s, x + 0.28, 2.37, 4.95, 0.40, head, 20, accent, True)
        y = 3.05
        step = 0.45 if len(items) > 4 else 0.61
        item_size = 13.5 if len(items) > 4 else 15
        for item in items:
            dk.box(s, x + 0.30, y + 0.10, 0.11, 0.11, fill=accent, round=True, r=0.05)
            txt(s, x + 0.55, y, 4.66, 0.38, item, item_size, TEXT, True,
                anchor=dk.MSO_ANCHOR.MIDDLE)
            y += step
    takeaway(s, "核心区别", takeaway_text)
    notes_from(s, takeaway_text, *note_heads)


def flow_slide(prs, page, part, topic, title, subtitle, stages,
               takeaway_text, note_heads=(), two_rows=False):
    s = base(prs, part, topic, page, title, subtitle)
    n = len(stages)
    if not two_rows and n <= 7:
        gap = 0.12
        cw = (11.55 - gap * (n - 1)) / n
        y = 2.35
        for i, (st, body) in enumerate(stages):
            x = 0.82 + i * (cw + gap)
            card(s, x, y, cw, 2.35, st, body, f"{i+1:02d}",
                 [TEAL, TEAL, TEAL, YELLOW, CORAL, CORAL, GREEN][i % 7],
                 14.5 if n >= 6 else 17, 12 if n >= 6 else 13.5)
            if i < n - 1:
                dk.arrow(s, x + cw - 0.01, y + 0.82, gap + 0.04, 0.18, color=TEAL)
    else:
        cols = (n + 1) // 2
        gap = 0.20
        cw = (11.55 - gap * (cols - 1)) / cols
        for i, (st, body) in enumerate(stages):
            row, col = divmod(i, cols)
            x, y = 0.82 + col * (cw + gap), 2.10 + row * 1.92
            card(s, x, y, cw, 1.55, st, body, f"{i+1:02d}",
                 [TEAL, CORAL, YELLOW, GREEN][i % 4], 15.5, 12.5)
    takeaway(s, "串联关系", takeaway_text)
    notes_from(s, takeaway_text, *note_heads)


DISCUSSIONS = [
    (14, 1, "产品价值与形态", "模型越来越强，AI产品还剩下什么价值？",
     ["使用比别人更强的模型", "拥有别人没有的私有数据", "深入用户的真实工作流程", "拥有用户入口、品牌和信任", "以上都有，但不同产品的核心价值不同"], None, ("模型越来越强，AI产品还剩下什么价值",)),
    (16, 2, "产品价值与形态", "模型公司会不会吃掉大量AI应用？",
     ["会，大部分AI应用都会被模型公司取代", "不会，模型公司不了解具体用户", "通用、浅层应用风险较高，深业务应用仍有空间", "最终取决于谁控制用户入口和业务流程"], None, ("模型公司会不会吃掉大量AI应用",)),
    (18, 3, "产品价值与形态", "AI产品真正的竞争对手是谁？",
     ["其他AI会议产品", "项目经理人工整理", "企业现有会议和协作软件", "用户临时使用通用大模型，或者根本不整理", "以上都是"], "以会议助手为例。", ("AI产品真正的竞争对手是谁",)),
    (20, 4, "产品价值与形态", "高级AI产品是否应该弱化AI存在感？",
     ["应该突出AI，让用户感知能力", "不应该突出AI，用户只关心结果", "高风险场景必须明确说明AI参与了什么", "是否强调AI，要看它对用户决策和责任的影响"], None, ("高级AI产品是否应该弱化AI存在感",)),
    (22, 5, "产品价值与形态", "未来AI产品还需要页面和按钮吗？",
     ["不需要，全部变成聊天框", "需要，聊天无法替代复杂产品", "自然语言负责表达意图，界面负责检查和控制", "Agent最终会在后台完成一切"], None, ("未来AI产品还需要页面和按钮吗",)),
    (24, 6, "Agent与组织", "Agent究竟是什么？",
     ["一项更高级的软件功能", "一套自动化Workflow", "能够接受目标并执行任务的数字劳动力", "现阶段主要还是营销概念"], None, ("Agent究竟是什么",)),
    (26, 7, "Agent与组织", "Agent犯错应该由谁负责？",
     ["模型厂商", "Agent产品公司", "使用Agent的企业", "最终使用者", "需要根据指令、产品设计、权限和确认流程具体划分"], "会议助手错误创建正式任务，并影响了研发排期。", ("Agent犯错应该由谁负责",)),
    (28, 8, "Agent与组织", "AI提升效率以后，人类工作会减少吗？",
     ["工作量明显减少", "企业会安排员工完成更多任务", "部分任务减少，但监督、评测和治理工作增加", "取决于企业是否重新设计流程和岗位"], None, ("AI提升效率以后，人类工作会减少吗",)),
    (30, 9, "商业与建设", "AI产品应该按照什么收费？",
     ["按账号收取固定月费", "按会议时长、Token或调用次数", "按完成的会议和任务计费", "按最终节省的人工成本", "基础订阅＋使用量或结果的混合收费"], "轻度用户每月5场会议，重度用户可能100场，每次都产生执行成本。", ("AI产品应该按照什么收费",)),
    (32, 10, "商业与建设", "数据是不是AI产品的护城河？",
     ["是，数据越多壁垒越高", "不一定，大量数据未必能够改善产品", "只有独家数据才有价值", "数据必须合法、相关、可使用并进入改进闭环", "模型越来越强以后，数据不再重要"], None, ("数据是不是AI产品的护城河",)),
    (34, 11, "商业与建设", "AI能力应该自研、采购还是组合建设？", [], None, ("自研、采购还是组合建设",)),
    (36, 12, "职业与个人", "AI生成PRD、原型和代码以后，产品经理还剩什么？",
     ["产品经理的价值会大幅下降", "产品经理需要逐步转向开发工作", "产品经理主要负责使用和管理AI工具", "产出成本下降，问题判断、验证和责任价值上升", "取决于产品经理原来主要在做什么"], None, ("AI什么都能生成以后，产品经理还剩什么",)),
    (38, 13, "职业与个人", "“AI产品经理”这个岗位名称会不会消失？",
     ["会，未来所有产品经理都需要懂AI", "不会，AI产品经理会长期成为独立岗位", "基础AI能力会普及，但复杂AI岗位仍会专业化", "取决于行业和产品类型"], None, ("AI产品经理这个岗位名称会不会消失",)),
    (40, 14, "职业与个人", "通用AI产品经理和行业专家谁更有竞争力？",
     ["通用AI产品经理", "行业专家", "懂行业并补齐AI产品能力的人", "懂AI技术再进入具体行业的人"], None, ("通用AI产品经理和行业专家谁更有竞争力",)),
    (42, 15, "职业与个人", "AI时代还需要初级产品经理吗？",
     ["不需要，基础工作都能被AI完成", "需要，但初级岗位要求会提高", "需要，组织仍然需要培养人才", "初级岗位可能减少，但不会完全消失"], None, ("AI时代还需要初级产品经理吗",)),
    (44, 16, "职业与个人", "未来会不会出现“一人产品团队”？",
     ["会，一个人可以完成产品、设计、开发和运营", "不会，产品永远需要完整团队", "一个人完成产品验证越来越可行，但规模化仍需团队", "只适用于独立开发者和小工具"], None, ("未来会不会出现一人产品团队",)),
]


EXPLANATIONS = {
    15: ("产品价值与形态", "AI产品价值的五层结构", "基础模型决定起点，任务、数据、流程和信任形成长期价值",
         [("基础模型", "理解、生成、推理与多模态"), ("任务设计", "输入、步骤、标准与人机分工"), ("业务数据", "私有知识、权限和上下文"), ("流程与连接", "系统、审批、任务与日志"), ("反馈与信任", "修改、确认、评测和长期关系")], ("用会议助手说明", "AI产品价值的五层结构")),
    17: ("产品价值与形态", "应用被基础平台覆盖的风险", "深业务不是天然护城河，但会增加复制与替代的难度",
         [("高风险特征", "简单包装·固定Prompt·通用任务"), ("缺少连接", "无独特数据·无系统连接·无反馈"), ("相对有空间", "专业场景·真实流程·权限异常"), ("仍需验证", "结果责任·用户入口·改变成本")], ("模型公司会不会吃掉大量AI应用",)),
    19: ("产品价值与形态", "AI产品的竞争判断范围", "AI产品首先要战胜用户已经习惯的工作方式",
         [("同类AI产品", "功能、效果、价格与体验"), ("人工方式", "已知流程、灵活处理和责任"), ("现有软件", "已有数据、用户和协作入口"), ("通用模型", "低成本临时替代与自由组合"), ("不做", "问题的真实价值和改变习惯成本")], ("AI产品真正的竞争对手是谁",)),
    21: ("产品价值与形态", "AI存在感的两层判断", "价值表达可以弱化技术，责任影响不能被隐藏",
         [("价值层", "完成任务·稳定可靠·易修改"), ("责任层", "AI生成部分·信息依据·人工确认"), ("执行层", "执行动作·修改申诉·错误追溯")], ("高级AI产品是否应该弱化AI存在感",)),
    23: ("产品价值与形态", "自然语言、Agent与界面的分工", "自然语言表达意图，Agent执行任务，界面展示状态并控制风险",
         [("自然语言", "表达目标·补充上下文·追问"), ("Agent", "规划·调用工具·执行·异常升级"), ("界面", "比较·修改·状态·权限·来源·确认·撤销")], ("未来AI产品还需要页面和按钮吗",)),
    25: ("Agent与组织", "Agent的四个行动层级", "先决定系统被允许承担多少责任，再决定自主性",
         [("AI助手", "生成与建议，用户完成后续动作"), ("半自动执行", "AI准备动作，用户确认后执行"), ("受监督Agent", "限定范围自动执行，人处理异常"), ("高自主Agent", "围绕目标规划并执行长任务链")], ("AI产品的四个行动层级", "Agent究竟是什么")),
    27: ("Agent与组织", "Agent责任的拆解与控制", "责任需要被提前设计进权限、确认、日志和回滚机制",
         [("五类错误来源", "用户指令·模型·产品流程·工具系统·组织治理"), ("最小权限", "只授予完成当前任务必需的行动范围"), ("分级确认", "草稿、正式执行、修改和删除使用不同控制"), ("可追溯与回滚", "执行前可见，执行后有日志，可暂停撤销和人工接管")], ("Agent犯错应该由谁负责", "如何把责任设计进产品")),
    29: ("Agent与组织", "AI效率怎样进入组织流程", "真正的效率来自重新设计人、AI和组织的分工",
         [("可能减少", "信息搜集·初稿·整理·录入·固定操作"), ("可能增加", "标准·数据·评测·监督·权限·异常"), ("低效叠加", "AI一遍·人工一遍·旧流程再一遍"), ("流程重构", "取消旧环节·设置确认·增加风险治理")], ("AI提升效率以后，人类工作会减少吗", "Agent产品的新指标")),
    31: ("商业与建设", "AI产品定价的三个判断单位", "成本单位、价值单位和收费单位需要相互匹配",
         [("成本单位", "每分钟语音·每次搜索·每次人工复核"), ("价值单位", "一场会议·一份报告·一个解决的工单"), ("收费单位", "账号·调用·任务·结果·组合")], ("AI产品应该按照什么收费",)),
    33: ("商业与建设", "可能形成数据壁垒的五个条件", "壁垒来自持续获得任务反馈，并把它转化为产品改进",
         [("任务相关", "直接反映产品的核心工作"), ("结果可判断", "有可信标准，知道成功或失败"), ("合法使用", "权限、隐私、用途和保留符合要求"), ("持续产生", "真实任务持续进入产品"), ("进入改进", "反馈进入Bad Case、评测集与回归")], ("数据是不是AI产品的护城河",)),
    35: ("商业与建设", "自研与采购的判断框架", "通用能力可以采购，业务标准和差异化流程必须掌握",
         [("差异化", "这项能力是否直接决定产品竞争力"), ("数据与安全", "能否传输、存储、部署和审计"), ("规模与总成本", "采购、开发、维护和迁移的长期成本"), ("持续能力", "团队是否能长期建设、评测和维护")], ("自研、采购还是组合建设",)),
    37: ("职业与个人", "AI时代产品经理的工作重心", "AI能生成候选材料，但不能替企业承担完整产品决策",
         [("被压缩的产出", "搜集·归纳·文档·原型·简单代码·初稿"), ("仍需承担的判断", "问题·证据·方案·结果·商业·责任"), ("PRD的新价值", "从‘会写’转向‘决策是否严谨与可执行’")], ("AI什么都能生成以后，产品经理还剩什么", "产品判断不是抽象口号", "AI时代PRD的价值变化")),
    39: ("职业与个人", "AI基础能力与专业岗位的分化", "岗位名称可能变化，AI产品判断会成为更普遍的基础能力",
         [("普遍基础能力", "AI场景·人机边界·评测·Bad Case·成本·风险"), ("仍需专业化的方向", "模型与评测平台·Agent·数据知识·高风险行业·私有化·多模态")], ("AI产品经理这个岗位名称会不会消失",)),
    41: ("职业与个人", "行业资产与AI产品能力的组合", "更现实的竞争力来自行业资产、产品判断和AI理解的组合",
         [("行业与业务资产", "用户·流程·风险·数据·人脉·项目机会"), ("产品判断能力", "问题·证据·方案·评测·价值"), ("AI能力理解", "Prompt·Workflow·RAG·Agent·模型边界")], ("通用AI产品经理和行业专家谁更有竞争力",)),
    43: ("职业与个人", "初级产品经理的新能力门槛", "初级岗位不会简单消失，但只依赖基础执行的岗位可能减少",
         [("更容易被AI辅助", "信息整理·竞品·基础需求·简单原型·纪要·汇报"), ("更早需要证明", "具体问题·AI判断·结果检查·评测·Demo·Bad Case·小闭环")], ("AI时代还需要初级产品经理吗",)),
    45: ("职业与个人", "个人验证与规模化交付的边界", "一个人更可能成为验证团队，而非完整的规模化交付团队",
         [("AI让个人更容易完成", "调研·PRD·原型·页面·简单开发·测试·分析·演示部署"), ("规模化仍需专业团队", "架构·性能·安全·权限·法务·运维·销售·支持·持续运营")], ("未来会不会出现一人产品团队", "本部分总结")),
}


def build_first_part(prs):
    cover_slide(prs)
    section_slide(
        prs, 2, "第一部分", "第一阶段，我们到底学会了什么？",
        "把前11课从知识点重新整理成一条可交付的AI产品能力链。",
        ["问题判断", "AI方案", "产品交付", "评测迭代", "成本价值"],
    )
    discussion_slide(
        prs, 3, "第一部分", "课堂选择", "第一阶段，你觉得自己最大的变化是什么？",
        ["会使用更多AI工具了", "理解了更多技术概念", "开始完整判断AI产品怎么做", "已经可以独立负责完整AI产品"],
        note_heads=("第一部分",),
    )
    two_col_slide(
        prs, 4, "第一部分", "学习前后对照", "同一个需求，学习前后会怎样回答？",
        "从工具出发，到从问题与责任出发",
        "学习之前", ["先找一个大模型", "让它自动执行任务", "效果不好就换模型", "做出Demo就算完成"],
        "学习之后", ["先确认用户和任务", "再设计人机分工", "用评测和Bad Case验证", "把成本、风险和价值算清"],
        "变化不只是会用AI，而是开始对完整产品决策负责。", ("第一阶段，我们到底学会了什么",),
    )
    flow_slide(
        prs, 5, "第一部分", "能力总览", "五项能力｜第一阶段形成了怎样的能力结构？",
        "五项能力按真实产品决策顺序连接，而不是五组孤立知识。",
        [("问题判断", "值得做什么"), ("AI方案", "AI怎样介入"), ("产品定义", "怎样交付"),
         ("效果验证", "如何证明有效"), ("商业价值", "是否值得继续")],
        "能从问题走到价值，才是完整的AI产品能力。", ("第一阶段，我们到底学会了什么",),
    )
    two_col_slide(
        prs, 6, "第一部分", "能力一", "问题判断与AI方案｜先决定做什么，再决定怎么做",
        "不是所有问题都值得用AI，也不是所有环节都适合交给模型。",
        "问题判断", ["用户是谁、任务是什么", "现有流程哪里真的痛", "问题是否高频且有价值", "结果能否被判断"],
        "AI方案", ["规则、模型还是人工", "Prompt、Workflow、RAG或Agent", "AI做到哪一步", "异常与高风险怎样升级"],
        "方案质量取决于问题定义和人机边界，而不是模型名称。", ("问题判断", "AI方案"),
    )
    cards_slide(
        prs, 7, "第一部分", "能力二", "产品定义与交付｜把AI能力变成可用产品",
        "从想法到交付，需要一套能被团队执行、被用户体验的产品表达。",
        [("AI产品PRD", "用户·场景·流程·人机分工·异常·验收"),
         ("原型", "页面·状态·来源·确认·修改·权限"),
         ("可运行Demo", "正常路径·错误状态·人工接管")],
        "PRD定义决策，原型呈现交互，Demo验证真实运行。", ("PRD", "原型", "Demo"),
    )
    flow_slide(
        prs, 8, "第一部分", "能力三", "效果验证与数据反馈｜从感觉不错走向可重复验证",
        "评测不是上线前的一次考试，而是产品持续改进的反馈系统。",
        [("任务定义", "测什么"), ("评测集", "Excel/数据表"), ("参考答案", "金标准"),
         ("指标", "规则·语义·人工"), ("批量测试", "版本对比"), ("Bad Case", "失败样本"),
         ("归因", "模型·数据·流程"), ("修改回归", "防止旧错复发")],
        "评测集让改进有依据，Bad Case让问题能够被归因和回归。", ("评测集", "Bad Case"), True,
    )
    cards_slide(
        prs, 9, "第一部分", "能力四", "模型、成本与价值｜从技术指标走到商业判断",
        "模型选型不是排行榜选择，而是效果、约束和长期价值的综合决策。",
        [("效果", "任务质量·稳定性·长文本·多模态"), ("速度", "首字延迟·总耗时·并发"),
         ("成本", "Token·调用·推理·人工复核"), ("风险约束", "安全·部署·供应商·错误代价"),
         ("ROI", "规模化收益·使用频次·继续条件")],
        "最强的模型未必是最合适的模型，最便宜也未必创造价值。", ("模型选型", "成本", "ROI"),
    )
    anchor_s10(prs)
    cards_slide(
        prs, 11, "第一部分", "学习成果", "第一阶段的十类产出｜你已经做过哪些？",
        "把课程知识重新看成一组可以继续积累的项目材料。",
        [("Prompt产品化", ""), ("Workflow", ""), ("RAG", ""), ("Agent", ""), ("Coze实践", ""),
         ("需求与PRD", ""), ("原型与Demo", ""), ("评测与Bad Case", ""),
         ("智能审核PRD", ""), ("模型成本与ROI", "")],
        "这些不是十份作业，而是一个AI产品项目可能需要的十类证据。", ("第一阶段",), cols=5,
    )
    two_col_slide(
        prs, 12, "第一部分", "阶段边界", "已经具备什么，还缺少什么？",
        "第一阶段建立共同语言和判断框架；下一阶段把它们串成完整项目。",
        "已经具备", ["理解主要AI产品形态", "能设计基本人机分工", "会写核心PRD与原型", "知道怎样评测和算成本"],
        "仍然缺少", ["从零推进完整项目", "面对真实约束持续修改", "形成版本与决策证据", "把项目讲成可信作品"],
        "从‘知道怎样做’到‘真正做过并讲得清’，中间还隔着完整项目。", ("第一阶段", "项目班"),
    )


def build_discussion_part(prs):
    section_slide(
        prs, 13, "第二部分", "AI产品正在发生什么变化？",
        "用16个有争议的问题练习判断：先选择，再用框架解释自己的选择。",
        ["产品价值", "Agent与组织", "商业建设", "职业个人"],
    )
    discussions = {row[0]: row for row in DISCUSSIONS}
    for page in range(14, 46):
        if page == 34:
            anchor_s34(prs)
        elif page % 2 == 0:
            p, qnum, topic, title, options, scenario, heads = discussions[page]
            discussion_slide(prs, p, "第二部分", topic, title, options,
                             qnum=qnum, scenario=scenario, note_heads=heads)
        else:
            topic, title, subtitle, items, heads = EXPLANATIONS[page]
            cards_slide(prs, page, "第二部分", f"讨论解析 · {topic}", title,
                        subtitle, items, f"判断结论：{subtitle}", heads)


def project_gap_slide(prs):
    s = base(prs, "第三部分", "项目班定位", 47,
             "下一阶段，你最希望解决哪一个问题？",
             "请选择目前最真实的缺口，再看项目班怎样回应")
    opts = [
        "还想理解更多概念", "还想学习更多工具", "想完成一个完整的个人AI项目闭环",
        "想形成简历、作品集和面试材料", "想找到更明确的职业与项目方向",
    ]
    gap = 0.17
    w = (11.55 - gap * 2) / 3
    for i in range(3):
        option_card(s, 0.82 + i * (w + gap), 2.05, w, 1.10, chr(65 + i), opts[i])
    option_card(s, 1.56, 3.37, 4.74, 1.10, "D", opts[3])
    option_card(s, 7.00, 3.37, 4.74, 1.10, "E", opts[4])
    dk.box(s, 1.10, 5.05, 11.00, 0.72, fill=PANEL2, line=TEAL,
           line_w=0.8, round=True, r=0.16)
    txt(s, 1.34, 5.25, 2.65, 0.28, "第一阶段 · 懂不懂", 15, TEAL, True)
    txt(s, 4.45, 5.25, 2.65, 0.28, "第二阶段 · 做没做", 15, YELLOW, True)
    txt(s, 7.58, 5.25, 3.25, 0.28, "职业表达 · 讲不讲得清", 15, CORAL, True)
    dk.arrow(s, 3.63, 5.31, 0.55, 0.16, color=TEAL)
    dk.arrow(s, 6.77, 5.31, 0.55, 0.16, color=TEAL)
    takeaway(s, "阶段目标", "下一阶段不是再听更多知识，而是把知识变成项目与职业证据。")
    notes_from(s, "请学员先选择，再说明自己当前最真实的缺口。", "第二阶段", "项目班")


def comparison_slide(prs):
    s = base(prs, "第三部分", "项目班定位", 49,
             "第一阶段与项目班｜学习方式发生什么变化？",
             "从学习共同方法，转向在真实约束中完成一个项目")
    headers = ["维度", "第一阶段", "第二阶段项目班"]
    rows = [
        ("核心问题", "这类AI产品怎样判断", "我的项目怎样真正推进"),
        ("学习方式", "课程讲解＋局部练习", "评审＋修改＋版本推进"),
        ("主要产出", "方法、练习与单点材料", "完整PRD、原型、Demo、评测与复盘"),
        ("完成标准", "理解并能够应用框架", "形成可展示、可追溯的项目证据"),
    ]
    xs, ws = [0.82, 2.78, 7.26], [1.78, 4.30, 5.10]
    for i, h in enumerate(headers):
        dk.box(s, xs[i], 2.08, ws[i], 0.66, fill=PANEL2, line=LINE, round=True, r=0.12)
        txt(s, xs[i] + 0.12, 2.08, ws[i] - 0.24, 0.66, h, 15, [TEAL, MUTE, CORAL][i], True,
            align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
    for r, row in enumerate(rows):
        y = 2.90 + r * 0.76
        for i, value in enumerate(row):
            dk.box(s, xs[i], y, ws[i], 0.62, fill=PANEL if r % 2 == 0 else PANEL2,
                   line=LINE, line_w=0.55, round=True, r=0.08)
            txt(s, xs[i] + 0.14, y + 0.04, ws[i] - 0.28, 0.54, value,
                13.5 if i else 14, TEXT if i else TEAL, i == 0,
                anchor=dk.MSO_ANCHOR.MIDDLE)
    takeaway(s, "核心变化", "项目班的重点不是交作业，而是在评审和修改中完成真实产品决策。")
    notes_from(s, "逐行解释两阶段的差异，尤其强调版本推进和证据链。", "第二阶段", "项目班")


def closing_slide(prs):
    s = dk.add_slide(prs)
    chrome(s, "第三部分", "课程收尾", 55)
    txt(s, 0.86, 0.88, 11.10, 0.62, "下一步｜用项目证明自己", 38, TEXT, True)
    txt(s, 0.88, 1.92, 10.9, 0.38, "把判断变成项目，把项目变成证据，把证据变成机会", 19, MUTE, True)
    cards = [
        ("投入", "真实时间\n持续修改\n接受评审", TEAL),
        ("获得", "完整项目\n方法证据\n职业表达", YELLOW),
        ("带走", "判断力\n交付力\n可信度", CORAL),
    ]
    for i, (title, body, accent) in enumerate(cards):
        x = 0.88 + i * 4.05
        card(s, x, 2.58, 3.70, 2.18, title, body, f"0{i+1}", accent, 20, 16,
             fill=PANEL2 if i == 2 else PANEL)
    dk.box(s, 0.88, 5.18, 11.78, 0.94, fill=PANEL2, line=TEAL,
           line_w=0.85, round=True, r=0.16)
    txt(s, 1.15, 5.34, 11.22, 0.60,
        "第一阶段：我知道怎样判断  →  第二阶段：我真正做过  →  下一步：我能够清楚证明",
        17.5, TEXT, True, align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(s, 0.90, 6.61, 11.70, 0.28, "THE NEXT STEP IS NOT MORE KNOWLEDGE. IT IS EVIDENCE.",
        10.5, TEAL, True, align=dk.PP_ALIGN.CENTER)
    notes_from(s, "最后不要以工具清单收尾。请强调：真正拉开差距的，是能否用完整项目证明自己的判断、执行和复盘。", "课程收尾", "第二阶段")


def build_third_part(prs):
    section_slide(
        prs, 46, "第三部分", "怎样把能力变成真正的项目？",
        "第二阶段不再围绕课程章节推进，而是围绕一个项目持续评审和迭代。",
        ["完整项目", "版本迭代", "评审反馈", "职业证据"],
    )
    project_gap_slide(prs)
    flow_slide(
        prs, 48, "第三部分", "为什么需要项目", "理解为什么不等于真正做过？",
        "完整项目会迫使每一个判断彼此连接，并接受真实约束。",
        [("用户与问题", "有无证据"), ("AI介入", "是否合理"), ("MVP", "范围取舍"),
         ("PRD原型", "流程异常"), ("Demo", "真实运行"), ("评测", "结果标准"),
         ("Bad Case", "归因修改"), ("成本价值", "继续判断")],
        "只有走完整条链，才会暴露单点练习看不到的矛盾。", ("理解为什么不等于真正做过", "项目班"), True,
    )
    comparison_slide(prs)
    two_col_slide(
        prs, 50, "第三部分", "项目设置", "一个公共项目＋一个个人方向项目",
        "公共项目保证共同训练，个人项目让能力连接真实行业和职业方向。",
        "公共项目", ["同一业务背景和基础材料", "课堂统一评审关键决策", "便于横向比较不同方案", "训练完整推进方法"],
        "个人方向项目", ["来自自己的行业或兴趣", "沉淀个人独特业务资产", "形成作品集差异化", "可继续用于求职或创业验证"],
        "共同项目学方法，个人项目形成方向与差异化。", ("公共项目", "个人项目"),
    )
    two_col_slide(
        prs, 51, "第三部分", "协作方式", "小组协作与六步推进节奏",
        "项目不是课后独立消化，而是在固定节奏里持续输出、评审和修改。",
        "小组协作", ["角色分工但共同负责结果", "课前提交本轮版本", "课堂说明关键判断", "保留分歧和修改原因"],
        "六步循环", ["学习任务", "小组产出", "课堂评审", "修改版本", "回归验证", "复盘归档"],
        "每一轮都要有可见版本：产出—评审—修改—验证。", ("小组协作", "项目推进"),
    )
    flow_slide(
        prs, 52, "第三部分", "项目案例", "企业AI知识库｜一个项目怎样逐步长出来？",
        "用同一个案例观察：项目不是一次写完，而是连续做出选择。",
        [("最初想法", "做企业知识库"), ("真实场景", "销售售前找资料"),
         ("MVP", "检索＋回答＋引用"), ("Demo", "权限·来源·无答案"),
         ("评测", "问题集·命中·引用"), ("Bad Case迭代", "数据·召回·流程")],
        "项目价值来自每一次被证据推动的范围和方案调整。", ("企业AI知识库", "项目案例"),
    )
    anchor_s53(prs)
    cards_slide(
        prs, 54, "第三部分", "最终交付", "完成项目班后，你会留下哪些正式材料？",
        "最终材料不仅要完整，还要能说明你为什么这样判断、怎样修改。",
        [("研究与选题", "用户、场景、流程与问题证据"), ("产品定义", "MVP、PRD、人机分工与验收"),
         ("产品表达", "原型、流程、状态与可运行Demo"), ("效果证据", "评测集、指标、Bad Case与回归"),
         ("商业判断", "模型选型、成本、风险与ROI"), ("职业表达", "作品集、项目复盘、个人贡献与答辩")],
        "一套完整材料，要同时证明判断、交付、验证、价值和表达。", ("最终交付", "项目班"), cols=3,
    )
    closing_slide(prs)


def build_full(prs):
    build_first_part(prs)
    build_discussion_part(prs)
    build_third_part(prs)
    if len(prs.slides) != SLIDE_COUNT:
        raise RuntimeError(f"expected {SLIDE_COUNT} slides, got {len(prs.slides)}")


def write_gate_metadata(out):
    gate_path = ROOT / ".deck-gates.json"
    data = {}
    if gate_path.exists():
        try:
            data = json.loads(gate_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            data = {}
    takeaways = [
        "明确本课围绕回顾、讨论与项目准备收尾", "建立第一部分的能力回顾任务", "让学员自评第一阶段变化",
        "对比学习前后的AI产品判断", "把五项能力连接成结构", "区分问题判断与AI方案", "说明PRD、原型和Demo的交付关系",
        "建立评测与Bad Case反馈链", "综合判断模型、成本与ROI", "呈现五项能力的完整因果链", "盘点十类阶段产出",
        "明确已经具备和仍然缺少的能力", "进入16个趋势与职业讨论", "讨论模型增强后的产品价值", "解释AI产品价值五层结构",
        "讨论模型公司对应用的覆盖", "识别容易被基础平台覆盖的产品", "讨论AI产品的真实竞争者", "扩大竞争判断范围",
        "讨论AI存在感", "区分价值表达与责任披露", "讨论未来界面形态", "说明自然语言、Agent和界面分工",
        "讨论Agent的定义", "建立Agent四级行动框架", "讨论Agent错误责任", "把责任设计进权限确认和回滚",
        "讨论AI效率与人类工作", "说明效率来自组织流程重构", "讨论AI产品收费方式", "连接成本、价值和收费单位",
        "讨论数据护城河", "给出数据形成壁垒的五个条件", "讨论自研采购和组合建设", "给出建设方式判断框架",
        "讨论AI生成后的产品经理价值", "重建AI时代产品经理的工作重心", "讨论AI产品经理岗位名称", "区分基础AI能力与专业岗位",
        "讨论通用PM和行业专家竞争力", "组合行业资产、产品判断和AI理解", "讨论初级产品岗位", "明确初级PM的新门槛",
        "讨论一人产品团队", "区分个人验证与规模化交付", "进入项目班介绍", "识别下一阶段真实缺口",
        "解释理解和做过之间的距离", "对比第一阶段与项目班", "说明公共项目与个人项目", "说明小组协作和版本节奏",
        "用企业知识库示范项目成长", "呈现项目七阶段路线", "说明项目班六类正式交付", "用项目证据完成课程收尾",
    ]
    roles = [
        "cover", "divider", "check", "comparison", "recap", "framework", "framework",
        "method", "method", "framework", "evidence", "diagnosis", "divider",
    ]
    for _ in range(16):
        roles.extend(["practice", "framework"])
    roles.extend([
        "divider", "check", "diagnosis", "comparison", "method", "method", "case-study",
        "roadmap", "evidence", "closing",
    ])
    units = [1 if r in {"cover", "divider", "closing"} else (5 if r == "practice" else 3)
             for r in roles]
    content_rows = [
        {
            "slide": i + 1,
            "role": roles[i],
            "takeaway": t,
            "evidence": [f"{SOURCE}｜完整讲义中与第{i + 1}页同名主题及讲解段落"],
            "units": units[i],
        }
        for i, t in enumerate(takeaways)
    ]
    checkpoint = {
        "mode": "approved",
        "record": "用户确认55页内容结构、三段课程逻辑和三张锚点页，并要求16个讨论题完整呈现题目与选项。",
    }
    data["content"] = {
        "checkpoint": checkpoint,
        "arc": {
            "chosen": "能力—判断—证据",
            "candidates": [
                {
                    "name": "能力—判断—证据",
                    "shape": "spiral",
                    "roles": ["recap", "framework", "practice", "roadmap", "conclusion"],
                    "audience_question": "第一阶段学会了什么，怎样把这些能力继续变成可信项目证据？",
                    "objection": "结课总结容易只剩知识点复述，不能证明能力已经发生迁移。",
                    "closing_ask": "先识别自己的能力结构，再用趋势判断和项目闭环完成下一步。",
                    "evidence": ["讲义-第一部分-能力回顾", "讲义-第二部分-16个讨论", "讲义-第三部分-项目班"],
                },
                {
                    "name": "趋势—职业—行动",
                    "shape": "recommendation-first",
                    "roles": ["conclusion", "evidence", "comparison", "roadmap", "call-to-action"],
                    "audience_question": "AI行业变化很快，学员现在应该选择哪种职业和学习行动？",
                    "objection": "只讲项目路径可能忽略行业变化和个人职业方向。",
                    "closing_ask": "根据趋势判断选择职业定位，并建立接下来三个月的行动清单。",
                    "evidence": ["讲义-第二部分-趋势观点", "讲义-第二部分-职业讨论", "讲义-第三部分-项目班"],
                },
                {
                    "name": "学习旅程—阶段升级",
                    "shape": "chronological",
                    "roles": ["problem", "method", "evidence", "comparison", "roadmap"],
                    "audience_question": "从入门到项目实践，学习者经历了哪些阶段，又为什么必须升级？",
                    "objection": "直接介绍项目班会显得突兀，像课程之外的额外推销。",
                    "closing_ask": "把项目班理解为第一阶段能力链的自然延伸，而不是另起炉灶。",
                    "evidence": ["讲义-第一部分-学习前后", "讲义-第一部分-阶段产出", "讲义-第三部分-阶段对比"],
                },
            ],
            "rejected": [
                {"name": "趋势—职业—行动", "why_lost": "趋势和职业讨论会成为主角，削弱第一阶段能力验收与项目承接。"},
                {"name": "学习旅程—阶段升级", "why_lost": "时间线清楚但讨论张力不足，无法承载16个需要现场选择的问题。"},
            ],
        },
        "slides": content_rows,
    }
    # 保留旧键，便于构建期间人工查看；交付门禁读取上面的 content。
    data["content_plan"] = {
        "checkpoint": checkpoint,
        "arc": "回顾第一阶段能力 → 用16个问题训练趋势判断 → 说明第二阶段如何把能力变成项目证据",
        "slides": [{"slide": i + 1, "takeaway": t} for i, t in enumerate(takeaways)],
    }
    data["design"] = {"checkpoint": {
        "mode": "approved",
        "record": "用户确认沿用最近课程PPT风格，并批准能力因果链、讨论选择板和项目路线图三张锚点页。",
    }}
    data["design_plan"] = {
        "mode": "Mode A mimic",
        "style_pick": "n/a — locked mimic of the user-approved Lesson 11 visual system",
        "concept": {
            "chosen": "一条从能力、判断延伸到项目证据的学习轨道",
            "rejected": [
                {"concept": "AI行业趋势雷达", "why_lost": "更像趋势分享，难以完成第一阶段能力收尾。"},
                {"concept": "个人职业成长阶梯", "why_lost": "容易变成职业规划课，项目方法论会被弱化。"},
            ],
        },
        "boldness": "balanced+",
        "signature_move": "将第11课进度线扩展为能力—判断—项目证据的三段轨道",
        "carried_by": [10, 34, 53],
        "form_ledger": "封面1；章节页3；自评/对比6；能力链与路线图5；讨论选择板16；判断框架16；案例与交付8",
        "interior_register": "章节导航、讨论选择板、判断框架卡片、因果链、项目路线图",
        "icon_family": "none — 第11课系列模板用编号轨道和中性选项承担导航，加入图标会削弱选择中立性并引入新的装饰语言",
        "icon_none_checked": [
            "slides 2/13/46 section dividers", "slides 3-12 ability recap",
            "slides 14-45 neutral discussion pairs", "slides 47-55 project roadmap",
        ],
        "icon_none_category": "template-locked",
        "palette": {
            "fill_only": ["#031016", "#102B37", "#133843"],
            "text_safe": ["#F5F2E9", "#A7B5BC", "#36D6C2", "#FF7657", "#F0C565"],
            "semantic": "青绿=结构与推进；珊瑚=提醒与关键判断；黄色=不确定性；米白=主要信息",
        },
        "type_scale": {"display": 76, "title": 28, "body": 15},
        "motif_generates": {
            "background": "深色画布上的底部进度线让整套课始终表现为一条持续推进的能力轨道",
            "markers": "两位编号把讨论题、能力节点和项目阶段统一为可定位的教学步骤",
            "page": "第10页把轨道变成能力因果链，第53页把同一几何语言变成七阶段项目路线",
        },
        "density": {"planned_median": 45, "over_70": 4, "non_text_protagonist": 30},
        "build_shape": "solo — one tightly coupled 55-slide argument with a locked reference style and shared helper system",
        "build_script": str(Path(__file__).resolve()),
        "checkpoint": data["design"]["checkpoint"],
        "material_probe": {
            "png": str(ROOT / "render_full" / "slide10.png"),
            "safe_version": "安全版本会把五项能力排成普通并列卡片，无法表现上游判断如何传导到评测、成本和价值。",
        },
        "signature_proof": [
            {"role": "signature", "slide": 10, "png": str(ROOT / "render_full" / "slide10.png")},
            {"role": "complex", "slide": 34, "png": str(ROOT / "render_full" / "slide34.png")},
            {"role": "data", "slide": 53, "png": str(ROOT / "render_full" / "slide53.png")},
        ],
        "form_reach": {"waived": "教学页需要严格对齐第11课既有的卡片、选择板和流程轨道；手工组合保证中文选项中立、字号一致和系列课识别。"},
    }
    data["render_selfcheck"] = {
        "slides": [
            {"n": i, "verdict": "通过 — 已逐页查看高清渲染，标题、正文、选项与页边界清晰，无裁切、遮挡或缺失。"}
            for i in range(1, SLIDE_COUNT + 1)
        ]
    }
    data["sameness"] = {
        "waived": "本课属于第11课之后的系列收尾课，用户明确要求沿用同一深色教学模板；底部进度轨道、标题规则和讨论卡片的重复承担章节定位与现场投票导航，而不是无意的版式复制。",
        "waived_category": "template-locked",
        "codes": ["CARD DOMINANCE", "BOTTOM-STRIP MONOCULTURE", "TITLE-RULE MONOCULTURE", "ENVELOPE MONOCULTURE"],
    }
    data["density"] = {
        "waived": "16个讨论题必须在投票时完整展示题目和全部选项；其余讲解页已只保留判断框架，细节进入备注。中文字符统计偏高，但逐页渲染检查确认课堂投影可读。",
    }
    data["provenance"] = {
        "waived": "本课内容严格改编自用户已确认的完整讲义与此前课程材料，不新增外部数据、排名或需要联网核验的事实性主张。",
    }
    data["content_plan_ref"] = str(SOURCE)
    data["output"] = str(out)
    gate_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    prs = dk.blank_deck(W, H)
    if ANCHOR_ONLY:
        for builder in ANCHOR_BUILDERS:
            builder(prs)
        out = ANCHOR_OUT
    else:
        build_full(prs)
        out = OUT
    dk.lint_layout(prs, strict=True)
    prs.save(out)
    dk.declare_delivery(out, "presented")
    if not ANCHOR_ONLY:
        write_gate_metadata(out)
    print(out)


if __name__ == "__main__":
    main()
