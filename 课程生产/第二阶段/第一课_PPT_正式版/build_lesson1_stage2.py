#!/usr/bin/env python3
"""第二阶段第一课完整PPT构建脚本。

内容严格来自老师版完整讲义 v1.1，视觉沿用第11、12课的深色青绿教学体系。
"""
from pathlib import Path
import json
import sys

SKILL = Path("/Users/keivn/.codex/skills/slide-maker")
sys.path.insert(0, str(SKILL / "scripts"))
import deckkit as dk  # noqa: E402


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "第一课_AI项目立项、选题与完成标准_完整PPT_v1.0.pptx"
SOURCE = Path("/Users/keivn/Project/AI-Course/课程生产/第二阶段/02_第一课_AI项目立项选题与完成标准_老师版完整讲义_v1.0.md")
GATES = ROOT / ".deck-gates.json"

W, H = 13.333, 7.5
SLIDE_COUNT = 60

# Mode A mimic：沿用最近两课已确认的深色青绿教学系统。
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
    # Latin 和东亚文字都指向同一个 CJK 字体，避免 LibreOffice
    # 渲染时忽略 ea 字体并把中文渲染成方块。
    mono="Menlo", font="Hiragino Sans GB", display="Hiragino Sans GB",
    eadisplay="Hiragino Sans GB", eafont="Hiragino Sans GB",
    accents=[TEAL, CORAL, YELLOW, GREEN],
)
dk.set_ground(BG)
dk.set_geometry(radius=1.0, rule_w=0.85)


PARTS = {
    1: ("课程导入", "项目实战班第一课"),
    2: ("第一部分", "第二阶段需要完成什么"),
    9: ("第二部分", "什么样的材料才算完整AI项目"),
    26: ("第三部分", "课程项目怎样写进简历和作品集"),
    33: ("第四部分", "三个公共项目分别训练什么"),
    45: ("第五部分", "公共项目怎样迁移成个人项目"),
    48: ("第六部分", "怎样判断一个选题是否值得做"),
    52: ("第七部分", "现场诊断并确定第一个个人项目"),
    56: ("第八部分", "第一个项目接下来怎样推进"),
}


def part_for(page):
    current = PARTS[1]
    for n in sorted(PARTS):
        if page >= n:
            current = PARTS[n]
    return current


def txt(slide, x, y, w, h, value, size, color=TEXT, bold=False,
        align=dk.PP_ALIGN.LEFT, anchor=dk.MSO_ANCHOR.TOP, wrap=True,
        font="Hiragino Sans GB"):
    return dk.text(
        slide, x, y, w, h,
        [[(str(value), size, color, bold, False, font, "Hiragino Sans GB")]],
        align=align, anchor=anchor, space_after=0, wrap=wrap,
    )


def chrome(slide, page):
    part, topic = part_for(page)
    # 加大的章节标题，解决此前“第一节、第二节不够明显”的问题。
    txt(slide, 0.76, 0.20, 10.6, 0.36, f"{part}｜{topic}", 14.5, TEAL, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(slide, 12.03, 0.22, 0.55, 0.30, f"{page:02d}", 10.5, MUTE, True,
        align=dk.PP_ALIGN.RIGHT, anchor=dk.MSO_ANCHOR.MIDDLE)
    dk.box(slide, 0.76, 7.17, 11.82, 0.014, fill=LINE)
    progress = dk.box(slide, 0.76, 7.11, 11.82 * page / SLIDE_COUNT, 0.04, fill=TEAL)
    dk.tag_motif(progress, loud=False)


def heading(slide, title, subtitle=None, accent=TEAL):
    txt(slide, 0.78, 0.72, 11.72, 0.62, title, 29, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    rule = dk.box(slide, 0.78, 1.38, 1.46, 0.045, fill=accent)
    dk.tag_motif(rule, loud=False)
    if subtitle:
        txt(slide, 0.78, 1.50, 11.66, 0.40, subtitle, 14.5, MUTE)


def base(prs, page, title, subtitle=None, accent=TEAL):
    s = dk.add_slide(prs)
    chrome(s, page)
    heading(s, title, subtitle, accent)
    return s


def card(slide, x, y, w, h, title, body="", n=None, accent=TEAL,
         title_size=17, body_size=13.5, fill=PANEL, border=LINE):
    dk.box(slide, x, y, w, h, fill=fill, line=border, line_w=0.8,
           round=True, r=0.16)
    top = y + 0.22
    if n is not None:
        txt(slide, x + 0.22, top, 0.58, 0.25, n, 10.5, accent, True)
        top += 0.32
    txt(slide, x + 0.22, top, w - 0.44, 0.42, title, title_size, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    if body:
        txt(slide, x + 0.22, top + 0.46, w - 0.44,
            h - (top - y) - 0.58, body, body_size, MUTE)


def takeaway(slide, label, body, y=6.28, accent=CORAL):
    dk.box(slide, 0.82, y, 11.52, 0.60, fill=PANEL2, line=LINE,
           line_w=0.65, round=True, r=0.14)
    txt(slide, 1.04, y + 0.14, 1.24, 0.28, label, 11.5, accent, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(slide, 2.25, y + 0.11, 9.72, 0.34, body, 15, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)


def bullet_list(slide, items, x=0.90, y=2.02, w=7.10, h=3.95,
                color=TEAL, size=17, gap=0.64):
    for i, item in enumerate(items):
        yy = y + i * gap
        dk.box(slide, x, yy + 0.17, 0.10, 0.10, fill=color, round=True, r=0.05)
        txt(slide, x + 0.26, yy, w - 0.26, 0.48, item, size, TEXT, False,
            anchor=dk.MSO_ANCHOR.MIDDLE)


def extract_section(heading):
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
    return "\n".join(out).strip()


def notes_from(slide, intro, *headings):
    details = "\n\n".join(extract_section(h) for h in headings)
    dk.speaker_notes(slide, f"{intro}\n\n{details}\n\n[讲义来源]\n{SOURCE}")


def cover_slide(prs):
    s = dk.add_slide(prs)
    dk.slide_background(s, BG)
    dk.box(s, 0.28, 0, 0.16, H, fill=TEAL)
    txt(s, 0.92, 0.66, 4.7, 0.30, "AI PRODUCT MANAGER · PROJECT STAGE", 12, TEAL, True)
    txt(s, 0.92, 1.42, 10.6, 1.46, "AI项目立项、选题\n与完成标准", 42, TEXT, True)
    txt(s, 0.94, 3.22, 8.9, 0.44, "第二阶段第一课｜从一个想法走到可推进的项目", 20, TEXT, True)
    dk.box(s, 0.94, 3.78, 2.08, 0.05, fill=CORAL)
    txt(s, 0.94, 4.08, 8.9, 0.60,
        "看懂完整项目  ·  选定个人方向  ·  写出第一版边界", 16, MUTE)
    txt(s, 10.43, 4.72, 1.35, 1.00, "01", 62, TEAL, True,
        align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
    txt(s, 0.94, 6.72, 4.8, 0.25, "第二阶段 · 项目实战班", 11.5, MUTE, True)
    notes_from(s, "开场先告诉学员，这节课不急着写PRD，也不急着选模型。今天只做一件事：判断一个想法是否已经具备进入项目的条件。", "第一课｜AI项目立项、选题与完成标准")


def slide_statement(prs, page, title, statement, support, note_headings, accent=TEAL):
    s = base(prs, page, title)
    txt(s, 0.94, 2.12, 10.95, 1.34, statement, 30, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    dk.box(s, 0.95, 3.72, 0.08, 1.20, fill=accent)
    txt(s, 1.25, 3.72, 10.58, 1.20, support, 17, MUTE,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    notes_from(s, statement, *note_headings)


def slide_cards(prs, page, title, subtitle, items, note_headings,
                takeaway_text=None, cols=3):
    s = base(prs, page, title, subtitle)
    rows_n = (len(items) + cols - 1) // cols
    gap_x, gap_y = 0.18, 0.20
    x0, y0 = 0.82, 2.06
    total_w = 11.52
    bottom = 6.10 if takeaway_text else 6.72
    total_h = bottom - y0
    cw = (total_w - gap_x * (cols - 1)) / cols
    ch = (total_h - gap_y * (rows_n - 1)) / rows_n
    for i, item in enumerate(items):
        r, c = divmod(i, cols)
        x = x0 + c * (cw + gap_x)
        y = y0 + r * (ch + gap_y)
        title_i, body_i = item[0], item[1]
        accent = item[2] if len(item) > 2 else [TEAL, CORAL, YELLOW, GREEN][i % 4]
        card(s, x, y, cw, ch, title_i, body_i, f"{i+1:02d}", accent,
             title_size=16.5, body_size=13.2)
    if takeaway_text:
        takeaway(s, "本页重点", takeaway_text)
    notes_from(s, title, *note_headings)


def slide_compare(prs, page, title, subtitle, left_title, left_items,
                  right_title, right_items, note_headings, right_accent=CORAL,
                  takeaway_text=None):
    s = base(prs, page, title, subtitle)
    y, h = 2.06, 3.96 if takeaway_text else 4.58
    for x, heading_text, items, accent in [
        (0.82, left_title, left_items, TEAL),
        (6.70, right_title, right_items, right_accent),
    ]:
        dk.box(s, x, y, 5.64, h, fill=PANEL, line=accent, line_w=0.9,
               round=True, r=0.16)
        txt(s, x + 0.28, y + 0.22, 5.08, 0.40, heading_text, 18, accent, True)
        bullet_list(s, items, x=x + 0.30, y=y + 0.82, w=5.02, h=h - 1.02,
                    color=accent, size=14.4, gap=(h - 1.05) / max(len(items), 1))
    if takeaway_text:
        takeaway(s, "判断标准", takeaway_text)
    notes_from(s, title, *note_headings)


def slide_flow(prs, page, title, subtitle, stages, note_headings,
               takeaway_text=None, cols=None):
    s = base(prs, page, title, subtitle)
    if cols is None:
        cols = min(len(stages), 6)
    rows_n = (len(stages) + cols - 1) // cols
    x0, y0, total_w = 0.82, 2.12, 11.52
    gap_x, gap_y = 0.16, 0.42
    bottom = 5.98 if takeaway_text else 6.55
    ch = (bottom - y0 - gap_y * (rows_n - 1)) / rows_n
    cw = (total_w - gap_x * (cols - 1)) / cols
    for i, (t, b) in enumerate(stages):
        r, c = divmod(i, cols)
        x = x0 + c * (cw + gap_x)
        y = y0 + r * (ch + gap_y)
        card(s, x, y, cw, ch, t, b, f"{i+1:02d}",
             [TEAL, TEAL, YELLOW, CORAL, GREEN][min(i, 4)],
             title_size=14.7 if cols >= 5 else 16.2,
             body_size=11.6 if cols >= 5 else 13.0)
        if c < cols - 1 and i + 1 < len(stages):
            dk.arrow(s, x + cw - 0.01, y + ch / 2 - 0.08,
                     gap_x + 0.04, 0.16, color=TEAL)
    if takeaway_text:
        takeaway(s, "处理结果", takeaway_text)
    notes_from(s, title, *note_headings)


def slide_table(prs, page, title, subtitle, rows, col_w, note_headings,
                size=13.0, row_h=0.48, highlight=None, takeaway_text=None):
    s = base(prs, page, title, subtitle)
    dk.table(s, 0.82, 2.05, 11.52, rows, col_w=col_w, header=True,
             highlight=highlight, size=size, row_h=row_h,
             head_c=TEAL, body_c=MUTE, rule_c=LINE,
             hi_fill=PANEL2, hi_c=CORAL, font="Hiragino Sans GB")
    if takeaway_text:
        takeaway(s, "本页重点", takeaway_text)
    notes_from(s, title, *note_headings)


def slide_quote_analysis(prs, page, title, quote, analyses, note_headings,
                         quote_label="原始表达", takeaway_text=None):
    s = base(prs, page, title)
    dk.box(s, 0.84, 1.88, 11.48, 1.25, fill=PANEL2, line=TEAL,
           line_w=0.8, round=True, r=0.16)
    txt(s, 1.08, 2.06, 1.30, 0.28, quote_label, 11.5, TEAL, True)
    txt(s, 2.23, 1.98, 9.65, 0.72, f"“{quote}”", 19, TEXT, True,
        anchor=dk.MSO_ANCHOR.MIDDLE)
    n = len(analyses)
    gap = 0.18
    cw = (11.48 - gap * (n - 1)) / n
    for i, item in enumerate(analyses):
        x = 0.84 + i * (cw + gap)
        card(s, x, 3.48, cw, 2.27 if takeaway_text else 2.90,
             item[0], item[1], f"{i+1:02d}",
             [CORAL, YELLOW, TEAL, GREEN][i % 4], 15.5, 12.5)
    if takeaway_text:
        takeaway(s, "修改方向", takeaway_text)
    notes_from(s, title, *note_headings)


def build_deck():
    # 本系列课程使用 13.333×7.5 英寸宽屏画布；DeckKit 默认为
    # 10×5.625，宽高比相同但坐标尺度不同，因此需要显式声明尺寸。
    prs = dk.blank_deck(W, H)
    cover_slide(prs)

    # 第一部分
    slide_statement(prs, 2, "第二阶段的学习任务",
                    "把前面学过的内容，连续地用在一个项目里",
                    "第一阶段已经讲过模型、提示词、工作流、知识检索、智能体、产品文档、原型、评测、成本和投入产出比。现在要练习的是：面对一个不完整的问题，怎样一步步把它做成产品项目。",
                    ["为什么第二阶段不再继续讲新概念"])
    slide_flow(prs, 3, "从知识点到完整项目", "知识会在项目需要时重新调用", [
        ("业务问题", "谁在什么场景遇到什么困难"),
        ("产品方案", "AI进入哪一步，为什么"),
        ("PRD与原型", "流程、页面与边界"),
        ("测试与修改", "样本、错误和版本变化"),
        ("价值判断", "风险、成本和是否继续"),
    ], ["为什么第二阶段不再继续讲新概念"],
       "第二阶段把分散的知识放回同一条项目推进过程。", cols=5)
    slide_cards(prs, 4, "三个个人项目", "每个项目对应一种常见的企业AI产品问题", [
        ("对话与业务办理", "理解自然语言，把咨询转成可控的服务流程", TEAL),
        ("材料审核与风险核验", "提取业务材料，结合规则发现差异并交给人工确认", CORAL),
        ("专业研究与方案生成", "理解复杂要求，调用企业知识形成专业文档", YELLOW),
    ], ["第二阶段的三个个人项目"],
       "项目按业务问题组织；知识检索、工作流和智能体会按需组合。", cols=3)
    slide_compare(prs, 5, "三个项目的完成深度", "三个都要完整，其中一个继续做深",
                  "三个项目都要具备", ["业务背景和目标用户", "现状流程与主要问题", "输入、输出和人机分工", "MVP 范围、流程、页面、测试和修改记录"],
                  "主项目继续补齐", ["可运行验证", "更系统的测试", "典型错误迭代", "风险、成本和作品集"],
                  ["为什么三个项目不要求做成相同深度"], GREEN,
                  "最后形成三个完整方案，其中一个有更扎实的验证过程。")
    slide_cards(prs, 6, "两个选修方向", "它们扩展视野，不增加统一作业数量", [
        ("AI图像判断与结构化报告", "适合美业、质检、巡检、设计等多模态场景", YELLOW),
        ("客户研究与销售方案助手", "适合销售、咨询、保险、教育和客户成功场景", CORAL),
        ("三种使用方式", "了解项目、完成简化练习，或经确认后替换一个个人项目", TEAL),
    ], ["两个选修方向为什么不增加必修项目数量"],
       "五个方向可以学习，仍然只要求完成三个个人项目。", cols=3)
    slide_compare(prs, 7, "公共项目和个人项目", "课堂案例建立标准，个人项目重新建立业务内容",
                  "公共项目", ["老师完整拆解", "统一讲清问题、流程和边界", "展示页面、测试与修改方式"],
                  "个人项目", ["结合自己的行业或经历", "重新确定用户、材料和风险", "保留自己的判断和版本变化"],
                  ["公共项目和个人项目有什么区别"], CORAL,
                  "可以复用项目方法，不能只换项目名称。")
    slide_cards(prs, 8, "第二阶段的学习方式", "每次课都围绕同一个项目推进", [
        ("课堂", "老师拆解公共项目，解释为什么这样设计", TEAL),
        ("课后", "把方法迁移到个人项目", YELLOW),
        ("点评", "全班先判断，老师再讲问题和修改方向", CORAL),
        ("版本", "保留 V0.1 版本、点评意见和修改后的版本", GREEN),
    ], ["公共项目和个人项目有什么区别"],
       "学习结果看项目怎样被判断和修改，不看听过多少概念。", cols=4)

    # 第二部分
    slide_quote_analysis(prs, 9, "从一句项目名称开始", "我想做一个AI智能客服", [
        ("使用者", "谁会真正打开这个产品？"),
        ("使用时刻", "用户在什么情况下需要它？"),
        ("当前做法", "没有它时，任务怎样完成？"),
        ("项目边界", "第一版解决哪一个具体问题？"),
    ], ["项目不是一个名字"], takeaway_text="先把用户、场景和主要问题说清楚。")
    slide_quote_analysis(prs, 10, "功能清单还缺什么", "支持知识库问答、订单查询、客户投诉和商品推荐", [
        ("优先级", "哪个问题最频繁、最严重？"),
        ("材料", "知识、订单和商品数据从哪里来？"),
        ("责任", "哪些情况自动处理，哪些转人工？"),
        ("范围", "为什么第一版要同时做四个场景？"),
    ], ["项目也不是一组功能"], takeaway_text="功能多不代表项目已经成立。")
    slide_compare(prs, 11, "PRD和Demo各自能说明什么", "两类材料都重要，但都需要放回完整项目中",
                  "只有 PRD", ["能说明产品设计", "不能自动证明用户真实存在", "不能证明数据和流程可用", "不能证明 AI 效果"],
                  "只有 Demo", ["能说明一条演示路径跑通", "不能证明用户需要", "不能覆盖异常和风险", "不能证明成本可接受"],
                  ["PRD和Demo为什么都不能单独证明项目完整"], CORAL,
                  "完整度看问题、方案、页面和测试是否前后对应。")
    slide_flow(prs, 12, "一个完整项目需要回答的问题", "每一步都要和前后内容对得上", [
        ("用户与场景", "谁在何时工作"),
        ("当前流程", "任务现在怎样完成"),
        ("主要问题", "哪一步最需要改变"),
        ("材料与数据", "实际输入从哪里来"),
        ("能力分工", "模型、规则、系统和人工"),
        ("第一版", "做什么，暂时不做什么"),
        ("测试", "怎样判断结果能否使用"),
        ("修改", "发现问题后改了什么"),
    ], ["完整项目是一条连续的证据链"],
       "项目完整度来自前后对应，不来自材料页数。", cols=4)
    slide_cards(prs, 13, "项目例会待办整理助手的业务背景", "只处理会后整理这一小步", [
        ("整理耗时", "会议结束后还要花时间整理决定、待办和风险", TEAL),
        ("责任不清", "讨论了很多，却没有明确负责人和截止时间", CORAL),
        ("信息容易漏", "相对日期、重复任务和未解决风险常被遗漏", YELLOW),
        ("理解不一致", "纪要发出后，参会人对任务仍有不同理解", GREEN),
    ], ["业务背景"],
       "项目不解决所有会议管理问题，只聚焦会后整理。", cols=4)
    slide_statement(prs, 14, "项目定义",
                    "项目经理粘贴会议转写，系统整理决定、待办、负责人、时间和风险",
                    "负责人或时间不明确的内容会被标记。项目经理修改确认后，再导出结构化纪要。",
                    ["一句话项目定义"])
    slide_quote_analysis(prs, 15, "一段会议转写里有什么",
                         "登录接口这周可能来不及。小王先和后端确认一下，如果周四还没有结果，我们就把联调改到下周一。", [
        ("待办", "小王与后端确认接口进度"),
        ("时间", "周四是确认节点"),
        ("条件决定", "如果没有结果，联调改期"),
        ("风险", "接口本周可能无法完成"),
    ], ["为什么需要大模型"], quote_label="会议转写")
    slide_cards(prs, 16, "模型、规则、系统和人工的分工", "每一类能力只负责适合自己的部分", [
        ("大模型", "从自然语言中提取决定、待办、人员、时间、风险和对应原文", TEAL),
        ("规则", "检查缺失字段，把相对日期换成具体日期，检查格式", YELLOW),
        ("产品系统", "展示、编辑、确认、保留原文并导出纪要", GREEN),
        ("项目经理", "判断结果是否正确，补充信息并最终发布", CORAL),
    ], ["AI、规则与人工怎样分工"],
       "第一版不需要智能体自主执行，也不需要知识库。", cols=4)
    slide_flow(prs, 17, "第一版的范围", "只保留能验证核心问题的四步", [
        ("粘贴转写", "会议名称、日期和正文"),
        ("提取内容", "决定、待办和风险"),
        ("标记问题", "缺失、不确定和冲突"),
        ("人工确认", "修改并导出纪要"),
    ], ["MVP做什么"],
       "录音、转写、消息推送、任务系统和趋势看板暂不进入第一版。", cols=4)
    slide_compare(prs, 18, "两个核心页面", "原型先把主要操作和关键状态讲清楚",
                  "输入页", ["会议名称", "会议日期", "粘贴转写", "开始整理"],
                  "结果确认页", ["决定、待办和风险", "负责人和截止时间", "对应原文", "编辑、删除、确认和导出"],
                  ["核心页面"], GREEN,
                  "至少展示正常结果、字段缺失和人工修改三个状态。")
    slide_statement(prs, 19, "会议转写输入",
                    "支付页面已经评审通过，小李今天把最终稿发给开发",
                    "输入材料要像真实会议表达，而不是为了模型方便而写成标准表单。",
                    ["一个输入输出示例"], accent=YELLOW)
    slide_table(prs, 20, "结构化输出", "同一段转写会拆成不同类型的信息", [
        ["类型", "内容", "负责人", "截止时间", "状态"],
        ["风险", "登录接口本周可能无法完成", "未明确", "本周", "待确认"],
        ["待办", "与后端确认登录接口进度", "小王", "本周四", "待执行"],
        ["条件决定", "周四无结果则联调调整至下周一", "未明确", "周四确认", "待确认"],
        ["决定", "支付页面通过评审", "—", "—", "已确认"],
        ["待办", "将支付页面最终稿发送给开发", "小李", "今天", "待执行"],
    ], [1.18, 5.18, 1.35, 1.45, 1.55], ["一个输入输出示例"], size=12.4,
       row_h=0.56, takeaway_text="不能只看文字通不通顺，还要检查遗漏、误判和编造。")
    slide_cards(prs, 21, "最小测试集", "第一版先用20段仿真或脱敏会议片段", [
        ("明确任务", "负责人和日期都清楚"),
        ("字段缺失", "缺负责人或截止时间"),
        ("相对日期", "今天、明天、下周一"),
        ("混合表达", "一句话同时包含任务和风险"),
        ("没有结论", "讨论了问题但没有作出决定"),
        ("前后修改", "同一决定被重复或改动"),
        ("行动项召回", "应该提取的任务有没有漏"),
        ("人工修改率", "项目经理需要改多少内容"),
    ], ["怎样建立最小测试集"],
       "测试样本要覆盖真实变化，不只准备最容易成功的句子。", cols=4)
    slide_quote_analysis(prs, 22, "一个典型Bad Case", "这个暂时先不做，等客户确认以后再说", [
        ("错误输出", "模型写成：推进该功能开发"),
        ("问题原因", "没有理解否定和前置条件"),
        ("产品修改", "增加暂缓和条件事项类型"),
        ("使用控制", "保留原文，人工确认后再导出"),
    ], ["可能出现的Bad Case"], quote_label="原始表达",
       takeaway_text="发现错误以后，要修改字段、Prompt或流程。")
    slide_cards(prs, 23, "小项目的完整性", "没有复杂技术，也能形成完整产品闭环", [
        ("真实任务", "用户、场景和当前流程明确"),
        ("输入可得", "会议转写能够准备和测试"),
        ("分工清楚", "模型提取，规则检查，人工确认"),
        ("范围可控", "第一版只有两页和四步"),
        ("结果可测", "能够建立样本并检查错误"),
        ("可以修改", "Bad Case会推动字段和流程变化"),
    ], ["为什么它小但完整"],
       "复杂度可以低，项目的前后关系不能断。", cols=3)
    slide_table(prs, 24, "三种项目描述", "同一个方向，完整度差别很大", [
        ["版本", "描述", "当前状态"],
        ["A 项目想法", "我想做一个AI智能客服", "只有方向，不能立项"],
        ["B 功能方案", "支持知识问答、订单查询、投诉和推荐", "知道功能，缺少优先级、材料和责任"],
        ["C 可推进项目", "聚焦订单查询与异常分流，接入订单接口，异常转人工，并用历史问句测试", "具备继续做PRD和原型的条件"],
    ], [1.55, 7.00, 2.97], ["三个版本的项目对比"], size=13.0, row_h=0.82,
       highlight=2, takeaway_text="立项条件来自具体问题、可用材料、处理流程和测试方法。")
    slide_compare(prs, 25, "项目材料的可信度", "页数和技术名词并不能代表项目质量",
                  "看起来很多", ["行业介绍很长", "竞品和功能清单很多", "大量内容由模型一次生成", "缺少用户、样本和修改过程"],
                  "真正做过", ["说得清访谈或资料来源", "知道删掉了什么功能", "拿得出测试样本和错误", "保留修改前后的版本"],
                  ["判断完整度时不要只看材料数量"], GREEN,
                  "判断重点是做过哪些选择，后来为什么改变。")

    # 第三部分
    slide_compare(prs, 26, "课程项目怎样写进简历", "可以写，但要说明项目性质",
                  "可以这样标注", ["AI产品个人项目", "AI产品课程实战", "某行业业务改造方案", "基于公开资料的产品设计"],
                  "不要这样包装", ["冒充公司真实上线项目", "编造用户规模和业务收益", "把公共项目直接写成个人工作", "被追问时只会复述PRD"],
                  ["课程项目能不能写进简历"], CORAL,
                  "项目性质说清楚，不会削弱真正完成过的工作。")
    slide_compare(prs, 27, "两种简历描述", "面试官更想看到实际做了什么",
                  "描述 A", ["负责 AI 智能客服产品设计", "使用知识检索、智能体和大模型", "提升客服效率"],
                  "描述B", ["聚焦电商售后订单查询", "拆解意图识别、订单接口和异常分流", "用历史问句测试多意图和信息缺失", "补充澄清与转人工策略"],
                  ["简历描述应该呈现什么"], GREEN,
                  "技术名词少一些，项目过程更具体。")
    slide_cards(prs, 28, "能看出个人工作的材料", "项目过程越具体，越容易说明这是自己做的", [
        ("选题变化", "最初想做什么，后来为什么收缩"),
        ("业务材料", "访谈、公开资料、工作经验或脱敏文档"),
        ("版本记录", "V0.1 / V0.2 / V1.0 及修改原因"),
        ("功能取舍", "考虑过什么，为什么没有放进 MVP"),
        ("测试记录", "测试样本、结果和典型错误"),
        ("原型与Demo", "正常、异常、不确定和人工确认状态"),
        ("个人判断", "哪些是公共方法，哪些是个人设计"),
    ], ["什么内容可以证明项目是自己做的"],
       "面试时能讲清过程，比项目名称听起来宏大更重要。", cols=4)
    slide_cards(prs, 29, "项目中不能虚构的内容", "没有真实上线时，不能写成真实业务结果", [
        ("用户规模", "覆盖10万企业用户", RED),
        ("效果提升", "准确率提升35%", RED),
        ("商业收益", "节省公司500万元", RED),
        ("推广范围", "已在集团多个部门推广", RED),
    ], ["不能虚构什么"],
       "可以写测试结果和价值估算，但必须说明验证条件。", cols=4)
    slide_table(prs, 30, "真实结果、测试结果和估算", "三类信息要分别表达", [
        ["类型", "可以怎样写", "需要说明"],
        ["真实结果", "系统上线后的实际用户、效果或收益", "数据来源、时间范围和口径"],
        ["测试结果", "基于80条脱敏样本完成测试", "样本来源、指标和发现的问题"],
        ["价值估算", "根据人工步骤估算可能节省的时间", "尚未通过生产数据验证"],
    ], [1.45, 6.10, 3.97], ["不能虚构什么"], size=13.0, row_h=0.78,
       takeaway_text="如实说明边界，反而更容易获得信任。")
    slide_cards(prs, 31, "面试官常见追问", "真正做过项目的人通常能讲清当时的判断条件", [
        ("用户", "为什么选择这个用户？"),
        ("问题", "你怎么知道问题存在？"),
        ("材料", "输入材料具体长什么样？"),
        ("技术", "为什么用知识检索或规则？"),
        ("范围", "为什么没进入第一版？"),
        ("测试", "样本是谁整理的？"),
        ("错误", "最严重的问题是什么？"),
        ("修改", "后来具体改了哪一部分？"),
    ], ["面试官会怎样判断真实性"],
       "回答可以不完美，但必须回到真实信息和实际限制。", cols=4)
    slide_flow(prs, 32, "一段可信的项目描述", "简历只保留最关键的项目过程", [
        ("面向谁", "用户和具体任务"),
        ("解决什么", "当前流程最麻烦的地方"),
        ("做了什么", "关键产品工作"),
        ("怎样分工", "模型、规则和人工"),
        ("怎样验证", "样本、指标和错误"),
        ("怎样修改", "版本变化和原因"),
    ], ["简历描述应该呈现什么"],
       "把自己真正完成过的内容写清楚，不需要夸大上线结果。", cols=6)

    # 第四部分
    slide_statement(prs, 33, "三个公共项目的选择逻辑",
                    "三个项目分别训练服务、核验和专业内容生成",
                    "它们代表三类常见企业问题：开放表达怎样进入稳定服务流程，复杂材料怎样被可靠核验，分散知识怎样形成有依据的专业文档。",
                    ["为什么不是随便选三个热门项目"])
    slide_cards(prs, 34, "三个公共项目", "项目越来越复杂，产品责任也越来越重", [
        ("企业智能客服与业务办理", "开放对话、知识问答、业务查询和人工接管", TEAL),
        ("业务材料审核与风险核验", "文档提取、规则核对、风险提示和人工复核", CORAL),
        ("专业知识研究与方案生成", "长文档理解、知识检索、结构生成和事实检查", YELLOW),
    ], ["三个公共项目分别训练什么"],
       "三个项目都会同时调用多种AI和产品能力。", cols=3)
    slide_cards(prs, 35, "企业智能客服与业务办理", "电商售后只是公共案例，重点是服务流程", [
        ("用户表达", "同一个问题有很多说法，信息经常不完整"),
        ("知识来源", "产品资料、售后制度和FAQ分散"),
        ("实时数据", "订单、物流和账户状态需要查业务系统"),
        ("人工接管", "投诉、异常和高风险问题不能只靠自由对话"),
    ], ["公共项目一：企业智能客服与业务办理助手"],
       "客服产品要真正解决问题，不只是回答得像人。", cols=4)
    slide_flow(prs, 36, "智能客服的处理流程", "对话最终要进入可控的知识、查询或办理路径", [
        ("发起咨询", "自然语言描述问题"),
        ("识别意图", "补齐必要信息"),
        ("选择路径", "问答、查询或办理"),
        ("调用能力", "知识、接口和规则"),
        ("返回结果", "回答或引导操作"),
        ("转交人工", "不确定、异常和高风险"),
    ], ["公共项目一：企业智能客服与业务办理助手"],
       "人工接管时要把已经收集的信息一并交过去。", cols=6)
    slide_cards(prs, 37, "智能客服项目的练习重点", "从开放语言走到稳定服务", [
        ("意图与澄清", "识别需求，缺信息时继续追问"),
        ("知识和数据", "区分知识库回答与实时系统查询"),
        ("流程收束", "重要操作使用规则、按钮和表单"),
        ("异常处理", "低置信度、知识缺失和情绪问题"),
        ("人工接管", "触发条件和上下文交接"),
        ("产品评估", "解决率、完成率、转人工和典型错误"),
    ], ["公共项目一：企业智能客服与业务办理助手"],
       "核心能力是把用户表达转成可以执行、可以负责的服务流程。", cols=3)
    slide_cards(prs, 38, "采购单据核验的业务场景", "一笔采购会经过多份材料和多个系统", [
        ("业务材料", "采购申请、订单、入库单、发票和ERP记录"),
        ("人工工作", "提取供应商、日期、数量和金额"),
        ("核验任务", "判断材料是否完整，前后是否一致"),
        ("风险要求", "差异必须能回到原始材料，由审核人员确认"),
    ], ["公共项目二：AI业务材料审核与风险核验系统"],
       "公共案例关注采购穿行中的单据核验和风险提示。", cols=4)
    slide_flow(prs, 39, "材料核验的处理流程", "字段提取和业务核验是两类不同工作", [
        ("上传材料", "业务数据和单据"),
        ("识别类型", "判断是什么文档"),
        ("提取字段", "金额、日期、数量等"),
        ("关联业务", "找到同一笔采购"),
        ("规则核对", "比较字段和必要材料"),
        ("标记风险", "展示差异和原始位置"),
        ("人工复核", "审核人员确认"),
    ], ["公共项目二：AI业务材料审核与风险核验系统"],
       "模型负责理解材料，确定性的金额和日期比较交给规则。", cols=4)
    slide_cards(prs, 40, "材料核验项目的练习重点", "准确性和责任边界要求更高", [
        ("识别与提取", "文字识别、视觉模型和大模型怎样配合"),
        ("字段字典", "每类单据需要哪些字段"),
        ("业务关系", "不同材料怎样对应同一笔业务"),
        ("核验规则", "金额、日期、数量和缺失材料"),
        ("风险处理", "低置信度、误报、漏报和人工确认"),
        ("两层测试", "字段提取准确度和风险判断效果"),
    ], ["公共项目二：AI业务材料审核与风险核验系统"],
       "AI提示风险，专业人员承担最终判断。", cols=3)
    slide_cards(prs, 41, "RFP需求拆解与智能标书", "先理解复杂要求，再生成专业内容", [
        ("外部要求", "截止时间、资格、强制项、参数和评分标准"),
        ("内部知识", "企业资质、产品资料、历史案例和方案能力"),
        ("主要风险", "遗漏要求、引用错误、事实编造和前后冲突"),
        ("最终产出", "需求清单、材料缺口、目录和分章节初稿"),
    ], ["公共项目三：专业知识研究与方案生成平台"],
       "一次性生成长文档很容易漏项，必须拆成多个可检查步骤。", cols=4)
    slide_flow(prs, 42, "专业方案生成的处理流程", "从要求清单开始，逐步走到完整文档", [
        ("上传RFP", "几十页到几百页"),
        ("提取要求", "资格、强制项和评分点"),
        ("标记位置", "保留原文页码或段落"),
        ("检索资料", "资质、产品和案例"),
        ("生成目录", "一级和二级结构"),
        ("分段写作", "按章节生成"),
        ("检查内容", "遗漏、引用和事实"),
        ("人工校对", "补充并最终确认"),
    ], ["公共项目三：专业知识研究与方案生成平台"],
       "先抽取和核对要求，再开始生成。", cols=4)
    slide_cards(prs, 43, "专业方案生成项目的练习重点", "信息量更大，过程必须可检查", [
        ("长文档理解", "处理复杂结构和大量要求"),
        ("要求抽取", "先形成完整清单和材料缺口"),
        ("知识检索", "使用企业资料并保留来源"),
        ("树状拆解", "从目录走到章节和段落"),
        ("缺失处理", "材料不足时要求人工补充"),
        ("结果测试", "遗漏、错误引用、编造和冲突"),
    ], ["公共项目三：专业知识研究与方案生成平台"],
       "核心能力是把复杂要求转成有依据、可复核的专业内容。", cols=3)
    slide_table(prs, 44, "三个必修项目和两个选修方向", "必修负责主干能力，选修补充个性化方向", [
        ["项目", "主要输入", "重点练习"],
        ["智能客服与业务办理", "用户对话、企业知识和业务数据", "对话理解、流程控制和人工接管"],
        ["材料审核与风险核验", "业务数据、单据和审核规则", "模型、规则和人工可靠协作"],
        ["专业研究与方案生成", "复杂要求、企业知识和专业资料", "长文档理解、生成和校验"],
        ["图像判断与报告生成", "图片和业务标准", "多模态理解和结构化报告"],
        ["客户研究与销售方案", "外部研究、产品资料和客户目标", "信息研究、产品匹配和推荐"],
    ], [3.10, 4.55, 3.87], ["三个项目组成什么能力", "两个选修项目补充什么"],
       size=12.6, row_h=0.56, takeaway_text="三个个人项目保持不变，选修方向按个人需要使用。")

    # 第五部分
    slide_compare(prs, 45, "项目迁移和改名的区别", "业务内容需要重新建立",
                  "只改名称", ["把电商客服改成SaaS客服", "功能、流程和测试全部照搬", "不知道账号、权限和工单怎样处理"],
                  "真正迁移", ["重新确定用户和使用时刻", "找到知识、业务数据和真实问句", "重新设计权限、风险和转人工条件", "建立自己的测试样本"],
                  ["迁移不是改名字"], GREEN,
                  "方法可以复用，项目答案必须来自新的业务场景。")
    slide_cards(prs, 46, "三类公共项目的个人化方向", "每一类结构都能迁移到不同业务", [
        ("对话与办理", "SaaS帮助台、员工服务、物业报修、课程咨询、保险咨询", TEAL),
        ("材料审核", "报销、理赔、合同履约、供应商资质、贷款材料", CORAL),
        ("专业方案", "项目申报、客户研究、咨询方案、尽调、安全问卷", YELLOW),
    ], ["从公共项目一迁移", "从公共项目二迁移", "从公共项目三迁移"],
       "迁移后要重新说明角色、输入、规则、风险、输出和测试。", cols=3)
    slide_cards(prs, 47, "个人项目的三个入口", "优先寻找三者的交集", [
        ("过去的工作", "从重复、耗时或容易出错的任务开始", TEAL),
        ("可以获得的材料", "客服记录、业务单据、政策、手册、历史方案或用户反馈", YELLOW),
        ("目标岗位", "查看岗位希望你证明哪些业务和AI产品能力", CORAL),
    ], ["选择个人项目时优先从哪里开始"],
       "理解业务、拿得到材料、又能服务职业目标的方向最适合。", cols=3)

    # 第六部分
    slide_cards(prs, 48, "选题的六项标准", "每一项都回答项目能不能真正做下去", [
        ("业务熟悉", "能否还原用户现在的工作流程"),
        ("用户可达", "能否接触真实使用者或可靠从业者"),
        ("材料可得", "能否准备真实、脱敏或合理仿真的输入"),
        ("AI必要", "关键环节是否需要理解、提取、检索或生成"),
        ("范围可控", "第一版能否只完成一个明确任务"),
        ("职业相关", "能否证明目标岗位需要的能力"),
    ], ["怎样判断一个选题是否值得做"],
       "任何一项明显缺失，都要先调整选题。", cols=3)
    slide_table(prs, 49, "六项评分表", "每项0到2分，用来发现选题缺少什么", [
        ["标准", "0分", "1分", "2分"],
        ["业务熟悉", "基本不了解", "有一般认知", "有经验或可靠资料"],
        ["用户可达", "无法接触", "可以间接了解", "可以访谈或持续验证"],
        ["材料可得", "没有材料", "少量公开材料", "有可用或脱敏材料"],
        ["AI必要", "规则即可完成", "AI有一定价值", "AI明显改善关键环节"],
        ["范围可控", "范围巨大", "需要明显收缩", "可以形成最小闭环"],
        ["职业相关", "关系较弱", "部分相关", "高度相关"],
    ], [1.85, 3.08, 3.08, 3.51], ["六项评分表"], size=12.1, row_h=0.50,
       takeaway_text="9到12分可继续，6到8分先修改，0到5分暂缓立项。")
    slide_table(prs, 50, "三个选题的评分示例", "简单项目也可能比宏大项目更适合第一次练习", [
        ["选题", "业务", "用户", "材料", "AI", "范围", "职业", "总分"],
        ["AI老板助手", "0", "0-1", "0", "1", "0", "1", "2-3"],
        ["保险顾问家庭保障方案助手", "2", "2", "2", "2", "1", "2", "11"],
        ["项目例会待办整理助手", "2", "2", "2", "2", "2", "1", "11"],
    ], [3.25, 1.10, 1.10, 1.10, 1.10, 1.10, 1.10, 1.67], ["三个示例"],
       size=12.2, row_h=0.78, highlight=2,
       takeaway_text="评分不是预测模型，只是一张选题检查表。")
    slide_cards(prs, 51, "需要暂停的四类选题", "暂停的目的是换用户、缩范围、补材料或调整责任", [
        ("没有用户和材料", "只有行业想象，没有任何验证入口", RED),
        ("关键数据拿不到", "项目依赖医疗、风控或内部机密数据", RED),
        ("第一版覆盖全流程", "一次想做完客服、销售、营销和运营", RED),
        ("高风险自动决策", "自动诊断、拒赔、授信或审计，却没有人工确认", RED),
    ], ["四种需要直接暂停的选题"],
       "暂时不立项，不代表这个方向永远不能做。", cols=4)

    # 第七部分
    slide_flow(prs, 52, "现场诊断的八个问题", "所有项目按照同一顺序检查", [
        ("使用者", "实际打开产品的人"),
        ("使用时刻", "任务发生在什么时候"),
        ("当前做法", "没有产品时怎样工作"),
        ("主要问题", "第一版只选一个"),
        ("AI作用", "具体进入哪一步"),
        ("材料来源", "文档、字段和样本"),
        ("最小闭环", "输入、确认和可用输出"),
        ("删除内容", "主动缩小第一版"),
    ], ["每个项目按照同一顺序诊断"],
       "诊断先找缺口，再讨论具体功能。", cols=4)
    slide_cards(prs, 53, "课堂选择与项目收缩", "看完案例后，先让全班做判断", [
        ("A 可以进入项目", "用户、问题、材料和第一版已经明确", GREEN),
        ("B 需要缩小范围", "方向成立，但第一版包含太多任务", YELLOW),
        ("C 更换核心问题", "当前问题无法形成清楚的产品目标", CORAL),
        ("D 不适合大模型", "规则、搜索或普通系统已经足够", TEAL),
        ("E 缺少材料", "暂时无法判断效果和可行性", RED),
        ("收缩示例", "AI医疗助手可缩成营养师饮食记录整理与随访摘要", TEAL),
    ], ["让全班先做选择", "备用诊断案例一：AI医疗助手"],
       "选择之后再说明理由，重点是使用共同的判断条件。", cols=3)
    slide_cards(prs, 54, "项目一选题卡", "第一课结束前，每个人都要写出一版", [
        ("用户与场景", "项目名称、目标用户、发生场景、当前做法"),
        ("问题与材料", "最主要的问题、可以获得的材料或数据"),
        ("产品方案", "AI任务、最终输出和人工环节"),
        ("第一版边界", "明确不做什么、对应哪类公共项目"),
        ("当前判断", "六项总分、尚未确认的假设"),
    ], ["项目一选题卡"],
       "选题卡不是完整的产品需求文档，它只确认项目能不能继续。", cols=3)
    slide_table(prs, 55, "选题卡的三种处理结果", "老师根据项目是否具备推进条件给出处理意见", [
        ["处理结果", "常见情况", "下一步"],
        ["可以通过", "用户、场景、材料、AI任务、输出和人工责任都较明确", "进入公共项目学习并继续补充"],
        ["修改后通过", "方向合理，但范围偏大、任务不具体或材料不足", "按点评意见修改后继续"],
        ["暂不通过", "业务不清、材料拿不到、不需要AI或高风险无人确认", "换问题、补材料或重新选题"],
    ], [2.0, 6.65, 2.87], ["教师怎样判断选题卡"], size=12.7,
       row_h=0.76, highlight=1,
       takeaway_text="没有通过也要说清楚具体卡点，不能只写“还没想好”。")

    # 第八部分
    slide_flow(prs, 56, "最近三次课怎样衔接", "第一课只完成立项，不要求立刻写完整PRD", [
        ("第一课", "确定初步选题"),
        ("第二课", "学习智能客服公共项目"),
        ("24小时", "修订个人选题卡"),
        ("3天内", "提交项目骨架"),
        ("第三课前", "提交项目V0.1"),
        ("第三课", "点评并现场修改"),
        ("课后", "提交V1.0和修改说明"),
    ], ["最近三次课的关系"],
       "先建立参照，再修改个人项目，避免直接复制公共项目。", cols=4)
    slide_compare(prs, 57, "课后24小时和3天", "两个时间点解决不同问题",
                  "24小时内修订选题卡", ["用户是否具体", "输入材料是否真实存在", "知识、规则和人工怎样分工", "第一版是否仍然过大"],
                  "3天内提交项目骨架", ["背景、用户和当前流程", "主要问题、输入和输出", "初步能力分工", "MVP和不做清单"],
                  ["24小时内修订选题卡", "课后3天提交项目骨架"], GREEN,
                  "先发现方向错误，再投入时间写长文档。")
    slide_cards(prs, 58, "项目一 V0.1", "点评课前要形成一份可以被批评和修改的初稿", [
        ("项目说明", "背景、用户、流程与问题"),
        ("方案边界", "AI介入理由、输入输出、人机分工和MVP"),
        ("产品表达", "核心流程和1到3张核心页面"),
        ("初步测试", "10条样本和目前尚未解决的问题"),
    ], ["点评课前提交V0.1"],
       "V0.1的任务是暴露问题，不是看起来像最终答案。", cols=4)
    slide_flow(prs, 59, "点评课怎样进行", "全班用同一组项目判断方法参与", [
        ("汇总问题", "先看共性"),
        ("展示作业", "选择代表案例"),
        ("全班判断", "先做选择"),
        ("老师讲解", "说明问题和方向"),
        ("同步修改", "各自修改个人项目"),
        ("版本对比", "展示修改前后"),
    ], ["点评课怎样进行"],
       "点评重点是让每个人学会修改自己的项目。", cols=6)
    s = base(prs, 60, "修改记录和本课总结", "项目能力会在一次次修改中变得可见")
    card(s, 0.82, 2.04, 3.55, 3.48, "V0.1", "客户研究、线索评分、方案推荐、自动发邮件、CRM回写和销售看板全部进入第一版", "01", CORAL, 18, 14)
    card(s, 4.88, 2.04, 3.55, 3.48, "点评意见", "范围过大。核心问题是客户研究和方案准备，不是完整销售管理", "02", YELLOW, 18, 14)
    card(s, 8.79, 2.04, 3.55, 3.48, "V1.0", "只保留目标企业研究、带来源的客户简报、有限知识检索、方案建议和人工确认", "03", GREEN, 18, 14)
    takeaway(s, "最后三问", "谁遇到什么问题？材料从哪里来？第一版最少做到什么？")
    notes_from(s, "收尾时不要重复知识点。让学员回到自己的项目，回答最后三个问题，并提醒他们保留V0.1、点评意见、V1.0和修改原因。", "为什么要保留修改记录", "本课总结")

    dk.declare_delivery(OUT, "presented")
    dk.lint_layout(prs, strict=True)
    prs.save(OUT)

    slide_records = []
    titles = [
        "AI项目立项、选题与完成标准", "第二阶段的学习任务", "从知识点到完整项目", "三个个人项目",
        "三个项目的完成深度", "两个选修方向", "公共项目和个人项目", "第二阶段的学习方式",
        "从一句项目名称开始", "功能清单还缺什么", "PRD和Demo各自能说明什么", "一个完整项目需要回答的问题",
        "项目例会待办整理助手的业务背景", "项目定义", "一段会议转写里有什么", "AI、规则、系统和人工的分工",
        "第一版的范围", "两个核心页面", "会议转写输入", "结构化输出", "最小测试集", "一个典型Bad Case",
        "小项目的完整性", "三种项目描述", "项目材料的可信度", "课程项目怎样写进简历", "两种简历描述",
        "能看出个人工作的材料", "项目中不能虚构的内容", "真实结果、测试结果和估算", "面试官常见追问",
        "一段可信的项目描述", "三个公共项目的选择逻辑", "三个公共项目", "企业智能客服与业务办理",
        "智能客服的处理流程", "智能客服项目的练习重点", "采购单据核验的业务场景", "材料核验的处理流程",
        "材料核验项目的练习重点", "RFP需求拆解与智能标书", "专业方案生成的处理流程", "专业方案生成项目的练习重点",
        "三个必修项目和两个选修方向", "项目迁移和改名的区别", "三类公共项目的个人化方向", "个人项目的三个入口",
        "选题的六项标准", "六项评分表", "三个选题的评分示例", "需要暂停的四类选题", "现场诊断的八个问题",
        "课堂选择与项目收缩", "项目一选题卡", "选题卡的三种处理结果", "最近三次课怎样衔接",
        "课后24小时和3天", "项目一V0.1", "点评课怎样进行", "修改记录和本课总结",
    ]
    for i, title in enumerate(titles, 1):
        slide_records.append({
            "slide": i,
            "role": "cover" if i == 1 else ("closing" if i == 60 else "content"),
            "takeaway": title,
            "evidence": [f"{SOURCE}｜对应讲义章节与示例"],
            "units": 1 if i in (1, 2, 14, 19, 33, 60) else 3,
        })
    gates = {
        "delivery": "presented",
        "content": {
            "checkpoint": {"mode": "approved", "record": "用户已确认以最新讲义为内容来源，沿用先案例、再判断、最后总结方法的完整PPT结构。"},
            "arc": {
                "chosen": "小案例带路，再进入三个公共项目和个人选题",
                "candidates": [
                    {"name": "小案例带路，再进入三个公共项目和个人选题", "shape": "case-first", "roles": ["setup", "case", "method", "projects", "practice", "roadmap"], "audience_question": "怎样判断并启动第一个个人AI项目？", "objection": "直接讲立项标准会显得抽象。", "closing_ask": "提交项目一选题卡并保留修改过程。", "evidence": ["讲义第一至第八部分"]},
                    {"name": "先讲完整标准，再逐项举例", "shape": "framework-first", "roles": ["framework", "examples", "practice"], "audience_question": "完整项目包含哪些部分？", "objection": "标准较多，需要系统呈现。", "closing_ask": "按清单补齐项目。", "evidence": ["讲义第二、六部分"]},
                    {"name": "先介绍三个项目，再回到个人选题", "shape": "portfolio-first", "roles": ["projects", "comparison", "selection", "roadmap"], "audience_question": "第二阶段要做哪三个项目？", "objection": "学员需要先看到课程全貌。", "closing_ask": "从三个方向中选择项目。", "evidence": ["讲义第一、四、五部分"]},
                ],
                "rejected": [
                    {"name": "先讲完整标准，再逐项举例", "why_lost": "开场抽象，容易出现说教感。"},
                    {"name": "先介绍三个项目，再回到个人选题", "why_lost": "项目介绍过早，学员还没有完整度判断标准。"},
                ],
            },
            "slides": slide_records,
        },
        "design_plan": {
            "mode": "Mode A mimic",
            "style_pick": "n/a — locked mimic of the user-approved Lesson 11 and Lesson 12 visual system",
            "concept": {"chosen": "一条从项目想法逐步收束到可推进项目的轨道", "rejected": [{"concept": "项目工具箱", "why_lost": "容易变成功能和模板罗列。"}, {"concept": "职业作品集画廊", "why_lost": "会削弱立项判断和课堂练习。"}]},
            "boldness": "balanced+",
            "signature_move": "把项目从一句名称到可推进方案的变化做成贯穿全课的进度轨道",
            "carried_by": [12, 24, 56],
            "form_ledger": "陈述页6；比较页12；流程轨道页11；表格页8；案例拆解页5；多项框架页17；封面1",
            "interior_register": "加大章节标题、底部进度线、深色画布、青绿结构色和珊瑚提醒色贯穿全部正文页",
            "icon_family": "none — Mode A模板锁定，既有系列课使用编号、色块和流程轨道承担分类，增加图标会改变已确认的课程识别",
            "icon_none_category": "template-locked",
            "icon_none_checked": ["slides 2-60"],
            "palette": {"fill_only": ["#031016", "#102B37", "#133843"], "text_safe": ["#F5F2E9", "#A7B5BC", "#36D6C2", "#FF7657", "#F0C565"], "semantic": "青绿表示结构和正常推进，珊瑚表示提醒和风险，黄色表示待确认，绿色表示可继续"},
            "type_scale": {"display": 42, "title": 29, "body": 15},
            "motif_generates": {"background": "深色画布保持系列识别", "markers": "两位编号表示项目步骤和判断项", "page": "第12、24、56页把进度轨道变成项目链路和课程节奏"},
            "density": {"planned_median": 34, "over_70": 2, "non_text_protagonist": 49},
            "build_shape": "solo — host instruction prohibits subagent delegation; one locked reference style and one tightly coupled source lecture",
            "build_script": str(Path(__file__).resolve()),
            "checkpoint": {"mode": "approved", "record": "用户要求沿用最近课程PPT风格，并已确认案例先行、自然中文、强化章节结构。"},
            "material_probe": {"png": str(ROOT / "render" / "slide12.png"), "safe_version": "安全版本会把完整项目列成普通清单，无法看出前后怎样衔接。"},
            "signature_proof": [
                {"role": "signature", "slide": 12, "png": str(ROOT / "render" / "slide12.png")},
                {"role": "complex", "slide": 49, "png": str(ROOT / "render" / "slide49.png")},
                {"role": "data", "slide": 56, "png": str(ROOT / "render" / "slide56.png")},
            ],
            "form_reach": {"waived": "本课复用学员已经熟悉的系列教学模板。表格使用原生组件，其余页面按内容需要组合比较、流程和案例，不以组件数量作为设计目标。"},
        },
        "design": {"checkpoint": {"mode": "approved", "record": "Mode A沿用第11、12课已确认的深色青绿视觉体系。"}},
        "content_plan_ref": str(SOURCE),
        "output": str(OUT),
        "provenance": {"waived": "PPT仅改编用户已确认的本地讲义，不新增需要联网核验的外部事实。"},
        "density": {"waived": "这是老师现场讲授的课程PPT。主要页面只保留核心信息，完整解释放在演讲者备注中。"},
    }
    existing = {}
    if GATES.exists():
        try:
            existing = json.loads(GATES.read_text(encoding="utf-8"))
        except Exception:
            existing = {}
    existing.update(gates)
    GATES.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    build_deck()
