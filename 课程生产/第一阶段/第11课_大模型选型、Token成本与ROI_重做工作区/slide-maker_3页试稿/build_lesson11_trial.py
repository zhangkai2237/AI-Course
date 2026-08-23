#!/usr/bin/env python3
"""第11课 3页 slide-maker 试稿。"""
from pathlib import Path
import sys

SKILL = Path("/Users/keivn/.codex/skills/slide-maker")
sys.path.insert(0, str(SKILL / "scripts"))
import deckkit as dk  # noqa: E402


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "第11课_slide-maker_3页试稿.pptx"
SOURCE = Path("/Users/keivn/Project/AI-Course/课程生产/第一阶段/第11课_大模型选型、Token成本与ROI_重做工作区/第11课_大模型选型、Token成本与ROI_完整讲义.md")

W, H = 13.333, 7.5
BG = dk.RGBColor(0x03, 0x10, 0x16)
PANEL = dk.RGBColor(0x10, 0x2B, 0x37)
PANEL_HI = dk.RGBColor(0x13, 0x38, 0x43)
LINE = dk.RGBColor(0x41, 0x6A, 0x78)
TEAL = dk.RGBColor(0x36, 0xD6, 0xC2)
CORAL = dk.RGBColor(0xFF, 0x76, 0x57)
TEXT = dk.RGBColor(0xF5, 0xF2, 0xE9)
MUTE = dk.RGBColor(0xA7, 0xB5, 0xBC)
GHOST = dk.RGBColor(0x41, 0x6A, 0x78)


dk.set_palette(
    deep=TEXT,
    blue=TEAL,
    teal=TEAL,
    magenta=CORAL,
    slate=MUTE,
    mute=MUTE,
    mono="Menlo",
    font="Helvetica Neue",
    display="Helvetica Neue",
    eadisplay="Hiragino Sans GB",
    eafont="Hiragino Sans GB",
    accents=[TEAL, CORAL],
)
dk.set_ground(BG)
dk.set_geometry(radius=1.0, rule_w=0.85)


def txt(slide, x, y, w, h, value, size, color=TEXT, bold=False,
        align=dk.PP_ALIGN.LEFT, anchor=dk.MSO_ANCHOR.TOP, wrap=True):
    return dk.text(
        slide, x, y, w, h,
        [[(value, size, color, bold, False, "Helvetica Neue", "Hiragino Sans GB")]],
        align=align, anchor=anchor, space_after=0, wrap=wrap,
    )


def add_chrome(slide, kicker, page, total=3):
    txt(slide, 0.78, 0.32, 8.5, 0.28, kicker, 11, TEAL, True)
    txt(slide, 12.05, 0.31, 0.52, 0.28, f"{page:02d}", 11, MUTE, True,
        align=dk.PP_ALIGN.RIGHT)
    dk.box(slide, 0.76, 7.17, 11.82, 0.012, fill=LINE)
    prog = dk.box(slide, 0.76, 7.155, 11.82 * page / total, 0.038, fill=TEAL)
    dk.tag_motif(prog, loud=False)


def add_section_title(slide, title, subtitle=None):
    txt(slide, 0.78, 0.72, 11.7, 0.58, title, 29, TEXT, True)
    rule = dk.box(slide, 0.78, 1.39, 1.55, 0.045, fill=TEAL)
    dk.tag_motif(rule, loud=False)
    if subtitle:
        txt(slide, 0.78, 1.53, 11.6, 0.42, subtitle, 15.5, MUTE, False)


def slide_01(prs):
    """role=cover | form=course-series cover | static: single statement |
    takeaway='把“调用大模型”推进为可评审的选型与价值决策'"""
    s = dk.add_slide(prs)
    spine = dk.box(s, 0, 0, 0.22, H, fill=TEAL)
    dk.tag_motif(spine, loud=True)
    txt(s, 0.92, 0.78, 5.6, 0.35, "AI PRODUCT MANAGER", 13, TEAL, True)
    dk.text(
        s, 0.92, 2.00, 10.8, 1.78,
        [
            [("大模型选型", 48, TEXT, True, False, "Helvetica Neue", "Hiragino Sans GB")],
            [("Token 成本与 ROI", 44, TEXT, True, False, "Helvetica Neue", "Hiragino Sans GB")],
        ],
        space_after=2, line_spacing=1.02,
    )
    txt(s, 0.96, 3.96, 2.1, 0.35, "第 11 课", 17, MUTE, False)
    orange = dk.box(s, 0.92, 4.72, 3.05, 0.05, fill=CORAL)
    dk.tag_motif(orange, loud=True)
    txt(s, 0.92, 5.24, 9.8, 0.42,
        "从“调用大模型”到一份可评审的产品决策", 18, MUTE, True)
    txt(s, 0.92, 6.26, 10.4, 0.3,
        "MODEL SELECTION · TOKEN COST · ROI · GO / NO-GO", 10.5, MUTE, True)
    txt(s, 10.64, 5.10, 1.75, 1.18, "11", 72, GHOST, True,
        align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE, wrap=False)
    dk.speaker_notes(s, f"""第10课的PRD已经写明：复杂语义判断交给大模型。但这里还留下三个产品问题：到底用哪个模型、完整成本是多少、项目值不值得做。第11课就是把这三个问题做成可评审的决策。\n\n[Sources]\n- {SOURCE}（第1行、第8—21行、第1923—1934行）""")


def slide_02(prs):
    """role=objective | form=three-stage course map | static: scan-at-once map |
    takeaway='课程依次回答怎么选、多少钱、值不值得做'"""
    s = dk.add_slide(prs)
    add_chrome(s, "AI PRODUCT · OVERVIEW", 2)
    add_section_title(s, "第11课课程结构", "模型选择 → 成本核算 → 项目决策")

    xs = [0.78, 4.80, 8.82]
    w, y, h = 3.52, 2.36, 2.56
    # 先画连接关系，再画节点，保证箭头落在节点后面。
    dk.arrow(s, 4.34, 3.40, 0.38, 0.26, color=TEAL)
    dk.arrow(s, 8.36, 3.40, 0.38, 0.26, color=TEAL)
    stages = [
        ("01", "怎么选", "第1—4节", "场景 · 标准\n策略 · 多模态"),
        ("02", "多少钱", "第5节", "Token · 单次任务\n月度 · 完整成本"),
        ("03", "值不值得做", "第6节", "优化 · ROI\nGo / No-Go / 试点"),
    ]
    for i, (num, title, sec, body) in enumerate(stages):
        x = xs[i]
        dk.box(s, x, y, w, h, fill=PANEL, line=LINE, line_w=1.0, round=True, r=0.18)
        txt(s, x + 0.30, y + 0.22, 0.60, 0.30, num, 13, CORAL, True)
        txt(s, x + 0.30, y + 0.65, w - 0.60, 0.48, title, 24, TEXT, True)
        txt(s, x + 0.30, y + 1.28, w - 0.60, 0.28, sec, 12, TEAL, True)
        txt(s, x + 0.30, y + 1.67, w - 0.60, 0.62, body, 17, MUTE, False)

    dk.callout(
        s, 0.78, 5.54, 11.56, 0.78,
        "最终产出", "模型选型与价值决策表：选什么、花多少、继续还是停止",
        label_c=CORAL, fill=PANEL_HI, body_c=TEXT,
    )
    dk.speaker_notes(s, f"""整节课可以先记成三个问题。第一，怎么选：不是看排行榜，而是建立门槛、比较维度和实施策略。第二，多少钱：从Token单价升级到一次业务任务、月度和完整成本。第三，值不值得做：把成本、收益和风险合在一起，给出Go、No-Go或有条件试点。\n\n[Sources]\n- {SOURCE}（六节标题；第1472—1482行；第1923—1934行）""")


def slide_03(prs):
    """role=worked-example | form=exact-value comparison table + decision criteria |
    static: comparison must be visible at once | takeaway='最强模型不等于最适合的模型'"""
    s = dk.add_slide(prs)
    add_chrome(s, "第一节 · MODEL SELECTION", 3)
    add_section_title(s, "模型选型的工作场景", "案例：同一个审核任务，三个候选模型怎么选？")

    rows = [
        ["候选模型", "准确率", "响应时间", "单次成本"],
        ["模型 A", "98%", "5 秒", "0.30 元"],
        ["模型 B", "96%", "1 秒", "0.03 元"],
        ["模型 C", "91%", "0.3 秒", "0.005 元"],
    ]
    dk.table(
        s, 0.82, 2.16, 7.28, rows,
        col_w=[2.18, 1.70, 1.55, 1.85],
        header=True, highlight=1, numeric_cols=[1, 2, 3],
        size=17, row_h=0.66,
        head_c=TEXT, body_c=MUTE, rule_c=LINE,
        hi_fill=PANEL_HI, hi_c=TEAL,
        font="Helvetica Neue",
    )
    txt(s, 8.72, 2.13, 3.67, 0.42, "先问这 7 个问题", 18.5, TEXT, True)
    dk.bullet(
        s, 8.74, 2.72, 3.55,
        [
            ("任务：", "具体审什么？"),
            ("风险：", "错一次代价多大？"),
            ("体验：", "用户能等多久？"),
            ("规模：", "每天多少条？"),
            ("接管：", "能否人工复核？"),
            ("数据：", "能否发送外部？"),
            ("预算：", "能承担多少？"),
        ],
        size=17, gap=0.16, marker=CORAL, lead_c=TEXT, body_c=MUTE,
    )
    dk.bottom_callout(
        s, 0.82, 11.50,
        "判断原则", "最强模型 ≠ 最适合的模型。先设业务门槛，再比较候选模型。",
        footer_gap=0.30, label_c=CORAL, fill=PANEL, body_c=TEXT,
    )
    dk.speaker_notes(s, f"""这三个模型没有脱离场景的唯一答案。模型A最准确，但慢且贵；模型C最快最便宜，但漏判风险更高；模型B在普通营销审核场景里可能更均衡。真正的顺序不是先挑赢家，而是先明确任务、风险、等待时间、业务量、人工接管、数据边界和预算，再筛选候选模型。表中数据是教学假设，不代表真实厂商报价。\n\n[Sources]\n- {SOURCE}（第31—63行，教学示例）""")


SLIDES = [slide_01, slide_02, slide_03]


def main():
    prs = dk.blank_deck(W, H)
    for build in SLIDES:
        build(prs)
    dk.retrofit_ea(prs, "Hiragino Sans GB")
    dk.lint_layout(prs, strict=True)
    prs.save(OUT)
    dk.declare_delivery(OUT, "presented")
    print(f"saved -> {OUT} | slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
