#!/usr/bin/env python3
"""FDE：产品经理在 AI 时代的全栈化生存 — 41 页正式授课版。"""
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
    blank_deck, add_slide, box, text, picture, arrow, node, connect_boxes,
    flow_chain, timeline, step_list, org_tree, tier_stack, meter_bar,
    icon_card, source_note, speaker_notes, design_intent, tag_motif,
    lint_layout, declare_delivery, set_palette, set_geometry, set_ground, _flat,
)

OUT = ROOT / 'FDE_产品经理在AI时代的全栈化生存_正式授课PPT_v1.0.pptx'
ASSETS = ROOT / 'assets'

# Scope Chain visual system
BG = RGBColor(245, 242, 234)
INK = RGBColor(19, 33, 30)
GREY = RGBColor(67, 82, 78)
MUTE = RGBColor(122, 136, 131)
LINE = RGBColor(201, 211, 206)
LIGHT = RGBColor(227, 235, 231)
WHITE = RGBColor(255, 255, 255)
TEAL = RGBColor(0, 124, 108)
ORANGE = RGBColor(230, 106, 58)
BLUE = RGBColor(67, 97, 168)
OCHRE = RGBColor(181, 139, 42)
ORANGE_DARK = RGBColor(194, 71, 28)
OCHRE_DARK = RGBColor(132, 96, 8)
DEEP = RGBColor(19, 33, 30)
PALE_ORANGE = RGBColor(246, 226, 216)
PALE_BLUE = RGBColor(225, 231, 244)
PALE_OCHRE = RGBColor(241, 232, 207)
FONT = 'Noto Sans CJK SC'
EA = 'Noto Sans CJK SC'

W, H = 10.0, 5.625
M = 0.62

SECTIONS = {
    1: '开场', 2: '开场', 3: '课程地图',
    4: '01 重新理解产品经理', 5: '01 重新理解产品经理', 6: '01 重新理解产品经理',
    7: '01 重新理解产品经理', 8: '01 重新理解产品经理',
    9: '02 扩大 Scope', 10: '02 扩大 Scope', 11: '02 扩大 Scope', 12: '02 扩大 Scope',
    13: '03 能力链', 14: '03 能力链', 15: '03 能力链', 16: '03 能力链', 17: '03 能力链',
    18: '04 工作方式', 19: '04 工作方式', 20: '04 工作方式', 21: '04 工作方式',
    22: '04 工作方式', 23: '04 工作方式', 24: '04 工作方式',
    25: '05 转型路径', 26: '05 转型路径', 27: '05 转型路径', 28: '05 转型路径',
    29: '06 组织形态', 30: '06 组织形态', 31: '06 组织形态', 32: '06 组织形态', 33: '06 组织形态',
    34: '07 完整案例', 35: '07 完整案例', 36: '07 完整案例', 37: '07 完整案例',
    38: '07 完整案例', 39: '07 完整案例', 40: '结课', 41: '作业',
}

TITLES = {
    2: '如果资源没有到齐，产品还能继续吗？',
    3: '这节课只回答一件事：产品经理怎样重新成为建设者',
    4: '今天的岗位分工，并不是产品工作的天然形态',
    5: '产品经理最初出现，是为了有人对完整结果负责',
    6: '专业分工让复杂软件成为可能',
    7: '分工越细，产品经理越可能远离产品生产',
    8: 'AI同时降低了整条产品生产链的启动成本',
    9: '产品经理承担完整结果，却只能控制部分过程',
    10: 'FDE要拿回四种控制能力',
    11: 'FDE把“等待交付”改成“亲手完成闭环”',
    12: '扩大Scope不是无限加活，而是缩短关键路径',
    13: 'FDE是一条完整能力链，不是一张技能清单',
    14: '设计与前端：让方案成为真实体验',
    15: '后端与AI：让产品拥有数据、规则和智能',
    16: '部署与运营：跨过Demo与真实使用的边界',
    17: '关键环节要从“能协作”走向“能独立交付”',
    18: 'FDE的工作从业务现场贯穿到持续迭代',
    19: 'FDE接到的不是需求，而是一个业务问题',
    20: '先把“做AI客服”缩成一个可以上线的工作流',
    21: '第一版要薄，但链路必须完整',
    22: 'Demo成立，不代表真实业务成立',
    23: 'AI产品应该逐级增加自动化程度',
    24: '上线之后，才看得见问题真正发生在哪一层',
    25: '产品经理向FDE转型，要跨过四个阶段',
    26: '产出物要从“告诉别人怎么做”变成“现在已经能用”',
    27: '90天：把一个熟悉项目推进到真实使用',
    28: '不必重新想项目，先把已有项目向后做深',
    29: '传统组织依靠岗位之间的接力完成产品',
    30: 'FDE让一个人成为最小完整交付单元',
    31: '专家不会消失，而是从必经节点变成放大器',
    32: '0到1靠完整交付，规模化再引入专业分工',
    33: 'FDE组织必须同时扩大能力、责任和决策权',
    34: '同一个项目，可以停在方案，也可以推进到业务结果',
    35: '把“自动审核”改成“人工预审助手”',
    36: '第一版只需要跑通一个真实审核闭环',
    37: '测试不能只问“判断对不对”',
    38: '不同失败，需要在不同层级修改',
    39: 'FDE亲自完成的，是从现场到结果的整条链',
    40: '产品经理正在从协调者重新成为建设者',
    41: '课后作业：把一个已有项目FDE化',
}


def rgb(value):
    return value if isinstance(value, RGBColor) else RGBColor(*value)


def tx(slide, x, y, w, h, value, size=16, color=INK, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, font=EA, italic=False):
    lines = str(value).split('\n')
    runs = [[(line, size, rgb(color), bold, italic)] for line in lines]
    return text(slide, x, y, w, h, runs, align=align, anchor=valign,
                space_after=2, line_spacing=1.0)


def line(slide, x1, y1, x2, y2, color=LINE, width=1.4, dash=False):
    shp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = RGBColor(*color)
    shp.line.width = Pt(width)
    if dash:
        shp.line.dash_style = 2
    return _flat(shp)


def circle(slide, x, y, d, fill=TEAL, line_c=None, line_w=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor(*fill)
    if line_c:
        shp.line.color.rgb = RGBColor(*line_c); shp.line.width = Pt(line_w or 1)
    else:
        shp.line.fill.background()
    return _flat(shp)


def rounded(slide, x, y, w, h, fill=WHITE, line_c=LINE, r=True):
    return box(slide, x, y, w, h, fill=rgb(fill), line=rgb(line_c), line_w=1.0, round=r)


def icon(slide, name, x, y, size=0.42):
    return picture(slide, str(ASSETS / f'{name}.png'), x, y, size, size, fit='contain')


def quiet_chain(slide, dark=False):
    """Recurring interior register: a quiet two-node rail."""
    c = (111, 139, 132) if dark else LINE
    l = line(slide, 8.62, 0.43, 9.34, 0.43, c, 1.2)
    tag_motif(l, loud=False)
    for x, d in [(8.57, 0.12), (9.27, 0.18)]:
        s = circle(slide, x, 0.43-d/2, d, TEAL if not dark else (105, 185, 170))
        tag_motif(s, loud=False)


def chrome(slide, page, dark=False, title=None):
    box(slide, 0, 0, W, H, fill=DEEP if dark else BG)
    quiet_chain(slide, dark=dark)
    c1 = (174, 190, 185) if dark else MUTE
    tx(slide, M, 0.34, 3.8, 0.24, SECTIONS.get(page, ''), 9.5, c1, True)
    if title:
        tx(slide, M, 0.72, 8.78, 0.70, title, 25.5, WHITE if dark else INK, True)
    line(slide, M, 5.20, 9.38, 5.20, (65, 86, 81) if dark else LINE, 0.8)
    tx(slide, M, 5.29, 6.8, 0.16, 'FDE · 产品经理在AI时代的全栈化生存', 7.5, c1)
    tx(slide, 9.0, 5.27, 0.36, 0.18, f'{page:02d}', 8, c1, True, align=PP_ALIGN.RIGHT)


def note(slide, page, body, sources=None):
    src = '' if not sources else '\n\n资料来源：' + '；'.join(sources)
    speaker_notes(slide, f'第{page}页讲解建议：\n{body}{src}')
    if page in (1, 18, 40):
        design_intent(slide, reason='Scope Chain：用连续链路表达从问题到结果的端到端控制。')


def card(slide, x, y, w, h, title, body='', accent=TEAL, icon_name=None, dark=False, n=None):
    fill = (34, 51, 47) if dark else WHITE
    lc = (72, 94, 88) if dark else LINE
    rounded(slide, x, y, w, h, fill, lc)
    box(slide, x, y, 0.07, h, fill=accent, line=None)
    if n is not None:
        tx(slide, x+0.24, y+0.18, 0.42, 0.35, str(n), 18, accent, True)
        title_x = x+0.72
    elif icon_name:
        icon(slide, icon_name, x+0.23, y+0.20, 0.40)
        title_x = x+0.78
    else:
        title_x = x+0.25
    tx(slide, title_x, y+0.18, w-(title_x-x)-0.2, 0.38, title, 15, WHITE if dark else INK, True)
    if body:
        tx(slide, x+0.25, y+0.67, w-0.46, h-0.83, body, 11.3, (199, 211, 206) if dark else GREY)


def chain(slide, labels, y=2.55, x=0.75, w=8.5, colors=None, subs=None, loud=False):
    colors = colors or [TEAL] * len(labels)
    gap = 0.22
    cw = (w-gap*(len(labels)-1))/len(labels)
    centers = []
    for i, lab in enumerate(labels):
        xx = x+i*(cw+gap)
        rounded(slide, xx, y, cw, 0.86, WHITE, colors[i])
        tx(slide, xx+0.08, y+0.18, cw-0.16, 0.24, lab, 12.5, colors[i], True, align=PP_ALIGN.CENTER)
        if subs:
            tx(slide, xx+0.08, y+0.48, cw-0.16, 0.20, subs[i], 8.5, MUTE, align=PP_ALIGN.CENTER)
        centers.append((xx+cw/2, y+0.43))
        if i < len(labels)-1:
            a = arrow(slide, xx+cw+0.03, y+0.31, gap-0.06, 0.24, color=LINE)
            tag_motif(a, loud=False)
    rail = line(slide, centers[0][0], y+1.13, centers[-1][0], y+1.13, LINE, 1.2)
    tag_motif(rail, loud=loud)
    for i, (cx, _) in enumerate(centers):
        d = 0.10 + i*0.014 if loud else 0.10
        s = circle(slide, cx-d/2, y+1.13-d/2, d, colors[i])
        tag_motif(s, loud=loud)
    return centers


def two_col(slide, left_title, left_items, right_title, right_items, y=1.62, left_accent=ORANGE, right_accent=TEAL):
    for x, title, items, ac in [(0.65,left_title,left_items,left_accent),(5.12,right_title,right_items,right_accent)]:
        rounded(slide, x, y, 4.22, 3.16, WHITE, ac)
        tx(slide, x+0.28, y+0.23, 3.68, 0.36, title, 17, ac, True)
        for i, item in enumerate(items):
            yy = y+0.83+i*0.63
            circle(slide, x+0.29, yy+0.03, 0.13, ac)
            tx(slide, x+0.56, yy, 3.33, 0.43, item, 12.2, GREY)


def section_divider(prs, page, number, eyebrow, title, body):
    s = add_slide(prs); chrome(s, page, dark=True)
    tx(s, 0.70, 1.05, 1.2, 0.75, number, 44, TEAL, True)
    tx(s, 0.76, 1.92, 3.2, 0.28, eyebrow, 11, (157, 179, 173), True)
    tx(s, 0.76, 2.32, 8.1, 1.02, title, 29, WHITE, True)
    tx(s, 0.78, 3.57, 7.4, 0.72, body, 14, (190, 204, 199))
    # loud chain marker: one of three hero appearances
    rail = line(s, 6.55, 4.45, 9.10, 4.45, (74, 104, 96), 2)
    tag_motif(rail, loud=False)
    for i, d in enumerate([0.15,0.22,0.29,0.36]):
        shp = circle(s, 6.6+i*0.77-d/2, 4.45-d/2, d, [TEAL, BLUE, OCHRE, ORANGE][i])
        tag_motif(shp, loud=False)
    note(s, page, body)
    return s


def build_cover(prs):
    s = add_slide(prs)
    box(s, 0, 0, 4.05, H, fill=DEEP)
    box(s, 4.05, 0, 5.95, H, fill=BG)
    box(s, 0, 0, 0.12, H, fill=TEAL)
    tx(s, 0.72, 0.75, 2.4, 0.28, 'AI 产品经理专题课', 10.5, (168,190,183), True)
    tx(s, 0.72, 1.35, 3.0, 1.55, 'FDE', 48, WHITE, True)
    tx(s, 0.72, 3.05, 2.95, 0.95, '产品经理在AI时代的\n全栈化生存', 19.5, WHITE, True)
    tx(s, 0.72, 4.62, 2.8, 0.35, '从定义产品，到亲手把产品做出来', 11, (174,194,188))
    labels = ['业务', '产品', '设计', '代码', '上线', '运营']
    x0, y = 4.72, 2.73
    line(s, x0, y, 9.28, y, LINE, 2)
    for i, lab in enumerate(labels):
        d = 0.24 + i*0.055
        x = x0+i*0.91
        shp = circle(s, x-d/2, y-d/2, d, [MUTE,TEAL,TEAL,BLUE,OCHRE,ORANGE][i])
        tag_motif(shp, loud=True)
        tx(s, x-0.36, y+0.48, 0.72, 0.25, lab, 9, GREY, i>0, align=PP_ALIGN.CENTER)
    tx(s, 4.63, 1.23, 4.7, 0.70, '一条越来越完整的\n产品生产链', 24, INK, True)
    tx(s, 8.20, 4.92, 1.15, 0.22, '正式授课版', 8.5, MUTE, True, align=PP_ALIGN.RIGHT)
    note(s, 1, '开场先说明：这里的FDE不是狭义岗位介绍，而是一种产品经理在AI时代重新拿回端到端交付能力的工作方式。')


def build_slide_2(prs):
    s = add_slide(prs); chrome(s, 2, title=TITLES[2])
    qs = [('如果没有设计师','你能把体验做出来吗？','palette'),('如果没有前端','你能让别人真正使用吗？','code-2'),
          ('如果没有算法工程师','你能完成第一版智能能力吗？','brain-circuit'),('如果没有运营','你能让产品产生真实反馈吗？','users')]
    for i,(a,b,ic) in enumerate(qs):
        x=0.67+(i%2)*4.42; y=1.52+(i//2)*1.60
        icon_card(s,x,y,4.10,1.50,str(ASSETS/f'{ic}.png'),str(1+i).zfill(2)+'  '+a,'',fill=WHITE,line=LINE,accent=[TEAL,BLUE,OCHRE,ORANGE][i],ink=INK,body_c=GREY)
        tx(s,x+0.26,y+1.10,3.58,0.16,b,11.3,GREY)
    tx(s, 0.72, 4.77, 8.55, 0.25, '资源不到位时仍能推进，才是FDE真正改变产品经理的地方。', 12, TEAL, True, align=PP_ALIGN.CENTER)
    note(s,2,'让学员先代入真实工作：当下一个岗位没有空档期时，项目会不会停。不要急着给结论，把问题留到后面。')


def build_slide_3(prs):
    s=add_slide(prs); chrome(s,3,title=TITLES[3])
    labels=['为什么变','拿回什么','怎样工作','如何转型','组织如何变','案例怎么做']
    subs=['历史与冲击','四种控制','端到端闭环','90天路径','最小单元','审核助手']
    chain(s,labels,y=2.12,x=0.68,w=8.70,colors=[MUTE,TEAL,BLUE,TEAL,OCHRE,ORANGE],subs=subs)
    tx(s,1.10,3.72,7.80,0.52,'不是把每个岗位都学一遍，而是让一个业务结果不再卡在岗位边界上。',18,INK,True,align=PP_ALIGN.CENTER)
    note(s,3,'给出课程地图。强调主线不是技能罗列，而是从产品经理的历史分工一路走到一个真实项目闭环。')


def build_slide_5(prs):
    s=add_slide(prs); chrome(s,5,title=TITLES[5])
    timeline(s,0.85,2.05,8.25,[('1931','品牌管理'),('传统制造','跨职能协调'),('软件时代','产品经理'),('AI时代','FDE式建设者')],highlight=0,accent=TEAL,ink=INK,axis_c=LINE,h=1.35)
    card(s,0.88,3.57,8.18,1.12,'核心不是写需求，而是对一个完整结果持续负责','岗位名称会变，但“有人对结果负责”这件事一直没有变。',TEAL,'history')
    source_note(s,['P&G Innovation Timeline：1931年 Neil McElroy 创建品牌管理'],as_of='2026-10')
    note(s,5,'从P&G的品牌管理讲起。原点不是“写文档的人”，而是跨研发、生产、销售对一个品牌的完整结果负责的人。',['https://us.pg.com/innovation/'])


def build_slide_6(prs):
    s=add_slide(prs); chrome(s,6,title=TITLES[6])
    roles=[('产品','定义问题',TEAL),('设计','塑造体验',OCHRE),('前端','实现交互',BLUE),('后端','连接系统',BLUE),('算法','构建智能',ORANGE),('运维','保障稳定',MUTE)]
    for i,(r,b,c) in enumerate(roles):
        x=0.66+i*1.47
        rounded(s,x,1.86,1.25,1.74,WHITE,c)
        circle(s,x+0.46,2.12,0.32,c)
        tx(s,x+0.12,2.57,1.01,0.28,r,14,c,True,align=PP_ALIGN.CENTER)
        tx(s,x+0.12,3.00,1.01,0.26,b,9.5,GREY,align=PP_ALIGN.CENTER)
        if i<5: arrow(s,x+1.29,2.55,0.12,0.22,LINE)
    tx(s,0.98,4.17,8.04,0.50,'分工解决了复杂性，也让每个人只看见产品的一小段。',18,INK,True,align=PP_ALIGN.CENTER)
    note(s,6,'肯定专业分工的价值：软件复杂度上升后，岗位专业化是必然。后面讨论的不是取消专家，而是改变0到1阶段的工作单元。')


def build_slide_7(prs):
    s=add_slide(prs); chrome(s,7,title=TITLES[7])
    labels=['用户问题','需求文档','视觉稿','排期','开发','测试','上线','反馈']
    cols=[TEAL,MUTE,MUTE,MUTE,BLUE,MUTE,OCHRE,ORANGE]
    chain(s,labels,y=2.18,x=0.50,w=9.00,colors=cols,subs=['现场','描述','表达','等待','实现','交接','发布','结果'])
    for x in [2.70,4.95,7.20]:
        line(s,x,1.86,x,3.82,ORANGE,1.5,True)
        tx(s,x-0.37,3.88,0.74,0.25,'交接点',8.5,ORANGE_DARK,True,align=PP_ALIGN.CENTER)
    tx(s,1.17,4.47,7.66,0.35,'岗位越多，局部更专业；交接越多，完整反馈越慢。',16,INK,True,align=PP_ALIGN.CENTER)
    note(s,7,'讲分工的副作用：产品经理可能越来越擅长协调，却越来越少直接生产产品。每个交接点都会增加等待、损耗和误解。')


def build_slide_8(prs):
    s=add_slide(prs); chrome(s,8,title=TITLES[8])
    areas=[('理解研究','总结、分类、检索','search-check',TEAL),('设计表达','原型、文案、视觉','palette',OCHRE),
           ('开发集成','代码、接口、自动化','code-2',BLUE),('运营分析','数据、内容、反馈','gauge',ORANGE)]
    for i,(a,b,ic,c) in enumerate(areas):
        x=0.68+i*2.25
        icon_card(s,x,1.76,2.02,2.28,str(ASSETS/f'{ic}.png'),str(i+1).zfill(2)+'  '+a,b,fill=WHITE,line=LINE,accent=c,ink=INK,body_c=GREY)
    tx(s,0.82,4.35,8.36,0.53,'AI没有消灭岗位，但把“亲手完成第一版”的门槛整体向下拉。',18,TEAL,True,align=PP_ALIGN.CENTER)
    note(s,8,'把AI理解成全链路启动成本的下降：研究、设计、编码、部署、运营都更容易做出第一版。重点不是替代谁，而是谁可以更快跨过岗位边界。')


def build_slide_10(prs):
    s=add_slide(prs); chrome(s,10,title=TITLES[10])
    items=[('问题控制','定义真实业务问题','search-check',TEAL),('方案控制','决定产品形态','palette',OCHRE),
           ('实现控制','让关键链路跑通','code-2',BLUE),('结果控制','上线、观察、改进','gauge',ORANGE)]
    for i,(a,b,ic,c) in enumerate(items):
        x=0.72+i*2.20
        icon_card(s,x,1.78,1.98,2.36,str(ASSETS/f'{ic}.png'),a,b,fill=WHITE,line=c,accent=c,ink=INK,body_c=GREY)
        if i<3: arrow(s,x+2.01,2.74,0.13,0.23,LINE)
    tx(s,1.20,4.51,7.60,0.32,'控制不是一个人包办全部，而是关键环节不再完全依赖别人启动。',14,INK,True,align=PP_ALIGN.CENTER)
    note(s,10,'四种控制能力是本课的核心框架。强调“控制”是可以推进和验证，不是拒绝协作，也不是替代所有专家。')


def build_slide_11(prs):
    s=add_slide(prs); chrome(s,11,title=TITLES[11])
    two_col(s,'传统产品经理：岗位接力',['写清需求，再等待设计','评审方案，再等待开发','跟进上线，再等待反馈'],
            'FDE式产品经理：亲手闭环',['先做可用原型','打通最小系统','进入真实使用再迭代'],left_accent=ORANGE,right_accent=TEAL)
    note(s,11,'用同一个项目对比两种推进方式。传统方式的主动作是交付给下一个岗位；FDE的主动作是让产品往可用状态前进一步。')


def build_slide_12(prs):
    s=add_slide(prs); chrome(s,12,title=TITLES[12])
    tx(s,0.76,1.68,2.55,0.32,'扩大范围 ≠ 承包所有工作',13.5,ORANGE_DARK,True)
    labels=['必须亲自掌握','可以AI辅助','需要专家放大']
    widths=[3.35,2.45,1.65]; cols=[TEAL,BLUE,OCHRE]
    x=0.76
    for i,(lab,ww,c) in enumerate(zip(labels,widths,cols)):
        rounded(s,x,2.22,ww,1.23,WHITE,c)
        tx(s,x+0.16,2.45,ww-0.32,0.32,lab,14,c,True,align=PP_ALIGN.CENTER)
        sub=['业务闭环与关键路径','研究、设计、代码初稿','安全、性能、规模化'][i]
        tx(s,x+0.16,2.91,ww-0.32,0.25,sub,9.5,GREY,align=PP_ALIGN.CENTER)
        x+=ww+0.18
    line(s,0.96,4.16,9.04,4.16,LINE,1.5)
    for i,(lab,c) in enumerate([('短：自己完成',TEAL),('中：AI协作',BLUE),('长：专家介入',OCHRE)]):
        d=0.15+i*0.08; xx=2.30+i*2.72
        circle(s,xx-d/2,4.16-d/2,d,c); tx(s,xx+0.20,4.01,1.60,0.30,lab,10.5,OCHRE_DARK if i==2 else c,True)
    note(s,12,'给Scope划边界。0到1阶段关键链路要能自己做；重复性工作让AI辅助；高风险、高性能和规模化问题仍然需要专家。')


def build_slide_13(prs):
    s=add_slide(prs); chrome(s,13,title=TITLES[13])
    labels=['业务理解','产品定义','交互设计','前端实现','后端与AI','部署上线','运营迭代']
    cols=[TEAL,TEAL,OCHRE,BLUE,BLUE,ORANGE,ORANGE]
    chain(s,labels,y=2.05,x=0.52,w=8.96,colors=cols,subs=['现场','边界','体验','可用','能力','真实','反馈'],loud=False)
    tx(s,0.92,3.86,8.16,0.30,'每一项能力的价值，都来自它对下一环节的直接连接。',15,INK,True,align=PP_ALIGN.CENTER)
    note(s,13,'把技能清单改成能力链。会做原型不重要，能否让原型进入前端；会调模型不重要，能否接进真实工作流。')


def build_slide_14(prs):
    s=add_slide(prs); chrome(s,14,title=TITLES[14])
    two_col(s,'设计：把抽象需求变成可体验',['信息架构与关键页面','状态、异常与人工接管','让用户知道系统正在做什么'],
            '前端：把体验变成可操作',['数据输入与结果呈现','权限、表单与工作台','快速收集真实使用反馈'],left_accent=OCHRE,right_accent=BLUE)
    note(s,14,'设计和前端不是为了“做得漂亮”，而是把讨论对象从文档变成可以使用的体验。强调状态、异常和接管入口。')


def build_slide_15(prs):
    s=add_slide(prs); chrome(s,15,title=TITLES[15])
    labels=['业务数据','后端服务','模型能力','规则约束','结果输出']
    subs=['来源与权限','接口与状态','理解与生成','边界与兜底','结构化可用']
    chain(s,labels,y=2.05,x=0.78,w=8.44,colors=[TEAL,BLUE,ORANGE,OCHRE,TEAL],subs=subs)
    card(s,1.30,3.64,7.40,1.05,'后端决定产品能否运行，AI决定产品能否处理不确定性','两者必须共同进入同一条业务链路，而不是各自做一个Demo。',BLUE,'brain-circuit')
    note(s,15,'讲清后端与AI的分工：后端负责数据、状态、权限、可追踪；AI负责理解、判断、生成。真正的产品需要把两者接在一起。')


def build_slide_16(prs):
    s=add_slide(prs); chrome(s,16,title=TITLES[16])
    stages=[('本地能跑','我能看到',MUTE),('测试环境','同事能用',BLUE),('受控上线','少量用户',OCHRE),('真实运营','持续改进',ORANGE)]
    for i,(a,b,c) in enumerate(stages):
        x=0.82+i*2.12; y=1.75+i*0.36
        rounded(s,x,y,1.82,1.05,WHITE,c)
        tx(s,x+0.14,y+0.20,1.54,0.27,a,13.5,c,True,align=PP_ALIGN.CENTER)
        tx(s,x+0.14,y+0.61,1.54,0.22,b,9.5,GREY,align=PP_ALIGN.CENTER)
        if i<3: arrow(s,x+1.84,y+0.42,0.24,0.23,LINE)
    tx(s,0.96,4.25,8.08,0.38,'产品只有进入别人的工作流，才开始接受真实世界的约束。',17,INK,True,align=PP_ALIGN.CENTER)
    note(s,16,'区分“能演示”和“能使用”。部署不是技术收尾，而是把产品放进真实用户、权限、数据和异常里。运营则负责让反馈回到产品。')


def build_slide_17(prs):
    s=add_slide(prs); chrome(s,17,title=TITLES[17])
    levels=[('能理解','知道原理与边界',0.34,MUTE,'起步'),('能协作','能拆任务、能验收',0.62,BLUE,'协作'),('能独立交付','没有资源也能跑通第一版',0.92,TEAL,'独立')]
    for i,(a,b,f,c,v) in enumerate(levels):
        y=1.73+i*0.98
        tx(s,0.83,y,1.48,0.28,a,14,c,True)
        tx(s,2.14,y,2.43,0.28,b,10.5,GREY)
        meter_bar(s,4.75,y-0.05,4.13,f,label='',value=v,accent=c,track=LIGHT,h=0.24,value_pos='right',ink=INK)
    tx(s,1.02,4.73,7.96,0.26,'不是所有能力都要精通，但关键闭环必须至少有一条能独立跑通。',13.5,INK,True,align=PP_ALIGN.CENTER)
    note(s,17,'能力成熟度分三档。对大多数产品经理而言，先在一个熟悉业务中把关键链路做到独立交付，比平均学习所有技术更有效。')


def build_slide_18(prs):
    s=add_slide(prs); chrome(s,18,dark=True,title=TITLES[18])
    labels=['进入现场','定义结果','做出原型','接通系统','受控上线','观察使用','持续迭代']
    cols=[TEAL,TEAL,OCHRE,BLUE,BLUE,ORANGE,ORANGE]
    chain(s,labels,y=2.10,x=0.52,w=8.96,colors=cols,subs=['看真实工作','定业务指标','先可体验','数据+AI','有人兜底','找真实失败','修改整条链'],loud=True)
    tx(s,0.82,3.87,8.36,0.46,'FDE的边界不是“需求完成”，而是一个可衡量的业务结果已经发生。',17,WHITE,True,align=PP_ALIGN.CENTER)
    source_note(s,['OpenAI Forward Deployed Engineering：embed → prototype → validate → go live → scale'],as_of='2026-10',ink=RGBColor(166,188,181))
    note(s,18,'这是本课最重要的一张流程图。FDE从业务现场开始，定义结果、做原型、接系统、上线，再根据真实使用迭代。',['https://openai.com/business/forward-deployed-engineering/'])


def build_slide_19(prs):
    s=add_slide(prs); chrome(s,19,title=TITLES[19])
    two_col(s,'收到的表面需求',['“做一个AI客服”','“加一个知识库”','“自动回复用户”'],
            '真正要追的问题',['哪类问题占用最多人力？','什么结果才算解决？','哪里必须人工负责？'],left_accent=MUTE,right_accent=TEAL)
    tx(s,1.12,4.77,7.76,0.24,'先把功能名拆掉，才看得见业务结果。',13,TEAL,True,align=PP_ALIGN.CENTER)
    note(s,19,'从需求转向业务问题。FDE不会直接接受功能名，而是去现场观察：谁在什么情况下做什么，哪里耗时，哪里出错，怎样算改善。')


def build_slide_20(prs):
    s=add_slide(prs); chrome(s,20,title=TITLES[20])
    labels=['用户提问','识别问题','检索知识','生成草稿','人工确认','发送答复','记录结果']
    subs=['入口','分类','来源','可编辑','高风险兜底','执行','回流']
    chain(s,labels,y=2.08,x=0.47,w=9.06,colors=[TEAL,BLUE,BLUE,ORANGE,OCHRE,TEAL,MUTE],subs=subs)
    card(s,1.34,3.62,7.32,1.05,'把一个大项目压缩成一条可验证工作流','先服务一类高频问题，再逐步扩展范围。',TEAL,'workflow')
    note(s,20,'示范如何把“AI客服”改写成具体工作流。第一版不是覆盖所有咨询，而是选一类高频、低风险、资料稳定的问题。')


def build_slide_21(prs):
    s=add_slide(prs); chrome(s,21,title=TITLES[21])
    labels=['真实输入','可用处理','人工确认','真实输出','结果记录']
    subs=['不是样例','不是假按钮','责任明确','进入业务','能复盘']
    chain(s,labels,y=2.02,x=0.82,w=8.36,colors=[TEAL,BLUE,OCHRE,ORANGE,TEAL],subs=subs)
    tx(s,1.03,3.72,7.94,0.36,'可以很薄：只做一类问题、一个入口、一个团队。',15,GREY,True,align=PP_ALIGN.CENTER)
    tx(s,1.03,4.18,7.94,0.43,'不能断：输入、处理、确认、输出、记录必须首尾相接。',17,TEAL,True,align=PP_ALIGN.CENTER)
    note(s,21,'解释“薄但完整”。功能可以少，覆盖范围可以窄，但不能用假数据、假按钮或口头人工替代关键链路。')


def build_slide_22(prs):
    s=add_slide(prs); chrome(s,22,title=TITLES[22])
    two_col(s,'Demo世界',['输入被挑选过','知识资料很干净','只有成功路径','没有权限和责任'],
            '真实业务',['输入混乱且不完整','资料冲突或过期','异常高频出现','每个动作都有责任人'],left_accent=BLUE,right_accent=ORANGE)
    note(s,22,'Demo证明技术可能，真实业务检验产品成立。重点讲脏数据、边界情况、权限、人工责任和失败恢复。')


def build_slide_23(prs):
    s=add_slide(prs); chrome(s,23,title=TITLES[23])
    tiers=['L4 自动闭环 · 稳定场景自动处理','L3 受控执行 · 低风险自动，高风险拦截','L2 草稿 · AI生成，人工确认','L1 建议 · AI只给建议']
    tier_stack(s,1.02,1.58,7.96,3.18,tiers,mode='pyramid',direction='up',accents=[TEAL,OCHRE,BLUE,MUTE],ink=INK,labels='inside')
    tx(s,6.95,1.76,1.52,0.28,'自动化程度 ↑',10,TEAL,True,align=PP_ALIGN.RIGHT)
    tx(s,6.95,4.30,1.52,0.28,'风险容忍度 ↓',10,WHITE,True,align=PP_ALIGN.RIGHT)
    note(s,23,'不要一步跳到全自动。AI产品上线应逐级增加自动化：建议、草稿、受控执行、自动闭环。每一级都要有清晰的责任和回退。')


def build_slide_24(prs):
    s=add_slide(prs); chrome(s,24,title=TITLES[24])
    layers=[('业务层','目标是否正确','briefcase-business',TEAL),('产品层','流程与边界是否合理','workflow',OCHRE),
            ('系统层','数据、接口、状态是否稳定','database',BLUE),('模型层','理解、检索、生成是否可靠','brain-circuit',ORANGE),('运营层','用户是否愿意持续使用','users',MUTE)]
    for i,(a,b,ic,c) in enumerate(layers):
        y=1.55+i*0.69
        rounded(s,0.78,y,8.44,0.55,WHITE,c)
        icon(s,ic,0.98,y+0.09,0.34)
        tx(s,1.46,y+0.12,1.10,0.24,a,12.5,c,True)
        tx(s,2.72,y+0.12,5.97,0.24,b,11,GREY)
    note(s,24,'上线后不要把所有问题归因于模型。按业务、产品、系统、模型、运营五层定位根因，才能改对地方。')


def build_slide_26(prs):
    s=add_slide(prs); chrome(s,26,title=TITLES[26])
    pairs=[('需求文档','可操作原型'),('评审结论','可调用接口'),('项目计划','已部署版本'),('复盘汇报','真实使用数据')]
    for i,(a,b) in enumerate(pairs):
        y=1.62+i*0.82
        rounded(s,0.85,y,3.28,0.60,LIGHT,LINE)
        tx(s,1.05,y+0.15,2.88,0.25,a,12.5,GREY,True,align=PP_ALIGN.CENTER)
        arrow(s,4.35,y+0.18,1.04,0.24,TEAL)
        rounded(s,5.62,y,3.48,0.60,WHITE,TEAL)
        tx(s,5.82,y+0.15,3.08,0.25,b,12.5,TEAL,True,align=PP_ALIGN.CENTER)
    note(s,26,'展示产出物的变化。文档仍然有用，但它不再是终点；每个阶段都应该多一个可以运行、可以使用、可以观察的对象。')


def build_slide_27(prs):
    s=add_slide(prs); chrome(s,27,title=TITLES[27])
    stages=[('第1–30天','选一个熟悉流程','完成可操作原型',TEAL),('第31–60天','接入真实数据与AI','打通最小后端',BLUE),('第61–90天','受控上线','收集反馈并迭代',ORANGE)]
    for i,(time,a,b,c) in enumerate(stages):
        x=0.72+i*3.02
        rounded(s,x,1.66,2.76,2.78,WHITE,c)
        tx(s,x+0.22,1.90,2.32,0.34,time,16,c,True)
        line(s,x+0.22,2.42,x+2.54,2.42,LINE,1)
        tx(s,x+0.22,2.70,2.32,0.45,a,13,INK,True)
        tx(s,x+0.22,3.38,2.32,0.50,b,11,GREY)
        if i<2: arrow(s,x+2.79,2.93,0.18,0.24,LINE)
    note(s,27,'给出可执行的90天路线。第一月先做出可操作体验；第二月接数据、后端和AI；第三月进入受控真实使用。')


def build_slide_28(prs):
    s=add_slide(prs); chrome(s,28,title=TITLES[28])
    labels=['已有方案','可操作原型','真实接口','受控上线','使用数据','下一轮改进']
    chain(s,labels,y=2.08,x=0.60,w=8.80,colors=[MUTE,OCHRE,BLUE,TEAL,ORANGE,TEAL],subs=['你已经有','继续做深','接系统','给用户','找问题','再迭代'])
    tx(s,0.98,3.78,8.04,0.36,'最好的FDE练习，不是再写一个新项目，而是把旧项目往后推进两步。',16,INK,True,align=PP_ALIGN.CENTER)
    note(s,28,'让学员从已有客服、RAG、审核或数据分析项目出发。先问：它现在停在哪里？下一步怎样进入真实输入、真实用户和结果记录。')


def build_slide_30(prs):
    s=add_slide(prs); chrome(s,30,title=TITLES[30])
    circle(s,3.72,1.62,2.56,LIGHT,TEAL,1.8)
    circle(s,4.25,2.15,1.50,WHITE,TEAL,1.8)
    tx(s,4.30,2.54,1.40,0.30,'FDE',22,TEAL,True,align=PP_ALIGN.CENTER)
    labels=[('业务','看现场',0.84,2.05),('产品','定边界',1.47,3.56),('设计','做体验',3.05,4.44),('技术','接系统',6.12,4.42),('上线','进业务',7.56,3.52),('运营','拿反馈',8.10,2.00)]
    for i,(a,b,x,y) in enumerate(labels):
        rounded(s,x,y,1.18,0.72,WHITE,[TEAL,TEAL,OCHRE,BLUE,ORANGE,ORANGE][i])
        tx(s,x+0.08,y+0.11,1.02,0.22,a,11.5,[TEAL,TEAL,OCHRE,BLUE,ORANGE,ORANGE][i],True,align=PP_ALIGN.CENTER)
        tx(s,x+0.08,y+0.41,1.02,0.16,b,8.2,GREY,align=PP_ALIGN.CENTER)
        line(s,x+(0 if x>5 else 1.18),y+0.36,5.0,2.90,LINE,1)
    note(s,30,'FDE不是“什么都做的人”，而是能独立形成一个最小完整交付单元的人：看懂业务、做出体验、接通系统、上线和运营。')


def build_slide_31(prs):
    s=add_slide(prs); chrome(s,31,title=TITLES[31])
    org_tree(s,1.05,1.55,7.90,3.20,('FDE 最小交付单元',[('安全专家',[]),('模型专家',[]),('平台工程',[]),('行业专家',[])]),accent=TEAL,node_h=0.58,gap_y=0.60,label_size=12,font=EA,line_c=LINE)
    tx(s,1.02,4.55,7.96,0.34,'专家在关键节点放大质量与规模，而不是每一步都成为等待对象。',14,INK,True,align=PP_ALIGN.CENTER)
    note(s,31,'说明专家不会消失。FDE负责快速形成完整闭环；专家在安全、模型、平台、行业深度等关键点介入，把已经成立的方案做得更稳、更大。')


def build_slide_32(prs):
    s=add_slide(prs); chrome(s,32,title=TITLES[32])
    phases=[('0→1','1个FDE','速度优先','完整闭环',TEAL),('1→10','FDE + 专家','质量优先','可复制方案',BLUE),('10→100','专业团队','规模优先','平台化能力',ORANGE)]
    for i,(a,b,c,d,col) in enumerate(phases):
        x=0.72+i*3.02
        rounded(s,x,1.62,2.76,2.92,WHITE,col)
        tx(s,x+0.20,1.88,2.36,0.48,a,28,col,True,align=PP_ALIGN.CENTER)
        tx(s,x+0.20,2.60,2.36,0.32,b,14,INK,True,align=PP_ALIGN.CENTER)
        tx(s,x+0.20,3.21,2.36,0.26,c,10.5,GREY,align=PP_ALIGN.CENTER)
        tx(s,x+0.20,3.70,2.36,0.30,d,12,col,True,align=PP_ALIGN.CENTER)
    note(s,32,'把FDE放回组织生命周期。0到1强调完整交付；1到10加入专家提升质量和复制；10到100进入平台化与专业分工。')


def build_slide_33(prs):
    s=add_slide(prs); chrome(s,33,title=TITLES[33])
    items=[('能力范围','能亲手推进多少环节',TEAL),('结果责任','对什么业务结果负责',ORANGE),('决策权限','能否改变方案和流程',BLUE)]
    for i,(a,b,c) in enumerate(items):
        x=0.78+i*3.02
        rounded(s,x,1.75,2.74,1.62,WHITE,c)
        circle(s,x+1.12,2.00,0.50,c)
        tx(s,x+0.20,2.62,2.34,0.30,a,15,c,True,align=PP_ALIGN.CENTER)
        tx(s,x+0.20,3.02,2.34,0.23,b,9.5,GREY,align=PP_ALIGN.CENTER)
    tx(s,1.18,3.82,7.64,0.48,'能力扩大 + 责任扩大 + 权限扩大',23,INK,True,align=PP_ALIGN.CENTER)
    tx(s,1.18,4.39,7.64,0.30,'只有三者一起变化，FDE才不是“多干活的产品经理”。',12.5,ORANGE_DARK,True,align=PP_ALIGN.CENTER)
    note(s,33,'这是组织边界。只扩大能力和责任、不给决策权，会把FDE变成救火队；只给权、不补能力，也无法形成完整交付。')


def build_slide_34(prs):
    s=add_slide(prs); chrome(s,34,dark=True,title=TITLES[34])
    two_col(s,'方案型完成',['需求分析','功能结构','原型与PRD'],
            'FDE型完成',['真实审核闭环','受控上线','使用数据与迭代'],y=1.53,left_accent=MUTE,right_accent=TEAL)
    tx(s,1.03,4.77,7.94,0.24,'差别不在PPT写得多完整，而在产品向真实结果推进了多远。',13,WHITE,True,align=PP_ALIGN.CENTER)
    note(s,34,'进入完整案例。先用同一个内容审核项目对比两种完成深度，告诉学员后面看的不是功能更多，而是闭环更完整。')


def build_slide_35(prs):
    s=add_slide(prs); chrome(s,35,title=TITLES[35])
    tx(s,0.78,1.58,3.52,0.30,'原命题',11,MUTE,True)
    rounded(s,0.78,1.96,3.52,1.18,LIGHT,LINE)
    tx(s,1.03,2.29,3.02,0.48,'系统自动判断并拦截\n所有违规营销内容',16,GREY,True,align=PP_ALIGN.CENTER)
    arrow(s,4.47,2.38,0.76,0.28,TEAL)
    tx(s,5.48,1.58,3.74,0.30,'第一版产品命题',11,TEAL,True)
    rounded(s,5.48,1.96,3.74,1.18,WHITE,TEAL)
    tx(s,5.73,2.29,3.24,0.48,'AI先做预审与说明\n人工做最终决定',16,TEAL,True,align=PP_ALIGN.CENTER)
    card(s,1.08,3.60,7.84,1.05,'为什么这样改','降低错误拦截风险，同时保留AI的效率价值和人工责任边界。',ORANGE,'shield-check')
    note(s,35,'做关键产品取舍：第一版不追求全自动审核，而是定位为人工预审助手。AI负责发现风险、定位条款、给理由；审核员负责最终决定。')


def build_slide_36(prs):
    s=add_slide(prs); chrome(s,36,title=TITLES[36])
    labels=['提交内容','AI预审','风险定位','人工确认','结果回写','问题样本回流']
    subs=['文案/图片','规则+模型','说明可编辑','通过/驳回','业务系统','下一轮测试']
    chain(s,labels,y=1.78,x=0.55,w=8.90,colors=[TEAL,BLUE,ORANGE,OCHRE,TEAL,MUTE],subs=subs)
    # architecture rail
    y=3.54
    for i,(a,c) in enumerate([('审核工作台',OCHRE),('审核服务',BLUE),('模型与规则',ORANGE),('内容系统',TEAL),('结果库',MUTE)]):
        x=0.84+i*1.72
        rounded(s,x,y,1.46,0.67,WHITE,c)
        tx(s,x+0.08,y+0.21,1.30,0.24,a,10.5,c,True,align=PP_ALIGN.CENTER)
        if i<4: arrow(s,x+1.49,y+0.20,0.17,0.24,LINE)
    tx(s,0.88,4.45,8.20,0.28,'闭环负责业务流，架构负责让每一步可以运行、追踪和回退。',12,INK,True,align=PP_ALIGN.CENTER)
    note(s,36,'上半部分讲业务闭环，下半部分讲最小系统架构。第一版只要打通提交、预审、说明、人工确认、回写和样本回流。')


def build_slide_37(prs):
    s=add_slide(prs); chrome(s,37,title=TITLES[37])
    layers=[('01','判断质量','风险是否识别正确',TEAL),('02','解释质量','理由是否能被审核员理解',BLUE),
            ('03','流程效率','是否缩短审核时长',OCHRE),('04','系统可靠','异常能否恢复与追踪',ORANGE),('05','业务接受','审核员是否愿意持续使用',MUTE)]
    for i,(n,a,b,c) in enumerate(layers):
        x=0.61+i*1.78
        rounded(s,x,1.72,1.57,2.70,WHITE,c)
        tx(s,x+0.16,1.97,0.46,0.30,n,15,c,True)
        tx(s,x+0.16,2.54,1.25,0.46,a,14,INK,True)
        tx(s,x+0.16,3.24,1.25,0.64,b,10,GREY)
    note(s,37,'测试分五层：判断、解释、流程效率、系统可靠、业务接受。这样才能避免只拿准确率判断整个产品是否成立。')


def build_slide_38(prs):
    s=add_slide(prs); chrome(s,38,title=TITLES[38])
    rows=[('漏判高风险内容','模型层','补充样本与规则',ORANGE),('判断正确但没人信','产品层','展示来源与理由',OCHRE),
          ('等待时间太长','系统层','异步队列与状态反馈',BLUE),('审核员绕开工具','流程层','嵌入原工作台',TEAL)]
    tx(s,0.79,1.55,2.86,0.25,'真实失败',10,MUTE,True)
    tx(s,3.89,1.55,1.40,0.25,'问题层级',10,MUTE,True)
    tx(s,5.55,1.55,3.50,0.25,'修改动作',10,MUTE,True)
    for i,(a,b,c,col) in enumerate(rows):
        y=1.93+i*0.72
        rounded(s,0.72,y,8.55,0.57,WHITE,LINE)
        box(s,0.72,y,0.07,0.57,fill=col,line=None)
        tx(s,0.97,y+0.14,2.62,0.24,a,11.5,INK,True)
        tx(s,3.89,y+0.14,1.25,0.24,b,11,col,True)
        tx(s,5.55,y+0.14,3.42,0.24,c,11,GREY)
    note(s,38,'展示跨层修改。一个失败可能不是模型问题：没人信要改解释和来源，速度慢要改系统，没人用要改嵌入方式。FDE要能跨层定位并推进修改。')


def build_slide_39(prs):
    s=add_slide(prs); chrome(s,39,title=TITLES[39])
    labels=['进审核现场','重定义产品','画工作台','接模型规则','部署试点','看使用记录','跨层修改']
    subs=['看真实动作','预审助手','可操作','可运行','小范围','找失败','持续迭代']
    chain(s,labels,y=2.04,x=0.48,w=9.04,colors=[TEAL,TEAL,OCHRE,BLUE,ORANGE,MUTE,TEAL],subs=subs)
    tx(s,0.92,3.83,8.16,0.62,'FDE的专业性，不是“每一项都比专家强”，\n而是能让所有关键环节围绕同一个业务结果闭环。',17,INK,True,align=PP_ALIGN.CENTER)
    note(s,39,'回看案例里FDE亲自做的事情：现场观察、产品取舍、工作台、模型与规则接入、试点、使用分析和跨层修改。专业性来自完整闭环。')


def build_slide_40(prs):
    s=add_slide(prs); chrome(s,40,dark=True)
    tx(s,0.72,1.35,8.5,0.72,TITLES[40],30,WHITE,True)
    tx(s,0.73,2.30,2.45,0.35,'协调者',22,(150,169,163),True)
    arrow(s,3.07,2.34,1.26,0.30,TEAL)
    tx(s,4.58,2.30,2.45,0.35,'建设者',22,TEAL,True)
    labels=['理解业务','定义产品','做出体验','接通系统','上线运营','持续迭代']
    x0=0.86; y=3.66
    line(s,x0,y,9.08,y,(71,98,91),2.0)
    for i,lab in enumerate(labels):
        d=0.17+i*0.055; x=x0+i*1.62
        shp=circle(s,x-d/2,y-d/2,d,[MUTE,TEAL,OCHRE,BLUE,ORANGE,TEAL][i]); tag_motif(shp,loud=True)
        tx(s,x-0.53,y+0.45,1.06,0.24,lab,9,(185,204,198),True,align=PP_ALIGN.CENTER)
    tx(s,0.78,4.79,8.46,0.30,'AI让一个人重新有机会，对产品从想法到结果的整条链负责。',13.5,WHITE,True,align=PP_ALIGN.CENTER)
    note(s,40,'收束全课：产品经理不是简单扩张工作量，而是从协调资源重新成为建设产品的人。FDE代表的是端到端责任、能力和控制权。')


def build_slide_41(prs):
    s=add_slide(prs); chrome(s,41,title=TITLES[41])
    steps=[('1','选择已有项目','不要重新想题'),('2','画出当前链路','标出断点与等待'),('3','补齐两个环节','原型、接口、上线任选两项'),('4','定义一次受控使用','谁用、何时用、如何回退'),('5','提交一页复盘','结果、问题与改进')]
    for i,(n,a,b) in enumerate(steps):
        x=0.56+i*1.80
        rounded(s,x,1.58,1.60,2.84,WHITE,[TEAL,BLUE,OCHRE,ORANGE,TEAL][i])
        tx(s,x+0.18,1.84,0.50,0.36,n,20,[TEAL,BLUE,OCHRE,ORANGE,TEAL][i],True)
        tx(s,x+0.18,2.48,1.24,0.52,a,13,INK,True)
        tx(s,x+0.18,3.31,1.24,0.58,b,10,GREY)
    tx(s,0.86,4.70,8.28,0.28,'完成标准：不是“方案更完整”，而是项目向真实使用推进了至少两步。',13,TEAL,True,align=PP_ALIGN.CENTER)
    note(s,41,'布置作业。要求学生选一个已有项目，把当前链路画出来，并真正补齐两个后续环节。重点评估是否接近真实使用。')


def build_generic(prs, page):
    # Only section pages use this path; all other slides have bespoke builders.
    if page == 4:
        return section_divider(prs,4,'01','重新理解产品经理',TITLES[4],'先理解分工从哪里来，才能理解AI为什么会重新改变分工。')
    if page == 9:
        return section_divider(prs,9,'02','扩大 Scope',TITLES[9],'问题不只是工作变多，而是责任与控制范围长期不匹配。')
    if page == 25:
        return section_divider(prs,25,'05','转型路径',TITLES[25],'从“会讨论”走向“能交付”，需要按项目逐步扩展自己的能力边界。')
    if page == 29:
        return section_divider(prs,29,'06','组织形态',TITLES[29],'FDE不是取消专业分工，而是重写0到1阶段的最小交付单元。')
    raise ValueError(page)


BUILDERS = {
    1: build_cover, 2: build_slide_2, 3: build_slide_3,
    5: build_slide_5, 6: build_slide_6, 7: build_slide_7, 8: build_slide_8,
    10: build_slide_10, 11: build_slide_11, 12: build_slide_12, 13: build_slide_13,
    14: build_slide_14, 15: build_slide_15, 16: build_slide_16, 17: build_slide_17,
    18: build_slide_18, 19: build_slide_19, 20: build_slide_20, 21: build_slide_21,
    22: build_slide_22, 23: build_slide_23, 24: build_slide_24,
    26: build_slide_26, 27: build_slide_27, 28: build_slide_28,
    30: build_slide_30, 31: build_slide_31, 32: build_slide_32, 33: build_slide_33,
    34: build_slide_34, 35: build_slide_35, 36: build_slide_36, 37: build_slide_37,
    38: build_slide_38, 39: build_slide_39, 40: build_slide_40, 41: build_slide_41,
}


def main():
    set_palette(deep=DEEP, blue=BLUE, teal=TEAL, magenta=ORANGE, slate=GREY,
                mute=MUTE, mono=INK, font=FONT, display=FONT, eafont=EA,
                eadisplay=EA, accents=[TEAL,ORANGE,BLUE,OCHRE])
    set_geometry(radius=0.12, rule_w=1.0)
    set_ground(BG)
    prs = blank_deck(W,H)
    prs.core_properties.title = 'FDE：产品经理在AI时代的全栈化生存'
    prs.core_properties.subject = 'AI产品经理专题课 · 41页正式授课版'
    prs.core_properties.author = 'AI Course'
    for page in range(1,42):
        if page in BUILDERS:
            BUILDERS[page](prs)
        else:
            build_generic(prs,page)
    # The build is static and intentionally dense enough for a two-hour live class.
    lint_layout(prs, strict=False)
    prs.save(str(OUT))
    declare_delivery(str(OUT), 'presented')
    print(f'saved -> {OUT} | slides: {len(prs.slides)}')


if __name__ == '__main__':
    main()
