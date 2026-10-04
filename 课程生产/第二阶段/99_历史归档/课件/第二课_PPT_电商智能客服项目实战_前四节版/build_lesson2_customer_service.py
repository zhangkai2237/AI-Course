#!/usr/bin/env python3
"""第二阶段第二课：电商智能客服项目实战（前四节，28页）。"""
from pathlib import Path
import importlib.util
import json

ROOT = Path(__file__).resolve().parent
STYLE = Path('/Users/keivn/Project/AI-Course/课程生产/第二阶段/第一课_PPT_正式版/build_lesson1_stage2.py')
SOURCE = Path('/Users/keivn/Project/AI-Course/课程生产/第二阶段/04_第二课_电商智能客服项目实战_老师版完整讲义_v1.0.md')
OUT = ROOT / '第二课_电商智能客服项目实战_前四节_正式版PPT_v1.0.pptx'
GATES = ROOT / '.deck-gates.json'

spec = importlib.util.spec_from_file_location('lesson1_style', STYLE)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
dk = m.dk
m.OUT = OUT
m.SOURCE = SOURCE
m.SLIDE_COUNT = 28
m.PARTS = {
    1: ('课程导入', '电商智能客服项目实战'),
    3: ('第一节', '电商客服业务背景与项目任务'),
    9: ('第二节', '整体方案与主工作流'),
    18: ('第三节', '订单查询工作流'),
    25: ('第四节', '售后咨询与处理工作流'),
}


def extract_block(keyword):
    lines = SOURCE.read_text(encoding='utf-8').splitlines()
    start = level = None
    for i, line in enumerate(lines):
        if line.startswith('#') and keyword in line:
            start = i
            level = len(line) - len(line.lstrip('#'))
            break
    if start is None:
        return '详细讲解参见老师版完整讲义中的对应部分。'
    out = []
    for line in lines[start + 1:]:
        if line.startswith('#'):
            lv = len(line) - len(line.lstrip('#'))
            if lv <= level:
                break
        out.append(line)
    return '\n'.join(out).strip()


def notes_from(slide, intro, *headings):
    blocks = []
    for h in headings or (intro,):
        block = extract_block(h)
        if block not in blocks:
            blocks.append(block)
    dk.speaker_notes(slide, intro + '\n\n' + '\n\n'.join(blocks) + f'\n\n[讲义来源]\n{SOURCE}')


m.notes_from = notes_from


def section(prs, page, number, title, question, next_items):
    s = dk.add_slide(prs)
    dk.slide_background(s, m.BG)
    dk.box(s, 0.28, 0, 0.16, m.H, fill=m.TEAL)
    m.txt(s, 0.94, 0.72, 4.8, 0.30, f'SECTION {number}', 12, m.TEAL, True)
    m.txt(s, 0.94, 1.50, 1.75, 1.10, number, 66, m.TEAL, True, wrap=False)
    m.txt(s, 2.85, 1.50, 8.95, 0.72, title, 35, m.TEXT, True)
    dk.box(s, 2.88, 2.40, 2.20, 0.05, fill=m.CORAL)
    m.txt(s, 2.88, 2.78, 8.80, 0.62, question, 20, m.TEXT, True)
    y = 4.15
    for i, item in enumerate(next_items, 1):
        m.txt(s, 2.90, y + (i-1)*0.55, 0.48, 0.28, f'{i:02d}', 11.5, m.TEAL, True)
        m.txt(s, 3.48, y + (i-1)*0.55, 7.90, 0.34, item, 16, m.MUTE)
    m.txt(s, 11.25, 6.48, 1.10, 0.30, f'{page:02d}', 11.5, m.MUTE, True,
          align=dk.PP_ALIGN.RIGHT)
    notes_from(s, question, title)


def chat(slide, x, y, w, who, text, accent, h=0.92):
    dk.box(slide, x, y, w, h, fill=m.PANEL2 if who == '助手' else m.PANEL,
           line=accent, line_w=0.8, round=True, r=0.16)
    m.txt(slide, x + 0.18, y + 0.12, 0.58, 0.24, who, 11, accent, True)
    m.txt(slide, x + 0.18, y + 0.39, w - 0.36, h - 0.48, text, 14.5, m.TEXT, True)


def simple_flow(slide, items, y=2.36, total_w=11.48, x0=0.90, h=1.48, accent=None):
    accent = accent or m.TEAL
    gap = 0.17
    w = (total_w - gap * (len(items)-1)) / len(items)
    for i in range(len(items)-1):
        x = x0 + i*(w+gap)
        dk.arrow(slide, x+w-0.01, y+h/2-0.09, gap+0.04, 0.18, color=accent)
    for i, (title, body) in enumerate(items):
        x = x0 + i*(w+gap)
        m.card(slide, x, y, w, h, title, body, f'{i+1:02d}', accent, 15.5, 12.5)


def cover(prs):
    s = dk.add_slide(prs)
    dk.slide_background(s, m.BG)
    bar = dk.box(s, 0.28, 0, 0.16, m.H, fill=m.TEAL); dk.tag_motif(bar, loud=True)
    m.txt(s, 0.94, 0.70, 7.0, 0.30, 'AI PRODUCT MANAGER · PROJECT 01', 12, m.TEAL, True)
    m.txt(s, 0.94, 1.50, 10.7, 1.35, '电商智能客服\n项目实战', 45, m.TEXT, True)
    m.txt(s, 0.96, 3.26, 8.8, 0.38, '第二课｜从业务背景到订单查询与售后受理', 20, m.TEXT, True)
    dk.box(s, 0.96, 3.85, 2.10, 0.05, fill=m.CORAL)
    m.txt(s, 0.96, 4.18, 9.4, 0.44, '不是做一个会聊天的窗口，而是把用户请求接进真实业务流程', 18, m.MUTE, True)
    m.txt(s, 10.35, 4.62, 1.70, 1.20, '02', 70, m.TEAL, True,
          align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE, wrap=False)
    notes_from(s, '正式进入第二阶段第一个公共项目。', '第二课：电商智能客服项目实战')


def build():
    prs = dk.blank_deck(m.W, m.H)
    cover(prs)

    m.slide_cards(prs, 2, '今天怎样拆开这个项目', '四个部分，从业务问题逐步进入可执行流程', [
        ('第一节', '业务背景与项目任务', m.TEAL),
        ('第二节', '整体方案与主工作流', m.YELLOW),
        ('第三节', '订单查询工作流', m.CORAL),
        ('第四节', '售后咨询与处理', m.GREEN),
    ], ['第二课：电商智能客服项目实战'], '先看业务，再看系统怎样把一次请求真正处理下去。', cols=4)

    section(prs, 3, '01', '电商客服业务背景与项目任务',
            '为什么不是做一个“会聊天的机器人”？',
            ['看懂购买前、下单后、收货后的客服任务', '理解一句回复背后的查询与操作', '确定本项目重点处理的业务范围'])

    m.slide_cards(prs, 4, '电商客服覆盖购买前后的完整旅程', '用户处在不同阶段，客服要完成的任务也不同', [
        ('购买前', '商品差异、使用场景、预算与推荐', m.TEAL),
        ('下单后', '发货、物流、地址与履约状态', m.YELLOW),
        ('收货后', '使用问题、退换货、退款与投诉', m.CORAL),
    ], ['电商客服每天在处理什么'], '客服不是只回答问题，还要帮助用户完成选择、查询和处理。', cols=3)

    s = m.base(prs, 5, '用户只说了一句话，后台却要完成一串工作', '示例：“上周买的耳机没有声音，我想换一个”')
    chat(s, 0.86, 2.22, 3.05, '用户', '上周买的耳机没有声音，我想换一个。', m.CORAL, 1.18)
    simple_flow(s, [('找订单', '定位商品'), ('问问题', '确认故障'), ('查政策', '判断要求'), ('补材料', '整理申请'), ('交接/提交', '返回状态')],
                y=2.10, x0=4.22, total_w=8.05, h=1.56)
    m.takeaway(s, '评价标准', '不能只看回复像不像人，还要看用户要办的事情有没有被正确推进。', y=5.35)
    notes_from(s, '用一句换货请求拆出后台操作链。', '用户看到的是聊天，客服做的是一连串操作')

    s = m.base(prs, 6, '客服自动化不是从大模型开始', '四种能力会长期共存，而不是互相完全替代')
    simple_flow(s, [('人工客服', '理解灵活\n成本较高'), ('FAQ 机器人', '稳定问答\n难接复杂请求'),
                    ('流程机器人', '步骤确定\n表达受限制'), ('大模型客服', '理解自然表达\n连接既有能力')],
                y=2.25, h=2.10)
    m.takeaway(s, '关键认识', '查询接口、固定流程和业务系统早已有之；大模型让自然表达更好地进入这些流程。', y=5.25)
    notes_from(s, '讲清人工、FAQ、流程机器人和大模型客服的差别。', '从人工客服到大模型客服')

    m.slide_cards(prs, 7, '引入大模型，最值得先改善什么', '请学员先选择，再说明理由', [
        ('A', '回复更自然、更像真人', m.TEAL),
        ('B', '用户用自己的话表达，系统仍能进入正确流程', m.CORAL),
        ('C', '所有售后操作都不再需要人工', m.YELLOW),
    ], ['我们希望用 AI 改善什么'], '本项目优先选择第二项：自然表达之后，仍然要有清楚、可靠的处理步骤。', cols=3)

    m.slide_table(prs, 8, '本项目要解决什么、暂时不做什么', '把业务问题和课堂范围同时说清', [
        ['业务问题', '产品需要改善', '本课边界'],
        ['重复咨询多', '复用商品、政策与流程资料', '不建设完整客服中台'],
        ['用户表达与字段不一致', '把自然表达对应到订单、商品和诉求', '不让模型绕过权限和规则'],
        ['回答后不知道下一步', '给出选择、补充、确认和状态', '不做真实自动退款与赔付'],
        ['转人工后信息断裂', '携带订单、诉求、已问内容和失败原因', '业务操作使用教学模拟环境'],
    ], [2.50, 5.05, 3.97], ['本次电商项目的业务背景', '今天这个助手需要处理哪些请求'],
       size=12.8, row_h=0.78, highlight=2, takeaway_text='重点深入订单查询和售后处理；商品推荐与兜底用于补全整体结构。')

    section(prs, 9, '02', '整体方案与主工作流',
            '用户发来一句话后，系统怎样知道该交给哪条流程？',
            ['先看三段真实感对话', '区分用户界面和后台处理', '设计分流、上下文与状态传递'])

    s = m.base(prs, 10, '同一个对话入口，背后是三种不同处理', '以下均为教学设计中的预期对话')
    cols = [(0.82, '订单查询', '上周买的耳机怎么还没到？', '查到两笔订单，请选择需要查询的一笔。', m.TEAL),
            (4.40, '售后处理', '耳机右侧无声，我想换货。', '先确认订单，再了解故障和要求。', m.CORAL),
            (7.98, '商品推荐', '预算三百，通勤用。', '先确认更看重降噪还是舒适度。', m.YELLOW)]
    for x, title, q, a, color in cols:
        m.card(s, x, 2.02, 3.30, 3.78, title, '', None, color, 18, 14)
        chat(s, x+0.18, 2.82, 2.94, '用户', q, color, 1.02)
        chat(s, x+0.18, 4.02, 2.94, '助手', a, color, 1.34)
    notes_from(s, '展示订单、售后和推荐三段预期对话。', '先看这个产品应该怎样工作')

    s = m.base(prs, 11, '用户看到的是产品，背后运行的是处理系统', '聊天窗口只是入口，不是完整方案')
    dk.box(s, 0.90, 2.08, 11.50, 1.06, fill=m.PANEL2, line=m.TEAL, line_w=1.0, round=True, r=0.16)
    m.txt(s, 1.18, 2.25, 2.0, 0.28, '用户界面', 17, m.TEAL, True)
    m.txt(s, 3.18, 2.23, 8.70, 0.36, '输入问题 · 选择订单或商品 · 补充信息 · 查看结果 · 转人工', 17, m.TEXT, True)
    simple_flow(s, [('主工作流', '决定去哪'), ('子工作流', '处理任务'), ('知识与系统', '提供事实'), ('人工服务', '承接例外'), ('处理记录', '保存状态')],
                y=3.70, h=1.46)
    m.takeaway(s, '产品结构', '工作流负责组织步骤、调用能力，并根据实际结果决定下一步。', y=5.62)
    notes_from(s, '区分用户看到的产品和背后的处理流程。', '先区分“用户看到的产品”和“背后的处理流程”')

    s = m.base(prs, 12, '主工作流：先判断请求要进入哪条业务路径', '主流程负责分流，子流程负责把具体任务处理下去')
    dk.box(s, 0.84, 2.04, 2.10, 1.16, fill=m.PANEL2, line=m.TEAL, round=True, r=0.16)
    m.txt(s, 1.02, 2.30, 1.72, 0.50, '用户消息\n+ 当前状态', 16, m.TEXT, True, align=dk.PP_ALIGN.CENTER)
    dk.arrow(s, 2.94, 2.50, 0.60, 0.20, color=m.TEAL)
    dk.box(s, 3.55, 2.10, 2.20, 1.02, fill=m.PANEL, line=m.CORAL, round=True, r=0.16)
    m.txt(s, 3.78, 2.37, 1.72, 0.32, '判断处理路径', 17, m.TEXT, True, align=dk.PP_ALIGN.CENTER)
    dk.arrow(s, 5.72, 2.50, 0.62, 0.20, color=m.TEAL)
    m.txt(s, 6.45, 2.38, 1.30, 0.30, '分流到', 14, m.MUTE, True)
    branches = [('订单查询', m.TEAL), ('售后与投诉', m.CORAL), ('商品咨询推荐', m.YELLOW), ('其他与兜底', m.GREEN)]
    branch_x0, branch_w = 1.02, 2.70
    branch_gap = (11.48 - branch_w * len(branches)) / (len(branches) - 1)
    for i, (name, color) in enumerate(branches):
        x = branch_x0 + i * (branch_w + branch_gap); y = 3.48
        dk.box(s, x, y, branch_w, 1.16, fill=m.PANEL, line=color, round=True, r=0.16)
        m.txt(s, x+0.18, y+0.36, branch_w-0.36, 0.32, name, 16, m.TEXT, True, align=dk.PP_ALIGN.CENTER)
    m.takeaway(s, '拆分原则', '订单流程的修改不应要求同时重写商品推荐流程。', y=5.47)
    notes_from(s, '讲解四条主分支和主、子工作流的分工。', '主工作流怎样组织')

    m.slide_table(prs, 13, '四条分支需要先写清业务定义', '分类不是只列几个标签，还要说明边界和示例', [
        ['分支', '主要接收的需求', '用户示例'],
        ['订单查询', '具体订单、发货与物流查询', '“耳机到哪了？”'],
        ['售后与投诉', '政策、故障、退换货诉求和投诉', '“耳机坏了，我想换。”'],
        ['商品咨询推荐', '功能、比较、使用场景和购买建议', '“通勤买哪款？”'],
        ['其他与兜底', '寒暄、表达不完整、范围外请求', '“你好。”“我想问一下。”'],
    ], [2.30, 5.30, 3.92], ['意图分类不能只写几个标签'],
       size=13.2, row_h=0.76, highlight=2, takeaway_text='分类规则本身也是产品设计：类别含糊，后面的流程再正确也会被送错。')

    m.slide_quote_analysis(prs, 14, '“这个耳机能不能换？”到底是什么意思', '这个耳机能不能换？', [
        ('购买前', '询问某款商品的换货政策'),
        ('购买后', '想为自己的订单申请换货'),
        ('当前推荐', '想把刚才推荐的耳机换成另一款'),
    ], ['意图分类不能只写几个标签'], quote_label='用户表达',
       takeaway_text='同一句话可能对应不同任务，必须结合当前消息、页面对象和对话上下文。')

    s = m.base(prs, 15, '多轮对话：先知道“正在做什么”，再理解这一句话', '“第二个”和“右边”脱离上下文都没有意义')
    m.card(s, 0.84, 2.05, 5.55, 3.30, '订单选择', '', '01', m.TEAL, 18, 14)
    chat(s, 1.08, 3.05, 5.02, '助手', '查询到两笔耳机订单，请选择一笔。', m.TEAL, 0.90)
    chat(s, 1.08, 4.18, 3.28, '用户', '第二个。', m.CORAL, 0.78)
    m.txt(s, 4.55, 4.36, 1.25, 0.30, '→ DEMO-102', 13.5, m.TEAL, True)
    m.card(s, 6.72, 2.05, 5.55, 3.30, '故障补充', '', '02', m.CORAL, 18, 14)
    chat(s, 6.96, 3.05, 5.02, '助手', '两边都没有声音，还是只有一边？', m.CORAL, 0.90)
    chat(s, 6.96, 4.18, 3.28, '用户', '右边。', m.TEAL, 0.78)
    m.txt(s, 10.22, 4.36, 1.62, 0.30, '→ 补充故障信息', 13.5, m.CORAL, True)
    notes_from(s, '用两个短句说明当前任务和等待字段的重要性。', '用几句话检查分类是否清楚', '不是每一句话都重新从头分流')

    m.slide_table(prs, 16, '四条业务分支与三种处理方式', '分支按用户任务组织；处理方式在分支内部连续配合', [
        ['业务分支', '知识问答', '业务查询', '业务办理'],
        ['订单查询', '发货说明', '订单与物流', '改址入口／人工'],
        ['售后与投诉', '退换货政策', '订单与申请', '补材料、确认、提交'],
        ['商品咨询推荐', '参数与使用说明', '库存与价格', '加入对比／购买'],
        ['其他与兜底', '服务范围', '已有任务状态', '转人工／结束'],
    ], [2.35, 3.05, 3.05, 3.07], ['“三类任务”和“四条分支”是什么关系'],
       size=12.8, row_h=0.78, highlight=2, takeaway_text='不要把所有知识问答、接口操作割裂开；同一个用户任务往往需要它们连续协作。')

    m.slide_table(prs, 17, '流程之间不能只传一句话', '输入要带上下文，输出要带处理状态', [
        ['传入子流程', '子流程返回'],
        ['当前消息与用户身份', '等待用户补充：任务继续'],
        ['已确认的订单或商品', '已完成查询：返回结果'],
        ['已经收集的必要信息', '需要人工：自动流程暂停'],
        ['当前任务与等待内容', '执行失败：进入重试或人工'],
    ], [5.70, 5.82], ['各流程之间需要传递什么'],
       size=14.2, row_h=0.80, highlight=2, takeaway_text='回复内容必须和实际状态一致：接口失败时，模型不能说“已经处理好了”。')

    section(prs, 18, '03', '订单查询工作流',
            '意图已经识别出来，系统怎样找到用户真正问的那笔订单？',
            ['准备可控的模拟订单数据', '处理无订单号与多订单选择', '区分无结果、待发货和接口失败'])

    m.slide_table(prs, 19, '先准备三笔可验证的模拟订单', '查询日固定为 9 月 9 日，避免“上周”随授课日期变化', [
        ['订单编号', '商品与规格', '下单时间', '订单状态', '最新物流'],
        ['DEMO-101', 'AirLite 耳机｜白色', '9 月 1 日', '已签收', '9 月 4 日签收'],
        ['DEMO-102', 'AirLite Pro 黑色', '9 月 3 日', '已发货', '9 月 8 日到达分拨中心'],
        ['DEMO-103', 'K68 键盘｜灰色', '9 月 5 日', '待发货', '暂无物流记录'],
    ], [1.85, 3.10, 1.65, 1.65, 3.27], ['先准备这个流程需要的数据'],
       size=12.8, row_h=0.86, highlight=2, takeaway_text='“上周买的耳机”可能对应两笔订单，不能直接替用户选择。')

    m.slide_cards(prs, 20, '用户没有订单号，下一步怎么做', '先选择，再比较用户成本和业务风险', [
        ('A', '要求用户找到订单号后再查询', m.TEAL),
        ('B', '根据当前账号查询近期相关订单，请用户确认', m.CORAL),
        ('C', '直接选择最近一笔耳机订单', m.YELLOW),
    ], ['没有订单号，应该怎样查询'], '当前场景选择第二项：系统先筛选候选订单，信息有歧义时由用户确认。', cols=3)

    s = m.base(prs, 21, '订单查询子工作流', '从自然语言线索进入确定的订单与物流数据')
    simple_flow(s, [('检查信息', '身份\n已有订单'), ('查询候选', '商品\n时间范围'), ('判断数量', '0 / 1 / 多笔'),
                    ('用户确认', '目标订单'), ('查询详情', '订单\n物流'), ('检查结果', '成功\n失败')],
                y=2.13, h=2.05)
    m.takeaway(s, '权限边界', '身份由登录系统提供；订单筛选和访问权限由系统完成，不把整个数据库交给模型。', y=5.25)
    notes_from(s, '完整讲解订单查询的六个步骤。', '展开订单查询子工作流')

    s = m.base(prs, 22, '“第二个”怎样准确对应到 DEMO-102', '自然语言与页面操作可以同时存在')
    m.card(s, 0.84, 2.02, 3.25, 3.40, '候选订单', '① DEMO-101 白色\n9 月 1 日下单\n\n② DEMO-102 黑色\n9 月 3 日下单', '01', m.TEAL, 18, 15)
    dk.arrow(s, 4.15, 3.05, 0.74, 0.22, color=m.CORAL)
    chat(s, 4.95, 2.70, 2.10, '用户', '第二个。', m.CORAL, 0.92)
    dk.arrow(s, 7.12, 3.05, 0.74, 0.22, color=m.TEAL)
    m.card(s, 7.93, 2.02, 4.28, 3.40, '会话状态更新', 'waiting_for: clear\nselected_order_id: DEMO-102\ncurrent_task: order_query', '02', m.CORAL, 18, 15)
    m.takeaway(s, '产品细节', '点击订单卡片可以直接完成选择；输入“第二个”则必须结合刚才展示的候选列表。', y=5.72)
    notes_from(s, '讲解候选列表、等待状态和选择结果的对应。', '用户说“第二个”，系统怎样接住')

    m.slide_compare(prs, 23, '拿到查询结果以后，模型负责什么', '事实来自系统，模型负责组织表达并遵守边界',
                    '系统直接提供', ['订单状态与物流节点', '事件发生时间', '是否存在预计送达时间', '接口是否成功'],
                    '模型可以完成', ['组合成清楚的用户回复', '说明当前已知与未知', '提出可执行的下一步', '不补写系统没有的事实'],
                    ['查询订单详情与物流信息', '哪些内容需要模型，哪些可以直接展示'], m.GREEN,
                    '没有预计到达时间，不承诺日期；没有延误原因，不编造天气或运输问题。')

    m.slide_table(prs, 24, '看不到物流，不代表同一种情况', '三个状态需要不同回复与后续处理', [
        ['实际状态', '真正含义', '产品处理'],
        ['没有找到订单', '查询成功，但没有匹配记录', '调整时间、确认账号或补充信息'],
        ['订单尚未发货', '已经找到订单，暂时无物流轨迹', '如实展示待发货，不承诺时间'],
        ['查询工具失败', '系统没有成功获取数据', '说明暂时无法查询，提供重试或人工'],
    ], [2.35, 4.25, 4.92], ['三种容易被混淆的查询结果', '用测试检查这条流程'],
       size=13.2, row_h=0.84, highlight=3, takeaway_text='课堂重点测试：正常查询、多订单选择、接口失败；错误必须定位到具体步骤。')

    section(prs, 25, '04', '售后咨询与处理工作流',
            '从“想换一个”到“申请已受理”，中间还缺哪些产品步骤？',
            ['区分问政策与发起处理', '复用订单和故障信息，查询适用要求', '确认、模拟提交并返回真实状态'])

    s = m.base(prs, 26, '售后主线：从政策咨询走到申请受理', '本节使用订单编号为 DEMO-101 的已签收订单，处理右侧耳机无声的换货诉求')
    simple_flow(s, [('判断诉求', '问政策\n还是要处理'), ('确认订单', 'DEMO-101'), ('了解故障', '右侧无声'),
                    ('查询要求', '政策与材料'), ('用户确认', '申请内容'), ('提交申请', '等待审核')],
                y=2.02, h=1.88, accent=m.CORAL)
    dk.box(s, 1.05, 4.45, 5.25, 0.74, fill=m.PANEL, line=m.TEAL, round=True, r=0.16)
    m.txt(s, 1.28, 4.67, 4.78, 0.30, '可以：受理申请、返回编号和当前状态', 15.5, m.TEXT, True)
    dk.box(s, 6.62, 4.45, 5.25, 0.74, fill=m.PANEL, line=m.CORAL, round=True, r=0.16)
    m.txt(s, 6.85, 4.67, 4.78, 0.30, '不可以：自行批准、退款、赔偿或承诺寄出', 15.5, m.TEXT, True)
    m.takeaway(s, '核心边界', '受理申请，不等于批准申请。', y=5.65)
    notes_from(s, '先讲清问政策、具体咨询和申请处理的差别，再讲范围边界。', '先区分“问政策”和“要处理”', '先确定这条售后流程处理到哪里', '先看售后流程全貌')

    s = m.base(prs, 27, '售后流程怎样决定下一步', '先复用用户已经说过的内容，再按业务要求补齐缺失信息')
    m.card(s, 0.82, 2.00, 3.55, 3.65, '复用已有信息', 'DEMO-101\n换货诉求\n收到后右侧无声\n已经重新连接', '01', m.TEAL, 18, 15)
    m.card(s, 4.62, 2.00, 3.55, 3.65, '知识与规则', '知识：排查 · 受理 · 流程 · 特殊情况\n\n规则：完整性 · 重复申请 · 权限', '02', m.YELLOW, 18, 15)
    m.card(s, 8.42, 2.00, 3.88, 3.65, '责任分工', '模型：理解与追问\n系统：查询、校验、提交\n人工：例外和最终处理', '03', m.CORAL, 18, 15)
    m.takeaway(s, '提问原则', '不是把整张表单重新问一遍，而是只问当前真正缺少、会影响处理的内容。', y=5.88)
    notes_from(s, '讲解复用信息、缺失字段、售后知识材料和模型/规则/系统/人工分工。', '确认订单，同时复用用户已经提供的信息', '根据处理要求决定下一句问什么', '售后知识库里应该准备什么', '知识库、规则和模型分别负责什么')

    s = m.base(prs, 28, '确认、提交与异常：最后一句话必须对应真实结果', '申请会改变业务数据，不能只关注回复是否流畅')
    m.card(s, 0.82, 2.00, 3.62, 3.75, '申请确认页', 'AirLite 白色\nDEMO-101\n换货｜收到后右侧无声\n材料已添加\n\n确认提交｜修改信息', '01', m.TEAL, 18, 14.5)
    m.card(s, 4.67, 2.00, 3.62, 3.75, '模拟提交结果', 'accepted: true\nAS-DEMO-001\npending_review\n\n申请已提交\n目前等待审核', '02', m.GREEN, 18, 14.5)
    m.card(s, 8.52, 2.00, 3.78, 3.75, '关键异常', '改诉求：重新确认\n提交超时：先查状态\n政策冲突：转人工\n投诉升级：携带上下文', '03', m.CORAL, 18, 14.5)
    m.takeaway(s, '最终检查', '不能把“已提交”说成“已批准”，也不能把“超时未知”说成成功或失败。', y=5.96)
    notes_from(s, '完整讲解确认页、修改失效、模拟提交、幂等与超时、投诉转人工和三组测试。', '把已收集的信息整理成申请确认页', '模拟提交后，按真实结果回复', '投诉与特殊情况怎样处理', '用三组案例检查这条流程')

    dk.declare_delivery(OUT, 'presented')
    dk.lint_layout(prs, strict=True)
    prs.save(OUT)
    gates = {
        'delivery': 'presented',
        'content_plan_ref': str(SOURCE),
        'output': str(OUT),
        'slide_count': 28,
        'content': {'checkpoint': {'mode': 'approved', 'record': '用户确认28页结构。'}},
        'design': {'checkpoint': {'mode': 'approved', 'record': '沿用第一课深色青绿体系，强化章节和流程。'}},
        'design_plan': {
            'mode': 'Mode A mimic',
            'style_pick': 'n/a — locked recent-course style',
            'boldness': 'balanced+',
            'signature_move': '用户消息沿服务路由穿过主工作流、订单与售后状态',
            'palette': {'ground': '#031016', 'text': '#F5F2E9', 'teal': '#36D6C2', 'coral': '#FF7657'},
            'type_scale': {'display': 45, 'section': 35, 'title': 29, 'body': 16},
            'form_ledger': '封面1；目录1；章节页4；对话页3；流程页5；表格页7；比较/框架页7',
            'interior_register': '大号章节编号、深色画布、青绿路由、珊瑚风险、底部进度线',
            'density': {'planned_median': 38, 'over_70': 2, 'non_text_protagonist': 17},
            'build_script': str(Path(__file__).resolve()),
        },
        'provenance': {'source_only': str(SOURCE), 'no_external_claims': True},
    }
    GATES.write_text(json.dumps(gates, ensure_ascii=False, indent=2), encoding='utf-8')
    print(OUT)


if __name__ == '__main__':
    build()
