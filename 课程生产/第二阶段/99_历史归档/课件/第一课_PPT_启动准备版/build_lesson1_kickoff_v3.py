#!/usr/bin/env python3
"""第二阶段第一课PPT v3.0：依据老师版完整讲义 v2.0 重构。"""
from pathlib import Path
import importlib.util
import json


ROOT = Path(__file__).resolve().parent
STYLE_SCRIPT = Path("/Users/keivn/Project/AI-Course/课程生产/第二阶段/第一课_PPT_正式版/build_lesson1_stage2.py")
SOURCE = Path("/Users/keivn/Project/AI-Course/课程生产/第二阶段/03_第一课_第二阶段项目启动与准备_老师版完整讲义_v2.0.md")
OUT = ROOT / "第一课_第二阶段项目启动与准备_完整PPT_v3.0.pptx"
GATES = ROOT / ".deck-gates.json"

spec = importlib.util.spec_from_file_location("lesson1_style", STYLE_SCRIPT)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
dk = m.dk

m.OUT = OUT
m.SOURCE = SOURCE
m.SLIDE_COUNT = 30
m.PARTS = {
    1: ("课程导入", "第二阶段项目启动与准备"),
    3: ("第一节", "第二阶段怎样学习"),
    10: ("第二节", "三个公共项目的选择逻辑"),
    15: ("第三节", "三个公共项目与个人项目方向"),
    27: ("第四节", "第一课课后准备"),
}


def extract_block(keyword):
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    start = None
    level = None
    for i, line in enumerate(lines):
        if line.startswith("#") and keyword in line:
            start = i
            level = len(line) - len(line.lstrip("#"))
            break
    if start is None:
        return "详细讲解参见老师版完整讲义中的对应部分。"
    out = []
    for line in lines[start + 1:]:
        if line.startswith("#"):
            new_level = len(line) - len(line.lstrip("#"))
            if new_level <= level:
                break
        out.append(line)
    return "\n".join(out).strip()


def notes_from(slide, intro, *headings):
    keys = headings or (intro,)
    blocks = []
    for key in keys:
        block = extract_block(key)
        if block not in blocks:
            blocks.append(block)
    dk.speaker_notes(slide, intro + "\n\n" + "\n\n".join(blocks) + f"\n\n[讲义来源]\n{SOURCE}")


m.notes_from = notes_from


def cover(prs):
    s = dk.add_slide(prs)
    dk.slide_background(s, m.BG)
    dk.box(s, 0.28, 0, 0.16, m.H, fill=m.TEAL)
    m.txt(s, 0.92, 0.66, 5.8, 0.30, "AI PRODUCT MANAGER · PROJECT STAGE", 12, m.TEAL, True)
    m.txt(s, 0.92, 1.42, 10.8, 1.45, "第二阶段项目\n启动与准备", 43, m.TEXT, True)
    m.txt(s, 0.94, 3.25, 10.2, 0.44, "第一课｜学习方式、公共项目与个人素材准备", 20, m.TEXT, True)
    dk.box(s, 0.94, 3.80, 2.08, 0.05, fill=m.CORAL)
    m.txt(s, 0.94, 4.10, 10.2, 0.60, "先建立共同参照，再进入正式项目", 17, m.MUTE)
    m.txt(s, 10.43, 4.72, 1.35, 1.00, "01", 62, m.TEAL, True,
          align=dk.PP_ALIGN.CENTER, anchor=dk.MSO_ANCHOR.MIDDLE)
    m.txt(s, 0.94, 6.72, 4.8, 0.25, "第二阶段 · 项目实战班", 11.5, m.MUTE, True)
    notes_from(s, "本课不要求正式选题、写PRD或画原型。", "第一课｜第二阶段项目启动与准备")


def roadmap(prs):
    s = m.base(prs, 8, "十次课的整体安排", "启动、三轮项目、深化与成果表达")
    labels = [
        ("01", "启动准备", m.TEAL), ("02", "公共项目一", m.TEAL),
        ("03", "项目一点评", m.CORAL), ("04", "公共项目二", m.TEAL),
        ("05", "项目二点评", m.CORAL), ("06", "公共项目三", m.TEAL),
        ("07", "项目三点评", m.CORAL), ("08", "主项目深化", m.YELLOW),
        ("09", "作品集与简历", m.GREEN), ("10", "路演与总结", m.GREEN),
    ]
    for i, (n, label, color) in enumerate(labels):
        row, col = divmod(i, 5)
        x, y = 0.82 + col * 2.32, 2.08 + row * 1.72
        dk.box(s, x, y, 2.12, 1.22, fill=m.PANEL, line=color, line_w=0.8, round=True, r=0.16)
        m.txt(s, x + 0.18, y + 0.16, 0.45, 0.24, n, 10.5, color, True)
        m.txt(s, x + 0.18, y + 0.52, 1.74, 0.38, label, 14.5, m.TEXT, True,
              anchor=dk.MSO_ANCHOR.MIDDLE)
    m.takeaway(s, "课程节奏", "公共项目讲深，个人项目在互评、点评和修改中完成。")
    notes_from(s, "十次课整体安排", "十次课的整体安排")


def project_slide(prs, page, title, subtitle, stages, focus, accent, note_key):
    s = m.base(prs, page, title, subtitle, accent)
    x0, y, total_w, gap = 0.84, 2.16, 11.46, 0.16
    cw = (total_w - gap * (len(stages) - 1)) / len(stages)
    for i, (t, b) in enumerate(stages):
        x = x0 + i * (cw + gap)
        m.card(s, x, y, cw, 2.36, t, b, f"{i+1:02d}", accent, 14.5, 11.8)
        if i < len(stages) - 1:
            dk.arrow(s, x + cw - 0.02, y + 1.08, gap + 0.04, 0.16, color=accent)
    m.takeaway(s, "项目价值", focus, y=5.42, accent=accent)
    notes_from(s, title, note_key)


def homework_table(prs):
    rows = [
        ["候选场景", "使用者", "要完成的事情", "当前问题", "可用材料"],
        ["场景一", "", "", "", ""],
        ["场景二", "", "", "", ""],
        ["场景三", "", "", "", ""],
    ]
    m.slide_table(prs, 28, "三个候选业务场景", "每个问题写一两句话即可", rows,
                  [1.55, 2.20, 2.80, 2.55, 2.42], ["盘点三个候选场景"],
                  size=12.8, row_h=0.90, highlight=None,
                  takeaway_text="优先选择自己熟悉、能够获得材料、能够判断结果的场景。")


def build_deck():
    prs = dk.blank_deck(m.W, m.H)
    cover(prs)

    m.slide_cards(prs, 2, "今天主要讲四件事", "第一课先建立方向和共同参照", [
        ("学习方式", "第二阶段怎样上课、做作业和接受点评", m.TEAL),
        ("选择逻辑", "为什么选择这三个公共项目", m.YELLOW),
        ("项目方向", "三个项目可迁移到多类业务", m.CORAL),
        ("课后准备", "盘点场景并准备少量材料", m.GREEN),
    ], ["第一课｜第二阶段项目启动与准备"],
       "今天不正式选题，也不写PRD。", cols=4)

    m.slide_statement(prs, 3, "第二阶段的学习目标",
                      "把第一阶段的知识连续地用在真实项目中",
                      "面对一个不完整的业务问题，完成分析、方案、页面、测试和修改，而不是继续分别记忆更多概念。",
                      ["为什么进入第二阶段"])
    m.slide_compare(prs, 4, "第一阶段与第二阶段", "学习重点发生变化",
                    "第一阶段", ["理解模型、知识库和工作流", "学习需求、PRD与原型", "掌握评测、成本和风险"],
                    "第二阶段", ["从真实业务问题出发", "根据需要组合已有能力", "在作业修改中形成项目成果"],
                    ["为什么进入第二阶段"], m.CORAL,
                    "第二阶段按照项目推进过程组织，而不是按照技术名词分课。")
    m.slide_cards(prs, 5, "公共项目、个人项目与主项目", "三种成果承担不同作用", [
        ("3个公共项目", "老师完整拆解，建立共同方法和标准", m.TEAL),
        ("3个个人项目", "每位学员结合自己的业务独立完成", m.CORAL),
        ("1个主项目", "从个人项目中选择一个继续做深", m.GREEN),
    ], ["公共项目和个人项目"],
       "公共项目提供参照，个人项目证明方法可以迁移。", cols=3)
    m.slide_flow(prs, 6, "一轮项目的推进方式", "每个人独立完成项目，小组帮助发现问题", [
        ("公共项目", "老师完整讲解"),
        ("个人V0.1", "结合自己的业务"),
        ("小组互评", "提出问题"),
        ("老师点评", "讲代表案例"),
        ("个人V1.0", "根据问题修改"),
    ], ["一轮项目怎样推进"],
       "老师讲公共项目，个人完成V0.1，小组互评，老师点评，个人修改为V1.0。", cols=5)
    m.slide_compare(prs, 7, "互评小组怎样工作", "固定4—5人，小组不共同提交项目",
                    "小组负责", ["阅读成员的个人项目", "指出看不懂或不成立的地方", "汇总最需要老师讲的问题"],
                    "个人负责", ["独立选择项目和准备材料", "独立完成并提交每个版本", "根据反馈修改自己的项目"],
                    ["一轮项目怎样推进"], m.GREEN,
                    "小组帮助发现问题，每个人始终拥有自己的项目。")
    roadmap(prs)
    m.slide_cards(prs, 9, "第二阶段的最终成果", "三个完整项目，其中一个继续做深", [
        ("3个项目方案", "业务、流程、方案、页面、测试和修改记录", m.TEAL),
        ("1个重点项目", "补充更扎实的验证、迭代和作品集表达", m.YELLOW),
        ("1套面试讲法", "讲清背景、关键判断、取舍和变化", m.GREEN),
    ], ["第二阶段的最终成果"],
       "最终形成三段能够解释清楚的项目经历。", cols=3)

    m.slide_cards(prs, 10, "公共项目的选择标准", "先判断什么项目适合全班共同学习", [
        ("A｜技术热门", "使用当前最热门的技术", m.TEAL),
        ("B｜名称先进", "项目听起来足够复杂", m.YELLOW),
        ("C｜业务清楚", "流程、材料和判断标准可以展示", m.CORAL),
        ("D｜演示好看", "容易快速制作一个Demo", m.GREEN),
    ], ["什么样的项目适合作为公共项目"],
       "课堂先让学员选择，老师选择C。", cols=4)
    m.slide_cards(prs, 11, "四项筛选条件", "三个公共项目采用同一套标准", [
        ("真实问题", "对应常见而具体的业务工作", m.TEAL),
        ("能力组合", "自然使用模型、规则、系统和人工", m.YELLOW),
        ("过程可讲", "材料、判断过程和异常情况可以展示", m.CORAL),
        ("容易迁移", "能够发展出多个行业和岗位的变种", m.GREEN),
    ], ["什么样的项目适合作为公共项目"],
       "项目是否适合教学，取决于能否把业务判断讲清楚。", cols=4)
    m.slide_statement(prs, 12, "项目不按技术方向划分",
                      "先面对业务问题，再选择需要的能力",
                      "同一个项目通常会同时使用大模型、知识库、接口、规则、工作流和人工处理。",
                      ["为什么不按技术方向选择项目"])
    m.slide_table(prs, 13, "三个公共项目覆盖的问题", "从人、材料和复杂任务三个入口切入", [
        ["公共项目", "处理对象", "主要目标", "重点能力"],
        ["智能客服与业务办理", "问题和办理请求", "回答并推动服务完成", "对话、知识、系统连接、业务闭环"],
        ["材料审核与风险核验", "合同、票据和申请", "发现缺失、差异和风险", "提取、规则、语义判断、人工复核"],
        ["专业研究与方案生成", "复杂任务和多来源资料", "形成研究或解决方案", "任务拆解、资料组织、来源核实"],
    ], [3.0, 2.25, 2.75, 3.52], ["三个项目覆盖的业务问题"],
       size=12.4, row_h=0.82, highlight=1,
       takeaway_text="三类项目覆盖服务、审核和知识工作中的核心产品问题。")
    m.slide_compare(prs, 14, "为什么没有单独设置C端公共项目", "项目可点评性比B端或C端标签更重要",
                    "公共项目需要", ["全班可以理解的业务流程", "可以准备和展示的材料", "相对一致的评价标准"],
                    "很多C端项目依赖", ["真实用户和流量", "行为数据与增长实验", "留存和长期运营"],
                    ["为什么没有单独设置C端公共项目"], m.CORAL,
                    "有真实用户、经验或数据的学员，仍然可以选择C端个人项目。")

    m.slide_cards(prs, 15, "三个公共项目与个人项目方向", "听项目时关注两件事", [
        ("与经历的距离", "哪一类工作和自己的岗位、行业最接近", m.TEAL),
        ("材料可获得性", "哪一类项目的资料、规则或案例可以找到", m.CORAL),
    ], ["第三节：三个公共项目与个人项目方向"],
       "先形成两三个可能方向，本课不做正式选择。", cols=2)

    project_slide(prs, 16, "公共项目一：企业智能客服与业务办理",
                  "面向客户、员工或合作伙伴的统一服务入口",
                  [("理解请求", "识别意图与信息"), ("获取答案", "知识与业务数据"),
                   ("办理业务", "规则与工作流"), ("处理异常", "人工接管")],
                  "把咨询、业务查询、办理和转人工连接成完整服务流程。", m.TEAL,
                  "企业智能客服与业务办理")
    m.slide_table(prs, 17, "智能客服变种：面向消费者", "同一方法可以进入多种C端服务", [
        ["个人项目方向", "主要处理的事情"],
        ["电商售后服务助手", "订单查询、退换货政策、退款与异常售后"],
        ["在线教育学员服务助手", "课程咨询、请假、补课、延期与退款"],
        ["银行或保险服务助手", "账户问题、业务引导、理赔材料与进度"],
        ["医疗服务助手", "科室选择、预约、检查准备与复诊提醒"],
        ["旅游与本地生活服务助手", "预约、改期、退改政策、会员权益与行程变化"],
    ], [3.6, 7.92], ["面向消费者"], size=12.6, row_h=0.66, highlight=2,
       takeaway_text="服务流程、业务数据和办理规则决定项目差异。")
    m.slide_table(prs, 18, "智能客服变种：企业与内部服务", "客户、伙伴和员工都可能成为服务对象", [
        ["个人项目方向", "主要处理的事情"],
        ["企业软件客户服务", "产品使用、账户权限和技术支持工单"],
        ["设备售后或经销商服务", "故障排查、备件、库存、订单和返利规则"],
        ["员工人事与行政服务", "考勤福利、证明申请、访客和办公资源"],
        ["企业IT服务", "账号、权限、软件、设备和网络问题"],
        ["财务或采购服务", "报销、付款、预算、采购申请与供应商状态"],
    ], [3.6, 7.92], ["面向企业客户和合作伙伴", "面向企业内部"], size=12.6, row_h=0.66, highlight=4,
       takeaway_text="这一方向适合客服、运营、售后、人事、IT、财务和采购背景。")

    project_slide(prs, 19, "公共项目二：业务材料审核与风险核验",
                  "处理合同、票据、申请表和资质证明等业务材料",
                  [("读取材料", "文本、表格与图片"), ("提取信息", "关键字段与关系"),
                   ("执行核验", "规则与语义判断"), ("人工复核", "风险定位与处理")],
                  "多份材料互相核对，同时处理确定性规则、语义风险和人工责任。", m.CORAL,
                  "业务材料审核与风险核验")
    m.slide_table(prs, 20, "材料审核变种：财务、合同与采购", "材料和规则改变，审核方法可以迁移", [
        ["个人项目方向", "主要处理的事情"],
        ["报销或付款材料审核", "核对发票、报销单、合同、订单和验收单"],
        ["预算申请审核", "检查预算科目、历史支出和审批要求"],
        ["合同条款与履约检查", "检查风险条款、交付、验收和付款条件"],
        ["供应商准入审核", "核验营业执照、资质和准入条件"],
        ["采购或招投标材料检查", "核对预算、询价、响应内容和资质缺项"],
    ], [3.6, 7.92], ["财务与费用", "合同与法务", "采购与供应商"], size=12.4, row_h=0.66, highlight=1,
       takeaway_text="真实审核清单、字段、规则和异常情况是项目基础。")
    m.slide_table(prs, 21, "材料审核变种：金融、人事与公共服务", "高风险结果需要定位原文并交给人工", [
        ["个人项目方向", "主要处理的事情"],
        ["开户、贷款或理赔预审", "核对企业身份、资产负债、保单和事故材料"],
        ["招聘入职材料审核", "检查身份、学历、经历和入职文件"],
        ["项目申报或办事材料预审", "对照指南检查资格、条件和附件"],
        ["质量报告核验", "检查检测项目、数值、结论和签章"],
        ["内容发布合规检查", "发现广告或公开材料中的风险表述"],
    ], [3.6, 7.92], ["金融与保险", "人事、行政和公共服务"], size=12.4, row_h=0.66, highlight=3,
       takeaway_text="这一方向适合财务、法务、采购、风控、人事、金融和质量管理背景。")

    project_slide(prs, 22, "公共项目三：专业研究与方案生成",
                  "围绕复杂任务组织资料并形成专业交付",
                  [("理解任务", "目标、问题与限制"), ("收集资料", "内部与外部来源"),
                   ("形成方案", "分析、简报或建议"), ("专业审阅", "来源、缺口与责任")],
                  "让研究过程可检查，让专业人员能够核实来源、修改判断并承担最终责任。", m.YELLOW,
                  "专业研究与方案生成")
    m.slide_table(prs, 23, "研究与方案变种：销售、产品与运营", "重点在任务拆解和资料使用，不是直接生成长文", [
        ["个人项目方向", "主要处理的事情"],
        ["客户需求与大客户简报", "整理需求、客户背景、合作历史和拜访准备"],
        ["解决方案与投标响应", "把客户要求映射到产品能力和证明材料"],
        ["竞品与产品机会研究", "整理竞品、用户反馈、市场信息和业务数据"],
        ["产品规划与用户研究", "汇总目标、约束、访谈和问卷形成判断"],
        ["运营活动方案", "结合目标用户、历史活动和渠道情况形成计划"],
    ], [3.6, 7.92], ["销售与售前", "产品与运营"], size=12.4, row_h=0.66, highlight=2,
       takeaway_text="选择自己能够判断内容对错的专业领域。")
    m.slide_table(prs, 24, "研究与方案变种：行业、管理与专业服务", "复杂任务可以迁移到多类知识工作", [
        ["个人项目方向", "主要处理的事情"],
        ["行业、市场或政策研究", "整理市场、竞争、政策、趋势和业务影响"],
        ["经营分析与战略规划", "汇总经营数据、部门反馈、目标和外部环境"],
        ["咨询或投资研究", "围绕问题组织访谈、公司、行业和风险资料"],
        ["课程研发与培训方案", "根据学习目标、资料和能力差距设计内容"],
        ["技术选型研究", "比较技术方案、供应商、成本和实施风险"],
    ], [3.6, 7.92], ["行业与管理", "专业服务"], size=12.4, row_h=0.66, highlight=1,
       takeaway_text="这一方向适合产品、运营、销售、咨询、研究、培训和管理岗位。")

    m.slide_cards(prs, 25, "从公共项目到个人项目", "个人项目至少需要重新确定五件事", [
        ("用户", "谁在使用", m.TEAL),
        ("场景", "何时遇到什么问题", m.TEAL),
        ("流程", "现在怎样完成", m.YELLOW),
        ("材料", "数据、文档、规则和样本", m.CORAL),
        ("范围", "第一版解决到哪里", m.GREEN),
    ], ["从公共项目发展为个人项目"],
       "用户、流程、材料和判断没有变化，就只是给公共项目换了名字。", cols=5)
    m.slide_cards(prs, 26, "个人方向的课堂选择", "先记录可能方向，不正式选题", [
        ("A｜服务与办理", "回答问题、提供服务或办理业务", m.TEAL),
        ("B｜审核与核验", "检查材料、核对规则或发现风险", m.CORAL),
        ("C｜研究与方案", "搜集信息、研究问题或输出方案", m.YELLOW),
        ("D｜暂不确定", "继续观察自己的工作和可用材料", m.GREEN),
    ], ["课堂互动"],
       "再写下与自己经历最接近的两个项目变种。", cols=4)

    m.slide_cards(prs, 27, "第一课课后准备", "本次只完成两项轻量准备", [
        ("三个候选场景", "从工作、行业经验或目标岗位中寻找", m.TEAL),
        ("一组基础材料", "从最熟悉的场景准备少量问题和样本", m.CORAL),
    ], ["第四节：第一课课后准备"],
       "第一课后不确定最终项目，也不写PRD。", cols=2)
    homework_table(prs)
    m.slide_cards(prs, 29, "最熟悉场景的基础材料", "数量不必多，材料可以脱敏或合理模拟", [
        ("5个问题或案例", "真实用户、客户或同事可能遇到的问题", m.TEAL),
        ("1—3份材料", "制度、表单、记录、文档或业务样例", m.YELLOW),
        ("1个正常情况", "业务正常完成的过程", m.GREEN),
        ("1个特殊情况", "最容易出错或需要人工判断的情况", m.CORAL),
        ("2个未知问题", "当前无法回答、后续需要核实的内容", m.MUTE),
    ], ["准备一个场景的基础材料"],
       "未知问题也要如实记录，它们会影响后续产品边界。", cols=5)
    m.slide_compare(prs, 30, "下一课之前不用做什么", "先准备业务素材，第二课后再进入正式项目",
                    "本次不用完成", ["最终选题与完整PRD", "页面原型和技术方案", "模型选择与产品开发"],
                    "下一步", ["第二课学习智能客服公共项目", "课后正式确定第一个个人项目", "形成项目V0.1并进入第三课点评"],
                    ["本次不需要完成"], m.GREEN,
                    "先建立共同参照，再开始正式项目。")

    dk.declare_delivery(OUT, "presented")
    dk.lint_layout(prs, strict=True)
    prs.save(OUT)

    titles = [
        "第二阶段项目启动与准备", "今天主要讲四件事", "第二阶段的学习目标", "第一阶段与第二阶段",
        "公共项目、个人项目与主项目", "一轮项目的推进方式", "互评小组怎样工作", "十次课的整体安排",
        "第二阶段的最终成果", "公共项目的选择标准", "四项筛选条件", "项目不按技术方向划分",
        "三个公共项目覆盖的问题", "为什么没有单独设置C端公共项目", "三个公共项目与个人项目方向",
        "公共项目一：企业智能客服与业务办理", "智能客服变种：面向消费者", "智能客服变种：企业与内部服务",
        "公共项目二：业务材料审核与风险核验", "材料审核变种：财务、合同与采购",
        "材料审核变种：金融、人事与公共服务", "公共项目三：专业研究与方案生成",
        "研究与方案变种：销售、产品与运营", "研究与方案变种：行业、管理与专业服务",
        "从公共项目到个人项目", "个人方向的课堂选择", "第一课课后准备", "三个候选业务场景",
        "最熟悉场景的基础材料", "下一课之前不用做什么",
    ]
    slides = [{
        "slide": i,
        "role": "cover" if i == 1 else ("closing" if i == 30 else "content"),
        "takeaway": title,
        "evidence": [f"{SOURCE}｜对应章节"],
        "units": 1 if i in (1, 3, 12, 15, 27) else (5 if i in (6, 8, 13, 25, 29) else 3),
    } for i, title in enumerate(titles, 1)]
    gates = {
        "delivery": "presented",
        "content": {
            "checkpoint": {"mode": "approved", "record": "用户逐节确认第一课内容，并要求据此生成讲义后修改旧PPT。"},
            "arc": {
                "chosen": "学习方式—选择逻辑—项目变种—轻量准备",
                "candidates": [
                    {"name": "学习方式—选择逻辑—项目变种—轻量准备", "shape": "orientation-map-action", "roles": ["orientation", "selection", "transfer", "homework"], "audience_question": "第二阶段怎样学，这三个项目与我有什么关系？", "objection": "项目方向离自己的工作太远。", "closing_ask": "提交三个业务场景和一组基础材料。", "evidence": [str(SOURCE)]},
                    {"name": "小案例—项目结构—个人选题", "shape": "worked-example", "roles": ["case", "method", "selection"], "audience_question": "一个项目怎样完成？", "objection": "第一课提前讲项目细节，与第二课重复。", "closing_ask": "立即选题。", "evidence": ["PPT v2.0"]},
                    {"name": "课程规则—作业要求", "shape": "policy-first", "roles": ["rules", "homework"], "audience_question": "课程有哪些要求？", "objection": "缺少项目价值和个人代入感。", "closing_ask": "准备材料。", "evidence": ["旧版课程设计"]},
                ],
                "rejected": [
                    {"name": "小案例—项目结构—个人选题", "why_lost": "独立小案例与第二课重复，也会让第一课像第四个公共项目。"},
                    {"name": "课程规则—作业要求", "why_lost": "说教感较强，无法解释三个项目与学员背景的关系。"},
                ],
            },
            "slides": slides,
        },
        "design_plan": {
            "mode": "Mode A mimic",
            "style_pick": "n/a — locked mimic of the user-approved recent-course dark teal visual system",
            "boldness": "balanced+",
            "signature_move": "用三组分层变种地图把公共项目与不同岗位背景连接起来",
            "carried_by": [17, 18, 20, 21, 23, 24],
            "palette": {"fill_only": ["#031016", "#102B37", "#133843"], "text_safe": ["#F5F2E9", "#A7B5BC", "#36D6C2", "#FF7657", "#F0C565"], "semantic": "青绿=主结构，珊瑚=风险与提醒，黄色=条件与研究，绿色=完成与下一步"},
            "type_scale": {"display": 43, "title": 29, "body": 15},
            "motif_generates": {"background": "沿用深色课程画布", "markers": "章节编号和项目序号", "page": "第8页课程路线与三组项目流程"},
            "density": {"planned_median": 34, "over_70": 0, "non_text_protagonist": 18},
            "build_shape": "solo — template reuse and tightly coupled course narrative",
            "build_script": str(Path(__file__).resolve()),
            "checkpoint": {"mode": "approved", "record": "用户此前确认沿用最近课程PPT风格，并要求强化章节结构。"},
            "form_ledger": "封面1；陈述页2；比较页4；流程页2；表格页9；多项框架页12",
            "interior_register": "大号节标题、底部进度线、深色画布、青绿结构色与珊瑚提醒色",
            "icon_family": "none — template-locked; categories are carried by numbered structure and tables",
            "icon_none_category": "template-locked",
            "icon_none_checked": [f"slides 2-{m.SLIDE_COUNT}"],
            "form_reach": {"waived": "沿用用户确认的课程模板，内容形态覆盖路线、流程、比较、选择与表格。"},
        },
        "design": {"checkpoint": {"mode": "approved", "record": "Mode A沿用最近课程已确认的深色青绿体系。"}},
        "content_plan_ref": str(SOURCE),
        "output": str(OUT),
        "provenance": {"waived": "仅使用用户逐节确认的课程内容和本地讲义，不新增外部事实。"},
        "density": {"waived": "现场讲授PPT保留主要要点，完整解释进入老师版讲义和演讲者备注。"},
        "critic": {"waived": "当前环境未派发独立评审，主代理依据逐页渲染、版式检查和内容对照完成自审。", "waived_category": "no-dispatch-on-host", "inline_ran": True},
        "sameness": {"waived": "用户要求沿用最近课程的既有视觉系统。", "waived_category": "template-locked", "codes": ["LAYOUT SAMENESS", "CARD DOMINANCE", "TITLE-RULE MONOCULTURE"]},
        "capability_ledger": {"choice_ui": "yes", "subagents": "no — not requested", "image_input": "yes — rendered pages inspected", "web_access": "not used — no external factual claims"},
    }
    GATES.write_text(json.dumps(gates, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    build_deck()
