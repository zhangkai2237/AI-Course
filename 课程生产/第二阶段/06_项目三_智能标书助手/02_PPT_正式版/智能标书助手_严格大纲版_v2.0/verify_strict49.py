#!/usr/bin/env python3
"""Verify the built deck preserves the approved 49-page outline."""
from pathlib import Path

from pptx import Presentation


ROOT = Path(__file__).resolve().parent
PPTX = ROOT / "智能标书助手_严格按照49页大纲_正式授课PPT_v2.0.pptx"

EXPECTED = [
    ("智能标书助手",),
    ("核心链路：从RFP到可提交标书",),
    ("第一节", "为什么需要智能标书系统"),
    ("案例背景",),
    ("招标文件关键要求",),
    ("启明数科已有企业材料与事实",),
    ("AI生成的标书初稿",),
    ("课堂互动：找出这段初稿的问题",),
    ("参考答案：问题出在哪里",),
    ("通用AI写作工具 vs 智能标书系统",),
    ("第二节", "RFP需求拆解"),
    ("RFP的典型结构",),
    ("八类要求抽取框架",),
    ("摘要 vs 要求清单",),
    ("复合要求拆分",),
    ("要求卡片字段",),
    ("课堂练习：从原文中抽取要求",),
    ("参考答案：六项独立要求",),
    ("第三节", "响应矩阵与参标判断"),
    ("响应矩阵的作用",),
    ("五种响应状态",),
    ("资格与技术响应矩阵",),
    ("交付与服务响应矩阵",),
    ("参标初步判断",),
    ("课堂练习：判断响应状态",),
    ("参考答案",),
    ("第四节", "受控长文档生成"),
    ("一次生成整篇标书的问题",),
    ("两级标书树",),
    ("评分标准决定目录深度",),
    ("要求与章节映射",),
    ("章节生成任务包",),
    ("企业材料使用规则",),
    ("受控生成示例",),
    ("第五节", "一致性与风险检查"),
    ("七类检查",),
    ("跨章节一致性",),
    ("过度承诺检查",),
    ("课堂练习：找出方案中的风险",),
    ("规则、AI与人工的分工",),
    ("问题处理流程",),
    ("第六节", "MVP产品设计"),
    ("MVP产品流程",),
    ("页面1–2：项目主页与要求解析",),
    ("页面3–5：响应矩阵、章节编辑、风险检查与导出",),
    ("课堂MVP验收标准",),
    ("第七节", "场景迁移与课后作业"),
    ("三个迁移方向",),
    ("个人项目V0.1作业与结课",),
]


def slide_text(slide):
    return "\n".join(
        shape.text for shape in slide.shapes
        if getattr(shape, "has_text_frame", False) and shape.text
    )


def main():
    prs = Presentation(PPTX)
    assert len(prs.slides) == 49, f"expected 49 slides, got {len(prs.slides)}"
    failures = []
    for page, (slide, required) in enumerate(zip(prs.slides, EXPECTED), start=1):
        body = slide_text(slide)
        missing = [item for item in required if item not in body]
        if missing:
            failures.append((page, missing))
    if failures:
        raise SystemExit(f"outline mismatch: {failures}")
    print("STRICT49 PASS — 49/49 pages contain the approved title/task labels")


if __name__ == "__main__":
    main()
