export const SOURCE_POLICY = Object.freeze({
  version: "2.0",
  soleContentSource: "../07_专题课_便宜的答案昂贵的判断_前五节老师讲义_v1.0.md",
  forbiddenInputs: [".ppt", ".pptx", "旧版目录", "既有专题课PPT"],
  publicCaseLabel: "公共模拟案例",
  realProjectRule: "涉及学员本人项目时，只提供条件式表达模板；不得虚构上线、用户、订单、指标、结果或个人经历。",
  conflictCorrection: "讲义中凡把课堂演示写成真实上线、真实用户或真实结果的表述，统一改为条件式：若你的项目已发生，则展示证据；若尚未发生，则说明验证计划与待补证据。",
  evidenceRules: [
    "事实与推测分开",
    "有来源才陈述为事实",
    "无数据就写待验证，不编数字",
    "公共客服案例始终标注公共模拟案例",
    "真实项目只用本人能够举证的材料",
  ],
});

export function assertSourcePath(sourcePath) {
  const lower = sourcePath.toLowerCase();
  if (lower.endsWith(".ppt") || lower.endsWith(".pptx")) {
    throw new Error(`Forbidden legacy presentation input: ${sourcePath}`);
  }
  if (!lower.endsWith(".md")) throw new Error(`Only Markdown source is allowed: ${sourcePath}`);
}
