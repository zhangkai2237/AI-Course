import fs from "node:fs/promises";

const sourcePath = new URL("./build_lesson9_core.mjs", import.meta.url);
const outPath = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_课堂讲义_核心结构版.md";

const source = await fs.readFile(sourcePath, "utf8");
const start = source.indexOf("const S=[];");
const end = source.indexOf("\nfunction render(", start);
if (start < 0 || end < 0) throw new Error("无法从PPT生成脚本中读取页面结构");

const block = source.slice(start, end);
const S = Function(`${block}\nreturn S;`)();

const esc = (value) => String(value ?? "").replaceAll("|", "\\|").replaceAll("\n", "<br>");
const bullets = (items) => items.map((item) => `- ${item}`).join("\n");
const numbered = (items) => items.map((item, index) => `${index + 1}. ${String(item).replaceAll("\n", "：")}`).join("\n");
const table = (headers, rows) => {
  const head = `| ${headers.map(esc).join(" | ")} |`;
  const split = `| ${headers.map(() => "---").join(" | ")} |`;
  const body = rows.map((row) => `| ${row.map(esc).join(" | ")} |`).join("\n");
  return `${head}\n${split}\n${body}`;
};

const deepNotes = {
  "AI测试与传统软件测试": "传统软件测试通常可以把输入、操作步骤和预期结果写得非常明确，例如点击按钮后页面必须跳转到指定地址。AI产品不同：用户可能用很多种表达提出同一个需求，模型也可能用不同句子给出同样正确的答案。因此，AI评测不能只做字符串匹配，而要判断意思是否正确、证据是否可靠、风险是否可控。产品经理需要从“功能有没有执行”进一步走向“回答质量是否达到业务要求”。",
  "AI评测的三个核心难点": "这三个不确定性是后续所有评测设计的起点。输入不确定意味着评测集必须覆盖表达变体；输出不确定意味着评分标准不能只依赖唯一答案；原因不确定意味着发现错误后不能立刻把责任归给模型，而要检查知识、检索、Prompt、工作流和工具调用。三个难点分别对应评测集、评分器和Bad Case归因。",
  "上线前评测的作用": "线上反馈当然重要，但它属于问题已经发生之后的证据。对于隐私泄露、违规承诺、危险建议等错误，不能用真实用户来替我们完成测试。上线前评测的价值，是提前把业务规则、高风险边界和历史事故转成测试题，用可控成本发现问题，并据此决定是否允许发布。",
  "住宿报销回答怎么评分": "这个案例要让学员看到：一句看起来正确的回答，也可能只是“部分正确”。500元/晚的金额正确，解决的是正确性；是否说明普通员工或不同职级，关系到完整性；金额来自哪一版制度，关系到忠实性；有没有越权披露或诱导违规，则属于安全性。评测不是只问“对不对”，而是拆成多个可以判断的维度。",
  "指标门槛的三层结构": "三层结构不能互相替代。一票否决项是底线，只要发生一次就可能阻止上线；上线门槛用于判断核心能力是否已经达到可用水平；优化指标是在产品可用之后继续改善体验。最常见的错误，是用总体平均分掩盖红线问题，或者产品还没有达到核心正确率，就先花大量时间优化文风。",
  "评测样本的建议比例": "比例不是固定行业标准，而是一种资源分配思路。高频核心场景决定大多数用户的日常体验，因此占比最高；历史Bad Case已经证明产品容易出错，必须持续保留；边界、异常、高风险和无答案样本虽然频率低，却决定产品是否稳健。实际项目中应根据业务风险调整比例，而不是机械照抄数字。",
  "规则评测：检查明确条件": "规则评测适合判断可枚举、可计算、可精确验证的内容。这里可以直接检查JSON字段是否存在、类型是否正确、城市是否为北京、金额是否等于500。它速度快、成本低、结果稳定，也方便定位错误。但规则只能判断写进规则里的条件，无法自动理解一段开放解释是否真正帮助了用户。",
  "语义相似度的局限": "参考答案和被测答案大部分词语相同，所以向量相似度可能很高，但3至5天与5至7天在业务上是明显不同的承诺。这个例子用来说明：语义相似度擅长识别“是不是在谈同一件事”，不擅长保证数字、日期、否定词和责任边界完全正确。关键事实仍应交给规则、Judge或人工复核。",
  "LLM-as-Judge需要哪些输入": "Judge并不是把两个答案随手丢给另一个大模型。它至少需要知道用户问了什么、回答时有哪些必要上下文、理想答案是什么、要从哪些维度评分，以及实际被测答案。缺少上下文时，Judge可能把合理回答判错；缺少评分维度时，它又容易凭整体印象打分。Judge质量首先取决于评测任务是否定义清楚。",
  "Judge Prompt的五个组成": "设计Judge Prompt时，应像设计一份阅卷标准。先明确裁判任务，再定义评分维度；为每个分数提供锚点，避免“好、一般、差”过于模糊；写清楚一票否决的风险；最后要求输出结构化结果和证据。结构化输出便于自动统计，证据字段则方便人工复核Judge为什么这样打分。",
  "Judge输出示例": "这个JSON不仅给出总的好坏判断，而是分别记录正确性、相关性、完整性和安全性。示例中金额正确，所以正确性得2分；回答确实回应了问题，相关性得2分；但遗漏职级差异，完整性只有1分。reason字段必须指向回答中的具体缺失，不能只写“回答不够好”，否则无法指导后续修复。",
  "住宿案例的组合评测": "组合评测的思想是让不同工具各做自己擅长的工作。金额和单位是硬条件，用规则最稳定；是否真正回答问题、是否遗漏职级属于开放质量判断，可以交给Judge；涉及高风险、争议或制度边界的样本再由人工抽查。不是评测方法越高级越好，而是组合之后既可靠，又能控制成本。",
  "自动测试Agent是什么": "自动测试Agent通常不是面向所有业务的万能Agent，而是围绕一套评测协议搭建的执行系统。底层框架可以通用，例如读取用例、调用产品、保存结果和生成报告；但每个产品的用例、评分规则、红线和版本信息必须定制。可以把它理解成通用评测引擎加业务专用配置。",
  "总体通过率与红线风险": "99%的总体通过率看起来很高，但如果剩下的1%是隐私泄露，这个版本仍然不能上线。因此报告不能只给一个平均数，必须拆分场景、风险等级和失败类型。产品决策关心的不是“平均上做得不错”，而是核心任务是否达到门槛、红线是否为零、与上一版本相比是否出现退化。",
  "上线门槛示例": "门槛应在测试前确定，不能看到结果后再临时修改。红线必须为零代表底线；核心事实正确率和高风险通过率代表产品是否可用；响应时间和成本则保证方案能在真实业务里运行。具体阈值要结合业务风险、历史基线和交付目标设定，示例中的数字用于演示门槛结构。",
  "Bad Case分析的五步流程": "分析Bad Case时，第一步不是改Prompt，而是复现。只有确认相同输入、上下文和版本下确实失败，后续分析才有意义。然后明确正确预期，定位错误发生在哪个环节，再做根因归类，最后设计修复和回归验证。跳过任何一步，都可能出现修错问题或只修表象的情况。",
  "住宿标准错误的完整归因": "回答600元不等于模型单点出错。检查链路后可能发现，知识库同时存在新旧制度，检索把旧版排在前面；系统又没有询问用户职级；输出前也没有版本校验。因此主要原因是知识优先级，次要原因是流程缺失，最后还有一道防线没有建立。完整归因会自然导出多个修复动作，而不是只在Prompt里加一句“不要答错”。",
  "Bad Case处理优先级": "频率高的问题影响用户多，影响程度高的问题后果严重，可修复性决定投入是否能快速产生回报。三个因素共同帮助团队排优先级，但安全与合规红线通常不应仅按乘法计算，而应直接提升到最高优先级。这个公式是决策框架，不是绝对精确的数学模型。",
  "回归测试的两个目标": "修复验证只回答“原来的问题好了没有”，防退化验证则回答“修这个问题时有没有把别的能力弄坏”。例如为了避免报销金额答错，模型可能被改得过度保守，之后所有报销问题都拒绝回答。一个合格的回归集必须同时包含原失败样本、相邻变体和原本正常的代表性样本。",
  "稳定性评测方法": "生成式模型具有随机性，一次答对不能证明稳定。重要问题应重复运行，并记录通过率、最差结果和波动范围。连续5次中有1次失败，说明产品仍存在真实风险，尤其在高风险场景中不能只展示平均表现。重复次数应根据成本、随机性和风险等级设置。",
  "版本对比报告看什么": "版本对比不是只看正确率有没有提升。V2修复了原Bad Case，核心正确率也上升，但新增了一次安全红线，同时响应更慢、成本超标。这样的版本不能因为平均能力更强就直接全量发布。报告的作用，是把效果、安全、效率和成本放在同一张决策表里。",
  "低分会话的处理流程": "线上低分会话不能未经处理直接进入训练或评测。首先要做隐私脱敏、去重和上下文补全，再由人工确认用户真实意图和正确预期，之后进行归因和严重度分级。只有稳定、代表性强、标注可信的案例才进入回归集，其余案例可以留在问题池中等待进一步确认。",
  "数据清洗与数据标注": "清洗解决的是数据能不能安全、完整地使用，例如删除隐私、合并重复、补齐上下文和排除恶意输入；标注解决的是团队如何理解问题，例如问题类型、严重度、期望行为和责任模块。两部分都可以让大模型辅助，但抽样质检、规则制定和高风险判断仍需要人工负责。",
  "标注规范需要正例和反例": "只有定义往往不够，因为不同标注员会对边界产生不同理解。正例告诉大家什么情况应该标为知识缺失；反例说明虽然现象相似，但根因并不属于这一类。示例中，知识库没有2026年制度属于知识缺失；如果资料已经召回但模型没有引用，则更接近模型使用知识或Prompt问题。",
  "从问题标注到产品实验": "统计发现知识召回错误占35%，只说明主要矛盾在哪里，不能直接证明某个方案一定有效。团队需要进一步形成假设，例如旧文档排序过高，再修改检索权重或版本过滤，然后用离线评测和灰度实验验证。如果错误下降但延迟大幅上升，也要重新权衡。",
  "数据飞轮": "飞轮成立的前提不是数据越多越好，而是数据能够被正确处理。更多使用带来更多真实问题，团队识别新的失败类型，把它们转成评测样本和产品改进，再通过回归和线上观察确认效果。只有经过清洗、标注、验证的数据回流，才会让产品越用越好；未经筛选的数据可能放大噪声。"
};

function autoTeaching(slide) {
  if (deepNotes[slide.title]) return deepNotes[slide.title];
  if (slide.type === "cover") return "开场先告诉学员：这节课不讨论如何让模型回答得更像人，而是讨论如何证明一个AI产品已经达到可用、可上线、可持续改进的状态。整节课会从评测标准出发，经过自动化评测、Bad Case和回归测试，最后进入真实数据反馈闭环。";
  if (slide.type === "section") return `这一部分围绕“${slide.subtitle}”展开。讲解时先帮助学员建立本模块的问题意识，再介绍方法和案例，不需要一开始就陷入工具或技术细节。`;
  if (["map", "flow", "closing", "loop"].includes(slide.type)) return `这页用连接关系呈现${slide.items.length}个环节或组成部分。讲解时按PPT顺序推进：${slide.items.map((item, i) => `第${i + 1}项“${String(item).replaceAll("\n", "：")}”`).join("；")}。如果页面表达的是流程，要说明前后步骤的输入输出；如果表达的是结构，则要说明各部分的分工以及为什么需要同时存在。`;
  if (slide.type === "cards") return `这页把主题拆成${slide.items.length}个观察角度。${slide.items.map((item) => `“${item.k}”关注${item.v}`).join("；")}。讲解时可以先给出整体框架，再用当前产品中的一个真实问题逐项套入，帮助学员理解这些概念不是彼此孤立的。`;
  if (slide.type === "compare") return `这页需要通过对比建立边界。左侧“${slide.left.h}”与右侧“${slide.right.h}”不是简单的好坏关系，而是适用于不同判断对象。讲完各自特点后，要落到产品决策：当前问题更接近哪一侧，应该采用哪套方法，以及是否需要组合使用。`;
  if (slide.type === "matrix") return `这张表用于把抽象方法落到具体对象。讲解时不要逐字念表，而是按行说明“对象是什么、判断依据是什么、最后影响什么决策”。可以邀请学员选择其中一行，用自己的业务场景替换示例，检查这套结构是否仍然成立。`;
  if (slide.type === "list") return `这页的重点是解释这些条件为什么会影响产品结果，而不只是记住条目。可以逐条联系上线决策：如果忽略这一点，会出现什么错误；如果提前处理，需要增加什么评测或流程。最后用页面中的结论句收束。`;
  if (slide.type === "case") return `案例页建议先遮住评测结论，只展示问题和AI回答，让学员先判断“这个答案能不能上线”。随后再逐项展开检查点。这样可以让学员体会到，直觉上的“看起来没问题”与可执行的评测标准之间存在差距。`;
  if (slide.type === "bars") return `柱状图表达的是相对优先级和资源分配。讲解时先看最高项为什么最高，再比较其余项目承担的风险。数字是当前课程中的示例配置，实际项目要根据业务频率、错误影响和历史数据重新校准。`;
  if (slide.type === "code") return "代码示例的重点不是语法，而是结构化结果。逐个解释字段分别判断什么、分值如何解释、reason为什么必须提供证据。结构化输出可以进入结果库做统计，理由则支持人工抽检和问题归因。";
  if (slide.type === "gate") return "从上到下讲清楚底线、可用和优化三个层级。上层门槛没有通过时，不应被下层的体验优势抵消；同样，达到底线也不代表产品已经足够好。发布决策必须同时满足对应层级的要求。";
  if (slide.type === "formula") return `公式“${slide.formula}”用于把团队争论转成相对一致的优先级判断。每个因子都需要基于真实业务证据估计，并对安全、合规等特殊问题设置额外规则，避免机械套用公式。`;
  if (slide.type === "summary") return `这一页用于阶段收束。按顺序回顾${slide.items.length}个产物或结论，并让学员确认：如果缺少其中任何一项，后续流程会在哪一步失去依据。讲完后自然过渡到下一模块。`;
  return "围绕页面中的概念说明它解决什么问题、依赖什么输入、最后影响什么产品决策，并结合当前课程的住宿报销案例帮助学员理解。";
}

function interaction(slide) {
  if (slide.type === "case") return "先不要公布评分，让学员用30秒判断：这个回答可以直接上线吗？请说出至少一个判断依据。";
  if (slide.type === "compare") return `请学员分别举出一个符合“${slide.left.h}”和“${slide.right.h}”的例子，并讨论两侧内容在真实项目中是否需要组合使用。`;
  if (["flow", "map", "loop"].includes(slide.type)) return "请学员指出：如果这条链路只允许优先改一个环节，应该选哪里？还需要什么证据才能确认？";
  if (slide.type === "matrix") return "选择表格中的一个场景，请学员补充一个真实问题，并说明它应该如何评测。";
  if (slide.type === "bars") return "请学员思考：如果产品属于金融、医疗或企业权限场景，这组比例或优先级需要怎样调整？";
  if (slide.type === "section") return "进入本部分前，可以请学员用一句话说出自己最想解决的问题，讲完本部分后再回看是否得到答案。";
  return "可以请学员用自己的产品举一个例子，判断这一页的方法在真实业务中需要哪些调整。";
}

function transitionText(slide, next) {
  if (!next) return "“到这里，第9课的完整闭环已经建立。最后请大家回到自己的产品，写下下一轮最需要补齐的评测标准、Bad Case或数据反馈。”";
  if (next.type === "section") return `“关于${slide.title}，我们先总结到这里。接下来进入新的模块：${next.title}。”`;
  if (slide.type === "section") return `“这一部分要解决的问题已经明确，下面先看${next.title}。”`;
  return `“理解了${slide.title}以后，下一步来看${next.title}。”`;
}

function renderSlide(slide, index) {
  const page = String(index + 1).padStart(2, "0");
  const lines = [`## 第${page}页｜${slide.title}`];
  if (slide.phase) lines.push(`\n> 模块：${slide.phase}`);
  if (slide.subtitle) lines.push(`\n${slide.subtitle}`);

  if (slide.type === "cover") {
    lines.push("\n**课程关键词**：评测 · Bad Case · 回归 · 数据反馈闭环");
  } else if (slide.type === "section") {
    lines.push("\n这一部分进入新的课程模块。先向学员说明本模块要解决的问题，再进入具体方法与案例。");
  } else if (["map", "flow", "closing"].includes(slide.type)) {
    lines.push(`\n${numbered(slide.items)}`);
    lines.push(`\n**讲解顺序**：${slide.items.map((item) => String(item).replaceAll("\n", "：")).join(" → ")}`);
  } else if (slide.type === "cards") {
    for (const item of slide.items) lines.push(`\n### ${item.k}\n\n${item.v}`);
  } else if (slide.type === "compare") {
    const length = Math.max(slide.left.items.length, slide.right.items.length);
    const rows = Array.from({ length }, (_, i) => [slide.left.items[i] ?? "", slide.right.items[i] ?? ""]);
    lines.push(`\n${table([slide.left.h, slide.right.h], rows)}`);
  } else if (slide.type === "matrix") {
    lines.push(`\n${table(slide.headers, slide.rows)}`);
  } else if (slide.type === "list") {
    lines.push(`\n${bullets(slide.items)}`);
    if (slide.quote) lines.push(`\n> ${slide.quote}`);
  } else if (slide.type === "case") {
    lines.push(`\n### 用户问题 / 评测输入\n\n${slide.q}`);
    lines.push(`\n### AI回答 / 被测结果\n\n${slide.a}`);
    if (slide.points?.length) lines.push(`\n${table(["检查项", "判断"], slide.points)}`);
  } else if (slide.type === "bars") {
    lines.push(`\n${table(["项目", "数值"], slide.items.map((item) => [item.k, `${item.v}${item.s ?? "%"}`]))}`);
  } else if (slide.type === "code") {
    lines.push(`\n\`\`\`json\n${slide.code}\n\`\`\``);
  } else if (slide.type === "gate") {
    for (const layer of slide.layers) lines.push(`\n### ${layer.k}\n\n${layer.v}`);
  } else if (slide.type === "formula") {
    lines.push(`\n> **${slide.formula}**`);
    lines.push(`\n${bullets(slide.items)}`);
  } else if (slide.type === "summary") {
    lines.push(`\n${numbered(slide.items)}`);
  } else if (slide.type === "loop") {
    lines.push(`\n${numbered(slide.items)}`);
    lines.push(`\n**循环关系**：${slide.items.join(" → ")}`);
  }

  lines.push(`\n### 讲解稿\n\n${autoTeaching(slide)}`);
  lines.push(`\n### 课堂互动\n\n${interaction(slide)}`);
  lines.push(`\n### 过渡语\n\n${transitionText(slide, S[index + 1])}`);

  return lines.join("\n");
}

const toc = S.map((slide, index) => `- 第${String(index + 1).padStart(2, "0")}页：${slide.title}`).join("\n");
const body = S.map(renderSlide).join("\n\n---\n\n");
const markdown = `# 第9课：AI产品评测与数据反馈｜课堂讲义\n\n` +
  `本讲义依据原始课程MD扩展，并与《第09课_正式PPT_核心结构版》保持相同的67页顺序、标题和内容结构。每页包含PPT呈现内容、详细讲解稿、课堂互动和过渡语，可直接按页码配合PPT授课。\n\n` +
  `## 课程主线\n\n` +
  `评测准备 → 评分方法 → 自动化评测 → Bad Case与回归 → 数据反馈闭环\n\n` +
  `<details>\n<summary>展开查看完整页码目录</summary>\n\n${toc}\n\n</details>\n\n---\n\n${body}\n`;

await fs.writeFile(outPath, markdown, "utf8");
console.log(JSON.stringify({ slides: S.length, outPath }));
