import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const inputPptx = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson10-narrative-revision/template-starter.pptx";
const outputPptx = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第10课_AI产品方案与PRD设计实战_重做工作区/第10课_正式PPT_教学主线修订版.pptx";
const renderDir = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson10-narrative-revision/final-render";
const layoutDir = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson10-narrative-revision/final-layout";

const edits = {
  2: ["sh/ozy1ofad", "AI PRODUCT · OVERVIEW", "课程衔接", "sh/x4r21kru", "第9课会评测，第10课把评测要求前置到产品方案", "课程衔接：为什么评测之后要学产品方案"],
  3: ["sh/d0jax03i", "AI PRODUCT · OVERVIEW", "课程主线", "sh/obq90bml", "第10课的产品设计主线", "项目主线：把模糊AI想法变成可验收PRD"],
  4: ["sh/0ba143al", "AI PRODUCT · OVERVIEW", "课程目标", "sh/doj29oba", "学习目标", "学习目标：完成一份AI产品方案"],
  5: ["sh/cf2tcr61", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/l4bupwny", "01 业务背景与产品目标", "01 业务背景与产品目标"],
  6: ["sh/zi98nu94", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/ml07i9sv", "项目从内容生产提速开始", "项目起点：为什么需要智能审核"],
  7: ["sh/dcbud0ra", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/kzmdova1", "人工审核流程出现五类问题", "业务现状：人工审核遇到哪些问题"],
  8: ["sh/g72x4zyd", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/tkby9kzm", "生产效率提升，审核成为新瓶颈", "核心矛盾：生产提速后，瓶颈转移到审核"],
  9: ["sh/0b65obm9", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/9gzml0nq", "审核不一致与规则变化同时存在", "业务难点：审核标准为什么难以统一"],
  10: ["sh/cbu58j2h", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/l036l83y", "系统既要沉淀数据，也要保留人工责任", "产品约束：为什么必须保留人工责任"],
  11: ["sh/wbih4b6d", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/98rytw72", "先补业务基线，再验证产品假设", "现状分析：立项前需要补哪些业务基线"],
  12: ["sh/0f2lgnmp", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/dcbm583y", "不要把业务需求直接写成产品需求", "需求澄清：把业务诉求转成产品问题"],
  13: ["sh/bq9orito", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/21gnuts7", "项目中有四类核心用户", "用户分析：谁会使用和管理智能审核"],
  14: ["sh/wbydknq1", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/5grehs7i", "产品目标要能被业务和评测共同验证", "目标定义：产品效果要如何验证"],
  15: ["sh/k3y5ov21", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/t87ml0ji", "第一版同时写清目标与非目标", "范围约束：第一版做什么、不做什么"],
  16: ["sh/9kby1g7m", "AI PRODUCT · BUSINESS", "项目第1步 · 业务问题", "sh/kv2h4rqp", "第一版项目定义", "阶段产出：完成第一版项目定义"],
  17: ["sh/gbedwfmx", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/t8ne10n6", "02 AI介入判断与能力分工", "02 AI介入判断与能力分工"],
  18: ["sh/hsn2l4bu", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/4felwzu5", "智能审核至少有四种候选方案", "方案选择：智能审核有哪些实现路线"],
  19: ["sh/8v2pobax", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/xkbq10be", "四种方案的取舍", "方案比较：四种路线分别如何取舍"],
  20: ["sh/rq50vmp8", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/2lczih8r", "AI是否介入，需要回答六个产品问题", "AI介入判断：什么情况下值得使用AI"],
  21: ["sh/0f2lon6p", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/dcb2t87y", "只使用大模型，上线时会暴露六类问题", "单一模型风险：为什么不能只接大模型"],
  22: ["sh/7a18rydc", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/il87etcv", "先按任务性质分配能力", "分工原则：先按任务性质选择技术方案"],
  23: ["sh/8zm9k7y5", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/xovqhcz2", "确定性规则负责稳定执行", "能力分工①：哪些任务交给规则"],
  24: ["sh/0r6p4nul", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/9wzq1svi", "大模型负责规则难以穷举的语义", "能力分工②：哪些任务交给大模型"],
  25: ["sh/cfyt4beh", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/pc7atgf6", "RAG提供依据，人工承担责任", "能力分工③：RAG和人工分别做什么"],
  26: ["sh/vitkzqlc", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/md0321kv", "混合审核的完整协作流程", "方案落地：四类能力怎样组成审核流程"],
  27: ["sh/lofadsni", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/8r698nm9", "第一版用Workflow，而不是Agent", "流程选择：第一版为什么使用Workflow"],
  28: ["sh/ralgf21g", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/exwzahkr", "能力分工要写到输入、输出和边界", "需求写法：能力分工要写清哪些边界"],
  29: ["sh/8jup8rad", "AI PRODUCT · SOLUTION", "项目第2步 · 方案与分工", "sh/h8n6lwbu", "同一条文案如何被分层审核", "贯穿案例：同一条文案如何被分层审核"],
  30: ["sh/o3yl4361", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/xsr2h8ni", "03 结果动作与人工复核", "03 结果动作与人工复核"],
  31: ["sh/cfa903yt", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/1k3ad8fa", "先把规则、RAG和模型的原始结果放在一起", "结果合并：三类能力分别发现了什么"],
  32: ["sh/lszilkfm", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/8vq1wzyd", "结果合并解决三个问题", "结果合并：为什么要去重、留证据、给建议"],
  33: ["sh/0jalwfm5", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/9o3m9knm", "一条正式风险结果必须结构化", "结果结构：一条风险结果包含哪些字段"],
  34: ["sh/ovq58fu1", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/d0zm54bi", "一篇内容需要四类整体处理结果", "整体结果：一篇内容会进入哪种状态"],
  35: ["sh/3ex47ip8", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/u943ador", "风险等级必须映射到明确产品动作", "动作映射：风险等级如何决定下一步"],
  36: ["sh/4rut0fa5", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/dwnax4bm", "运营端不是只看结果，而是完成修改闭环", "用户流程：运营怎样完成修改闭环"],
  37: ["sh/7m58vyhw", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/ixc7it0f", "六类情况触发人工复核", "人工接管：哪些情况必须触发复核"],
  38: ["sh/lw7idgva", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/8zi1obe1", "人工复核工作台要减少重复劳动", "复核工作台：人工需要看到和处理什么"],
  39: ["sh/4ze50fu9", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/do76d4vq", "审核任务需要明确状态", "任务状态：审核流程如何记录进度"],
  40: ["sh/zapg3qh8", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/alwfql0r", "机器结果和人工结果必须同时保留", "结果追溯：为什么机器和人工结论都要保留"],
  41: ["sh/szi98fi9", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/h4bqlkj6", "内容修改和机器失败都不能偷走安全步骤", "异常处理：内容修改和机器失败如何接管"],
  42: ["sh/nmxwfetc", "AI PRODUCT · FLOW", "项目第3步 · 结果与流程", "sh/yh4vi9cv", "完整案例从V1走到V2", "贯穿案例：一条文案如何从V1走到V2"],
  43: ["sh/43axcvqx", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/h0jy1gr6", "04 MVP范围与能力优先级", "04 MVP范围与能力优先级"],
  44: ["sh/s3qdcr2t", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/1szepg3q", "MVP验证的是最小审核闭环", "MVP定义：第一版要跑通哪条最小闭环"],
  45: ["sh/0rq1kjed", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/9gj2hovu", "第一版要验证三个核心假设", "核心假设：MVP首先验证什么"],
  46: ["sh/va14nmtg", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/ids3yhc7", "PoC、MVP与完整产品回答不同问题", "阶段区分：PoC、MVP和完整产品有何不同"],
  47: ["sh/4zadgjap", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/do3edor6", "功能是否进入MVP，连续问四个问题", "范围判断：一个功能是否应该进入MVP"],
  48: ["sh/83m9snyx", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/l0vaxsf6", "第一版主动收窄试点范围", "试点范围：第一版主动收窄哪些维度"],
  49: ["sh/1sj2psbq", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/cnq1snu9", "MVP必须具备八类能力", "必备能力：最小闭环需要哪些产品能力"],
  50: ["sh/pgzax8vq", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/gbq90jet", "这些能力暂时不进入第一版", "延后能力：哪些需求暂不进入第一版"],
  51: ["sh/07yt0vmp", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/9craxk36", "能力优先级围绕闭环排序", "优先级：先保障闭环，再提升效率和规模"],
  52: ["sh/gnq94n61", "AI PRODUCT · MVP", "项目第4步 · MVP范围", "sh/psjahcny", "MVP采用小范围试点，而不是一次性全量上线", "试点方式：第一版如何小范围安全上线"],
  53: ["sh/natwz2po", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/elkfmx87", "05 大模型能力输入与输出", "05 大模型能力输入与输出"],
  54: ["sh/0vqxgja5", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/dszyl4ru", "一项AI能力要求至少包含八个部分", "能力契约：一项AI能力要写哪些部分"],
  55: ["sh/r2hszqpc", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/ep8ba583", "不要让模型判断“是否合法”", "任务定义：大模型究竟应该判断什么"],
  56: ["sh/hcbqlob2", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/4f29g3ud", "模型需要五类输入信息", "输入设计：模型需要获得哪些信息"],
  57: ["sh/gb250ja1", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/tobm5ora", "模型输入不是一句Prompt，而是一组受控上下文", "输入示例：怎样把业务上下文传给模型"],
  58: ["sh/8v2p43qt", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/lsb6to72", "输出结构必须能够被系统稳定解析", "输出示例：怎样让系统稳定解析结果"],
  59: ["sh/velsb618", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/ihwrm1kz", "模型提供判断材料，系统决定业务动作", "动作决策：模型结果如何转成业务动作"],
  60: ["sh/rudg7et0", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/ipkfa9cj", "修改建议和最终改写是两个任务", "任务拆分：修改建议和自动改写有何不同"],
  61: ["sh/8bqdczap", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/l8zu14be", "不确定与失败结果也必须有产品定义", "异常定义：不确定和失败结果如何处理"],
  62: ["sh/cjyh4v29", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/pw7y903i", "AI能力边界必须写进PRD", "能力边界：哪些责任不能交给大模型"],
  63: ["sh/83e9o7yx", "AI PRODUCT · CAPABILITY", "项目第5步 · 能力契约", "sh/lgnqtsf6", "AI能力卡：营销文案语义风险识别", "阶段产出：完成语义风险识别能力卡"],
  64: ["sh/gjehonyh", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/58nilsfy", "06 评测集来源与验收约定", "06 评测集来源与验收约定"],
  65: ["sh/kzel0z6h", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/xwn25kn6", "首版评测集来自五类真实与补充数据", "数据来源：首版评测集从哪里获得样本"],
  66: ["sh/ovet0va9", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/1sna50ri", "历史标签不能直接等于标准答案", "标签治理：历史标签为什么不能直接使用"],
  67: ["sh/orex83qt", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/1o3yxo7i", "评测表每一行都要包含输入、标准和版本", "表格结构：评测集每一行要记录什么"],
  68: ["sh/ryt4nu9o", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/i903ap87", "首版评测集的整理过程", "整理流程：如何从历史数据形成评测集"],
  69: ["sh/xcbalcza", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/kfmtwryl", "没有历史数据，也可以建立首版评测集", "冷启动：没有历史数据怎样建立评测集"],
  70: ["sh/c3id0zqt", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/p0ru5k72", "开发、验收和回归样本要分开管理", "数据管理：开发、验收和回归样本为何分开"],
  71: ["sh/kruxk3q1", "AI PRODUCT · EVALUATION", "项目第6步 · 评测验收", "sh/twnex87i", "PRD中的验收约定至少写清六项", "验收约定：PRD必须提前写清哪些问题"],
  72: ["sh/gbi98j21", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/t87ad43a", "07 组装完整AI产品PRD", "07 组装完整AI产品PRD"],
  73: ["sh/vm9svehs", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/i90r69gj", "AI产品PRD比普通PRD多写三类内容", "PRD差异：AI产品需求多写了哪些内容"],
  74: ["sh/43ulwry5", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/h0jmlwze", "最终PRD由十三个章节组成", "文档结构：完整AI产品PRD包含哪十三章"],
  75: ["sh/orixgbe1", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/1orelgfa", "PRD首页要在一页说清项目", "首页摘要：怎样用一页说清整个项目"],
  76: ["sh/ov2p0fmd", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/1srq5032", "前六节内容按同一顺序组装进PRD", "组装顺序：前六步如何进入同一份PRD"],
  77: ["sh/7md4vuxk", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/up4nq9gb", "数据、规则和知识依赖要单独说明", "运行依赖：数据、规则和知识库要说明什么"],
  78: ["sh/tojm9c3i", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/kzqlc721", "风险与待确认项不是缺陷，而是决策清单", "决策清单：风险与待确认项如何写进PRD"],
  79: ["sh/0725sjep", "AI PRODUCT · PRD", "项目第7步 · 组装PRD", "sh/9cbm5ov6", "课堂用两项交付检验是否真正掌握", "课堂交付：用哪两项练习检验掌握程度"],
  80: ["sh/crid47ed", "AI PRODUCT · OVERVIEW", "课程复盘 · 完整链路", "sh/1gbe1cfu", "第10课的完整产品设计链路", "课程复盘：一份AI产品方案怎样形成"],
  81: ["sh/dknid4jm", "AI PRODUCT · NEXT", "下一步 · 投入判断", "sh/4fu10z25", "PRD完成以后，还要回答：这个项目值得投入吗？", "下一课问题：方案可开发后，值得投入吗"]
};

const chapterSubtitles = [
  ["sh/n6dcr65o", "先把真实业务问题说清楚，再讨论AI", "先回答：为什么做、为谁做、要达到什么结果"],
  ["sh/jilcvq54", "先比较方案，再决定每种能力负责什么", "业务问题明确后，判断AI是否介入，以及各能力如何分工"],
  ["sh/ru5k3ep8", "发现问题以后，产品要让用户继续完成任务", "方案选定后，设计审核结果如何进入真实业务流程"],
  ["sh/fulgba9k", "第一版不是功能最少，而是验证链路完整", "流程闭环后，决定第一版做到哪里、暂时不做什么"],
  ["sh/cjixknq1", "把“大模型审核”写成可实现、可接入的能力契约", "MVP范围确定后，把模型审核写成可实现的能力契约"],
  ["sh/rap0n2x4", "PRD要提前说明拿什么测、谁确认、何时重测", "能力契约写清后，约定用什么数据、如何判断能否上线"],
  ["sh/b2tsnylo", "把前六节的产品判断装进一份可执行文档", "把前六步的产品判断组装成可开发、可验收的PRD"]
];

function replaceText(presentation, id, oldText, newText) {
  const target = presentation.resolve(id);
  target.text.replace(oldText, newText);
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(inputPptx));

  for (const [slideNumber, [eyebrowId, oldEyebrow, newEyebrow, titleId, oldTitle, newTitle]] of Object.entries(edits)) {
    replaceText(presentation, eyebrowId, oldEyebrow, newEyebrow);
    replaceText(presentation, titleId, oldTitle, newTitle);
  }

  for (const [id, oldText, newText] of chapterSubtitles) {
    replaceText(presentation, id, oldText, newText);
  }

  presentation.resolve("sh/apc32p47").text.style = { fontSize: 21 };
  presentation.resolve("sh/ih0369sn").text.style = { fontSize: 21 };
  presentation.resolve("sh/il47adov").text.style = { fontSize: 18.5 };

  replaceText(
    presentation,
    "sh/5gvixcbu",
    "人工复核、异常、风险与待确认项",
    "12 人工复核与异常处理\n13 风险与待确认项",
  );
  replaceText(
    presentation,
    "sh/m9wjat0j",
    "第11课：模型选择、调用成本、人工成本与业务收益",
    "",
  );
  replaceText(presentation, "sh/alcvmh4v", "组装完整PRD", "组装PRD");

  await fs.mkdir(renderDir, { recursive: true });
  await fs.mkdir(layoutDir, { recursive: true });
  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(`${renderDir}/${stem}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
    await fs.writeFile(`${layoutDir}/${stem}.layout.json`, await (await slide.export({ format: "layout" })).text());
  }
  await writeBlob(
    "/Users/keivn/Project/AI-Course/.codex-tmp/lesson10-narrative-revision/final-montage.webp",
    await presentation.export({ format: "webp", montage: true, scale: 1 }),
  );

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPptx);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
