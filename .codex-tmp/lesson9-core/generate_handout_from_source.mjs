import fs from "node:fs/promises";

const pptSourcePath = new URL("./build_lesson9_core.mjs", import.meta.url);
const detailedSourcePath = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第9课 - AI产品评测与数据反馈.md";
const outPath = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_课堂讲义_核心结构版.md";

const pptSource = await fs.readFile(pptSourcePath, "utf8");
const pptStart = pptSource.indexOf("const S=[];");
const pptEnd = pptSource.indexOf("\nfunction render(", pptStart);
if (pptStart < 0 || pptEnd < 0) throw new Error("无法读取PPT页面结构");
const S = Function(`${pptSource.slice(pptStart, pptEnd)}\nreturn S;`)();

const detailedSource = await fs.readFile(detailedSourcePath, "utf8");
const sourceLines = detailedSource.trimEnd().split(/\r?\n/);

// 每个范围均为原详细版MD中的1-based闭区间。所有3687行恰好使用一次；
// 第23页因PPT先讲选择原则，使用原文后段内容，所以页面顺序与原文顺序略有调整。
const ranges = {
  1:[[1,2]], 4:[[3,18]], 5:[[19,48]], 6:[[49,122]], 7:[[123,164]], 8:[[165,242]],
  9:[[243,272]], 10:[[273,452]], 11:[[453,520]], 12:[[521,572]], 13:[[573,620]],
  14:[[621,700]], 15:[[701,736]], 16:[[737,856]], 17:[[857,884]], 18:[[885,970]],
  19:[[971,1010]], 20:[[1011,1048]], 21:[[1049,1078]], 22:[[1079,1108]],
  23:[[1564,1591]], 24:[[1109,1150]], 25:[[1151,1204]], 26:[[1205,1294]],
  27:[[1295,1392]], 28:[[1393,1467]], 30:[[1468,1525]], 31:[[1526,1563]],
  32:[[1592,1615]], 33:[[1616,1631]], 34:[[1632,1735]], 35:[[1736,1815]],
  36:[[1816,1853]], 37:[[1854,1933]], 38:[[1934,1963]], 39:[[1964,1990]],
  40:[[1991,2054]], 41:[[2055,2072]], 42:[[2073,2202]], 43:[[2203,2504]],
  44:[[2505,2522]], 45:[[2523,2582]], 46:[[2583,2646]], 47:[[2647,2754]],
  48:[[2755,2800]], 49:[[2801,2916]], 50:[[2917,2938]], 51:[[2939,2966]],
  52:[[2967,3044]], 53:[[3045,3100]], 54:[[3101,3114]], 55:[[3115,3142]],
  56:[[3143,3235]], 57:[[3236,3279]], 58:[[3280,3330]], 59:[[3331,3396]],
  60:[[3397,3440]], 61:[[3441,3479]], 62:[[3480,3521]], 63:[[3522,3537]],
  64:[[3538,3571]], 65:[[3572,3613]], 66:[[3614,3655]], 67:[[3656,3687]]
};

function normalizeOriginal(text) {
  return text
    .split("\n")
    .filter((line) => !/^# 第[一二三四五六七八]+部分[、：]/.test(line))
    .filter((line) => line.trim() !== "---")
    .filter((line) => line.trim() !== "#")
    .map((line) => {
      if (/^#{1,4} .*?(部分的结论|部分总结)$/.test(line)) return "### 本部分总结";
      const match = line.match(/^(#{1,4})\s+(.*)$/);
      if (!match) return line;
      const level = Math.min(6, match[1].length + 2);
      return `${"#".repeat(level)} ${match[2]}`;
    })
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function originalForPage(page) {
  const pageRanges = ranges[page] ?? [];
  return normalizeOriginal(pageRanges.map(([start,end]) => sourceLines.slice(start - 1, end).join("\n")).join("\n\n"));
}

function pptFallback(slide) {
  if (slide.type === "map" || slide.type === "flow" || slide.type === "closing") {
    return slide.items.map((item, i) => `${i + 1}. ${String(item).replaceAll("\n", "：")}`).join("\n");
  }
  if (slide.type === "cards") {
    return slide.items.map((item) => `### ${item.k}\n\n${item.v}`).join("\n\n");
  }
  if (slide.type === "code") return `\`\`\`json\n${slide.code}\n\`\`\``;
  if (slide.subtitle) return slide.subtitle;
  return "";
}

// 完整性校验：原详细版的每一行必须且只能被分配一次。
const coverage = new Array(sourceLines.length).fill(0);
for (const pageRanges of Object.values(ranges)) {
  for (const [start,end] of pageRanges) {
    for (let line = start; line <= end; line++) coverage[line - 1] += 1;
  }
}
const missing = coverage.flatMap((count, i) => count === 0 ? [i + 1] : []);
const duplicated = coverage.flatMap((count, i) => count > 1 ? [i + 1] : []);
if (missing.length || duplicated.length) {
  throw new Error(`原文分页范围异常：missing=${missing.slice(0,20)} duplicated=${duplicated.slice(0,20)}`);
}

const pages = S.map((slide, index) => {
  const page = index + 1;
  const body = originalForPage(page) || pptFallback(slide);
  return `## 第${String(page).padStart(2,"0")}页｜${slide.title}\n\n${body}`.trim();
});

const markdown = `# 第9课：AI产品评测与数据反馈｜课堂讲义\n\n` +
  `本讲义使用原详细版课程MD作为正文素材，按照核心结构版PPT的67页顺序重新分页。\n\n` +
  `${pages.join("\n\n---\n\n")}\n`;

await fs.writeFile(outPath, markdown, "utf8");
console.log(JSON.stringify({slides:S.length,sourceLines:sourceLines.length,outputBytes:Buffer.byteLength(markdown),outPath}));
