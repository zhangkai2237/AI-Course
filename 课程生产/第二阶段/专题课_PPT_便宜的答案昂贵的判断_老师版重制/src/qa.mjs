import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { slides, TOTAL_MINUTES } from "./content.mjs";
import { REQUIRED_NOTE_FIELDS } from "./notes.mjs";
import { SOURCE_POLICY } from "./source-policy.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const PPTX = path.join(ROOT, "dist", "专题课_便宜的答案昂贵的判断_老师版_v2.0.pptx");
const MANIFEST = path.join(ROOT, "qa", "build-manifest.json");
const REPORT_JSON = path.join(ROOT, "qa", "static-qa-report.json");
const REPORT_TXT = path.join(ROOT, "qa", "static-qa-report.txt");

const checks = [];
function check(name, pass, detail) { checks.push({ name, pass: Boolean(pass), detail }); }

check("44页内容层", slides.length === 44, `actual=${slides.length}`);
check("总时长120分钟", TOTAL_MINUTES === 120, `actual=${TOTAL_MINUTES}`);
check("44页均有备注", slides.filter(d => d.notes?.trim()).length === 44, `actual=${slides.filter(d => d.notes?.trim()).length}`);
for (const field of REQUIRED_NOTE_FIELDS) {
  const missing = slides.filter(d => !d.notes.includes(field)).map(d => d.id);
  check(`备注字段 ${field}`, missing.length === 0, missing.length ? `missing slides=${missing.join(",")}` : "all 44 slides");
}
const sourceMissing = slides.filter(d => !d.source?.trim()).map(d => d.id);
check("44页讲义来源标记", sourceMissing.length === 0, sourceMissing.length ? `missing=${sourceMissing.join(",")}` : "all slides have source mapping");
check("五幕分页", slides.slice(0,7).every(d=>d.act.includes("第一幕")) && slides.slice(7,14).every(d=>d.act.includes("第二幕")) && slides.slice(14,29).every(d=>d.act.includes("第三幕")) && slides.slice(29,36).every(d=>d.act.includes("第四幕")) && slides.slice(36,44).every(d=>d.act.includes("第五幕")), "P01-07 / P08-14 / P15-29 / P30-36 / P37-44");
check("P21—P23原生线框类型", [slides[20].type, slides[21].type, slides[22].type].join(",") === "wireConversation,wireOrder,wireHandoff", `${slides[20].type}, ${slides[21].type}, ${slides[22].type}`);
const publicSlides = slides.filter(d => d.publicCase);
check("公共案例均显式标记", publicSlides.length >= 12 && publicSlides.every(d => d.notes.includes("公共模拟") || d.kicker?.includes("公共模拟") || d.publicCase), `marked=${publicSlides.length}`);
const unsafe = [];
const claimText = JSON.stringify(slides.map(({ notes, ...visible }) => visible));
for (const phrase of ["真实用户自助率提升", "实际会话量达到", "已经接入真实订单", "在真实环境中运行"]) {
  if (claimText.includes(phrase)) unsafe.push(phrase);
}
check("无无条件真实成效断言", unsafe.length === 0, unsafe.length ? unsafe.join("; ") : "no prohibited unconditional claims");
check("唯一内容源为指定Markdown", SOURCE_POLICY.soleContentSource.endsWith("07_专题课_便宜的答案昂贵的判断_前五节老师讲义_v1.0.md"), SOURCE_POLICY.soleContentSource);
check("内容源不是PPT", !SOURCE_POLICY.soleContentSource.toLowerCase().match(/\.pptx?$/), SOURCE_POLICY.soleContentSource);
const requiredFiles = ["package.json", "README.md", "src/content.mjs", "src/theme.mjs", "src/components.mjs", "src/wireframes.mjs", "src/notes.mjs", "src/source-policy.mjs", "src/build.mjs", "src/qa.mjs"];
const fileResults = await Promise.all(requiredFiles.map(async f => { try { await fs.access(path.join(ROOT, f)); return true; } catch { return false; } }));
check("独立工程包结构", fileResults.every(Boolean), requiredFiles.filter((_,i)=>!fileResults[i]).join(",") || "all required files present");
let pptxExists = false, manifest = null;
try { const st = await fs.stat(PPTX); pptxExists = st.size > 0; } catch {}
try { manifest = JSON.parse(await fs.readFile(MANIFEST, "utf8")); } catch {}
check("正式PPTX已生成", pptxExists, path.relative(ROOT, PPTX));
check("构建清单44页/120分钟/44备注", manifest?.slideCount === 44 && manifest?.totalMinutes === 120 && manifest?.noteCount === 44, manifest ? JSON.stringify({ slides: manifest.slideCount, minutes: manifest.totalMinutes, notes: manifest.noteCount }) : "manifest missing");
const buildInputs = manifest?.buildInputs || [];
const forbiddenInputs = buildInputs.filter(p => /\.pptx?$/i.test(p));
check("构建清单无旧PPT输入", forbiddenInputs.length === 0, forbiddenInputs.length ? forbiddenInputs.join(",") : "0 PPT/PPTX inputs");
check("PptxGenJS固定4.0.1", manifest?.dependency === "pptxgenjs@4.0.1", manifest?.dependency || "manifest missing");

const passed = checks.filter(c=>c.pass).length;
const report = {
  generatedAt: new Date().toISOString(),
  status: passed === checks.length ? "PASS" : "FAIL",
  summary: { passed, failed: checks.length - passed, total: checks.length },
  checks,
  manualReviewRequired: [
    "在Microsoft PowerPoint中逐页检查文本溢出、字体回退与对齐",
    "检查P21—P23可编辑低保真UI在目标环境中的形状和层级",
    "确认演讲者备注在PowerPoint备注窗格中完整显示",
    "核对PDF导出后的44页分页、中文字体和裁切",
    "授课前用真实计时试讲并按班级互动速度微调",
  ],
};
await fs.mkdir(path.dirname(REPORT_JSON), { recursive: true });
await fs.writeFile(REPORT_JSON, JSON.stringify(report, null, 2), "utf8");
const text = [
  `静态QA：${report.status}`,
  `通过 ${passed}/${checks.length}；失败 ${checks.length - passed}`,
  "",
  ...checks.map(c => `${c.pass ? "PASS" : "FAIL"} | ${c.name} | ${c.detail}`),
  "",
  "仍需人工检查：",
  ...report.manualReviewRequired.map(v => `- ${v}`),
].join("\n");
await fs.writeFile(REPORT_TXT, text, "utf8");
console.log(text);
if (report.status !== "PASS") process.exitCode = 1;
