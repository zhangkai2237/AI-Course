import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const courseDir = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第10课_AI产品方案与PRD设计实战_重做工作区";
const outputPath = path.join(courseDir, "第10课_AI产品方案与PRD设计实战_完整讲义.md");

const names = (await readdir(courseDir))
  .filter((name) => /^0[1-7]_.*\.md$/.test(name))
  .sort((a, b) => a.localeCompare(b, "zh-CN", { numeric: true }));

if (names.length !== 7) {
  throw new Error(`Expected 7 section files, found ${names.length}: ${names.join(", ")}`);
}

const sections = [];
for (const name of names) {
  const content = await readFile(path.join(courseDir, name), "utf8");
  sections.push(content.trimEnd());
}

await writeFile(outputPath, `${sections.join("\n\n---\n\n")}\n`, "utf8");
console.log(outputPath);
