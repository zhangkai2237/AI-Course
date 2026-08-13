const fs = require('fs');
const JSZip = require('jszip');

function decodeXml(s) {
  return s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

async function extract(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file('word/document.xml').async('string');
  const paras = [];
  for (const p of xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || []) {
    const parts = [];
    for (const m of p.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) parts.push(decodeXml(m[1]));
    paras.push(parts.join('').trim());
  }
  return paras;
}

async function main() {
  const [file, ...keys] = process.argv.slice(2);
  if (!file) throw new Error('usage: node extract_docx_paragraphs.js file.docx [keyword ...]');
  const paras = await extract(file);
  console.log(`# ${file}\n# paragraphs=${paras.length}`);
  if (!keys.length) {
    paras.forEach((p, i) => { if (p) console.log(`P${i + 1}: ${p}`); });
    return;
  }
  const hits = paras.map((p, i) => keys.some(k => p.toLowerCase().includes(k.toLowerCase())) ? i : -1).filter(i => i >= 0);
  const ranges = [];
  for (const i of hits) {
    const a = Math.max(0, i - 2), b = Math.min(paras.length, i + 3);
    if (ranges.length && a <= ranges[ranges.length - 1][1]) ranges[ranges.length - 1][1] = Math.max(ranges[ranges.length - 1][1], b);
    else ranges.push([a, b]);
  }
  for (const [a, b] of ranges) {
    console.log(`\n## P${a + 1}—P${b}`);
    for (let i = a; i < b; i++) if (paras[i]) console.log(`P${i + 1}: ${paras[i]}`);
  }
}

main().catch(e => { console.error(e); process.exit(1); });
