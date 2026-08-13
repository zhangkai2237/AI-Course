const fs = require('fs');

const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('usage: node md_to_html.js input.md output.html');

const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const inline = s => esc(s)
  .replace(/`([^`]+)`/g, '<code>$1</code>')
  .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
  .replace(/\*([^*]+)\*/g, '<em>$1</em>');

const lines = fs.readFileSync(input, 'utf8').split(/\r?\n/);
let out = [];
let list = null;
const closeList = () => { if (list) { out.push(`</${list}>`); list = null; } };

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  if (!line.trim()) { closeList(); continue; }
  if (/^---+$/.test(line.trim())) { closeList(); out.push('<hr>'); continue; }
  const h = line.match(/^(#{1,4})\s+(.+)$/);
  if (h) { closeList(); out.push(`<h${h[1].length}>${inline(h[2])}</h${h[1].length}>`); continue; }
  if (line.startsWith('> ')) { closeList(); out.push(`<blockquote>${inline(line.slice(2))}</blockquote>`); continue; }
  if (line.includes('|') && i + 1 < lines.length && /^\s*\|?\s*:?-+/.test(lines[i + 1])) {
    closeList();
    const rows = [];
    const cells = s => s.trim().replace(/^\||\|$/g, '').split('|').map(x => x.trim());
    rows.push(cells(line)); i++;
    while (i + 1 < lines.length && lines[i + 1].includes('|') && lines[i + 1].trim()) rows.push(cells(lines[++i]));
    out.push('<table><thead><tr>' + rows[0].map(x => `<th>${inline(x)}</th>`).join('') + '</tr></thead><tbody>');
    for (const row of rows.slice(1)) out.push('<tr>' + row.map(x => `<td>${inline(x)}</td>`).join('') + '</tr>');
    out.push('</tbody></table>'); continue;
  }
  const ul = line.match(/^\s*-\s+(.+)$/);
  const ol = line.match(/^\s*\d+\.\s+(.+)$/);
  if (ul || ol) {
    const type = ul ? 'ul' : 'ol';
    if (list !== type) { closeList(); list = type; out.push(`<${type}>`); }
    out.push(`<li>${inline((ul || ol)[1])}</li>`); continue;
  }
  closeList(); out.push(`<p>${inline(line.trim())}</p>`);
}
closeList();

const html = `<!doctype html><html><head><meta charset="utf-8"><style>
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;line-height:1.65;color:#222;margin:40px}h1{color:#17365d;text-align:center}h2{color:#1f4e79;border-bottom:2px solid #5b9bd5;padding-bottom:4px}h3{color:#2f5597}h4{color:#444}table{border-collapse:collapse;width:100%;margin:10px 0 18px}th,td{border:1px solid #aaa;padding:6px;vertical-align:top}th{background:#d9eaf7}blockquote{background:#f3f7fb;border-left:4px solid #5b9bd5;margin:10px 0;padding:8px 12px}code{background:#eee;padding:1px 4px}li{margin:3px 0}hr{border:0;border-top:1px solid #bbb;margin:20px 0}
</style></head><body>${out.join('\n')}</body></html>`;
fs.writeFileSync(output, html);
