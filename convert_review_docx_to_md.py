from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


SOURCE = Path("/Users/keivn/Project/AI-Course/AI_Agent项目作业课堂讲评手册.docx")
OUTPUT = Path("/Users/keivn/Project/AI-Course/AI_Agent项目作业课堂讲评手册.md")


def inline_markdown(paragraph):
    parts = []
    for child in paragraph._p.iterchildren():
        if child.tag == qn("w:r"):
            text = "".join(node.text or "" for node in child.iter(qn("w:t")))
            if text:
                parts.append(text)
        elif child.tag == qn("w:hyperlink"):
            label = "".join(node.text or "" for node in child.iter(qn("w:t")))
            rel_id = child.get(qn("r:id"))
            url = paragraph.part.rels[rel_id].target_ref if rel_id else ""
            parts.append(f"[{label}]({url})" if url else label)
    return "".join(parts).strip()


def escape_cell(text):
    return " ".join(text.split()).replace("|", "\\|")


def paragraph_to_markdown(paragraph):
    text = inline_markdown(paragraph)
    if not text:
        return ""
    style = paragraph.style.name if paragraph.style else ""
    if style == "Title":
        return f"# {text}"
    if style.startswith("Heading"):
        try:
            level = int(style.split()[-1])
        except ValueError:
            level = 2
        return f"{'#' * min(level + 1, 6)} {text}"
    if style.startswith("List Bullet"):
        indent = "  " if style.endswith("2") else ""
        return f"{indent}- {text}"
    return text


def table_to_markdown(table):
    rows = [[escape_cell(cell.text) for cell in row.cells] for row in table.rows]
    if not rows:
        return []
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = ["| " + " | ".join(rows[0]) + " |"]
    lines.append("| " + " | ".join(["---"] * width) + " |")
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return lines


doc = Document(SOURCE)
lines = []
for child in doc.element.body.iterchildren():
    if child.tag == qn("w:p"):
        rendered = paragraph_to_markdown(Paragraph(child, doc))
        if rendered:
            lines.extend([rendered, ""])
    elif child.tag == qn("w:tbl"):
        lines.extend(table_to_markdown(Table(child, doc)))
        lines.append("")

content = "\n".join(lines).strip() + "\n"
OUTPUT.write_text(content, encoding="utf-8")
print(OUTPUT)
