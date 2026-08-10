#!/usr/bin/env python3
"""Convert a JSON blog payload (slug, title, meta_description, body markdown)
into a formatted .docx saved under blogs/{date}-{slug}.docx.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...", "meta_description": "...", "body": "..."}'
"""

import argparse
import json
import os
import re
from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")

LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1155CC")
    rpr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)

    new_run.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_runs_with_inline_markup(paragraph, text):
    """Add text to a paragraph, handling **bold** and [text](url) links."""
    tokens = []
    pos = 0
    combined_re = re.compile(r"(\*\*[^*]+\*\*|\[[^\]]+\]\([^)]+\))")
    for match in combined_re.finditer(text):
        if match.start() > pos:
            tokens.append(("text", text[pos:match.start()]))
        chunk = match.group(0)
        if chunk.startswith("**"):
            tokens.append(("bold", chunk[2:-2]))
        else:
            link_match = LINK_RE.match(chunk)
            tokens.append(("link", (link_match.group(1), link_match.group(2))))
        pos = match.end()
    if pos < len(text):
        tokens.append(("text", text[pos:]))

    for kind, value in tokens:
        if kind == "text":
            if value:
                paragraph.add_run(value)
        elif kind == "bold":
            run = paragraph.add_run(value)
            run.bold = True
        elif kind == "link":
            link_text, url = value
            add_hyperlink(paragraph, url, link_text)


def build_document(title, meta_description, body_markdown):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    heading = doc.add_heading(title, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

    meta_para = doc.add_paragraph()
    meta_run = meta_para.add_run(f"Meta description: {meta_description}")
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    doc.add_paragraph()

    for raw_line in body_markdown.split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        else:
            para = doc.add_paragraph()
            add_runs_with_inline_markup(para, line.strip())

    return doc


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="JSON payload with slug, title, meta_description, body")
    args = parser.parse_args()

    payload = json.loads(args.json)
    slug = payload["slug"]
    title = payload["title"]
    meta_description = payload["meta_description"]
    body = payload["body"]

    os.makedirs(BLOGS_DIR, exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"{date_str}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    doc = build_document(title, meta_description, body)
    doc.save(out_path)

    print(out_path)


if __name__ == "__main__":
    main()
