"""
Blog generator for Stretch Suite weekly blog automation.

Takes a JSON payload with slug/title/meta_description/body (markdown),
renders a .docx into /blogs/, and emails it via Resend as an attachment.
"""

import argparse
import base64
import datetime
import json
import os
import re
import sys
from pathlib import Path

import requests
from docx import Document
from docx.shared import Pt, RGBColor

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_URL = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite Blog <onboarding@resend.dev>"
TO_ADDRESS = "hammad@albertsgroup.net"

BLOGS_DIR = Path(__file__).parent / "blogs"


def add_hyperlink(paragraph, url, text):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(underline)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)
    return hyperlink


INLINE_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def add_inline_runs(paragraph, text):
    """Parse inline markdown (links + bold) and add runs to the paragraph."""
    pos = 0
    while pos < len(text):
        link = INLINE_LINK_RE.search(text, pos)
        bold = BOLD_RE.search(text, pos)

        next_match = None
        kind = None
        for m, k in [(link, "link"), (bold, "bold")]:
            if m is None:
                continue
            if next_match is None or m.start() < next_match.start():
                next_match = m
                kind = k

        if next_match is None:
            if pos < len(text):
                paragraph.add_run(text[pos:])
            break

        if next_match.start() > pos:
            paragraph.add_run(text[pos:next_match.start()])

        if kind == "link":
            add_hyperlink(paragraph, next_match.group(2), next_match.group(1))
        else:
            run = paragraph.add_run(next_match.group(1))
            run.bold = True

        pos = next_match.end()


def render_markdown_to_docx(title, meta_description, body_md, out_path):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    h1 = doc.add_heading(title, level=1)

    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run(f"Meta description: {meta_description}")
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    lines = body_md.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()

        if not line.strip():
            i += 1
            continue

        if line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            add_inline_runs(p, line[2:].strip())
        elif re.match(r"^\d+\.\s", line):
            content = re.sub(r"^\d+\.\s+", "", line)
            p = doc.add_paragraph(style="List Number")
            add_inline_runs(p, content)
        else:
            p = doc.add_paragraph()
            add_inline_runs(p, line.strip())

        i += 1

    doc.save(out_path)


def send_email(title, meta_description, docx_path):
    with open(docx_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")

    html = f"""
    <div style="font-family: -apple-system, Segoe UI, sans-serif; max-width: 640px;">
      <h2 style="color:#222;">{title}</h2>
      <p style="color:#555; font-style: italic;">{meta_description}</p>
      <p>The new Stretch Suite blog post is attached as a .docx file, ready to review and publish.</p>
      <p style="color:#888; font-size: 12px;">File: {docx_path.name}</p>
    </div>
    """

    payload = {
        "from": FROM_ADDRESS,
        "to": [TO_ADDRESS],
        "subject": "New blog is ready to view",
        "html": html,
        "attachments": [
            {
                "filename": docx_path.name,
                "content": encoded,
            }
        ],
    }

    resp = requests.post(
        RESEND_URL,
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Blog payload as JSON string")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    BLOGS_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.date.today().isoformat()
    filename = f"{today}-{slug}.docx"
    out_path = BLOGS_DIR / filename

    render_markdown_to_docx(title, meta_description, body, out_path)
    print(f"Wrote {out_path}")

    md_path = BLOGS_DIR / f"{today}-{slug}.md"
    md_path.write_text(
        f"# {title}\n\n_Meta: {meta_description}_\n\n{body}\n",
        encoding="utf-8",
    )
    print(f"Wrote {md_path}")

    result = send_email(title, meta_description, out_path)
    print(f"Resend response: {result}")


if __name__ == "__main__":
    main()
