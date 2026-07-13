#!/usr/bin/env python3
"""Stretch Suite blog generator.

Takes a JSON payload describing a blog post (slug, title, meta_description,
markdown body), renders it to a .docx file saved under ./blogs/, and emails it
as a base64 attachment to the Stretch Suite team via the Resend API.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...",
        "meta_description": "...", "body": "# markdown ..."}'
"""

import argparse
import base64
import datetime
import json
import os
import re
import sys

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_URL = "https://api.resend.com/emails"
EMAIL_FROM = "Stretch Suite <hello@stretchsuite.com>"
EMAIL_TO = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]
EMAIL_SUBJECT = "New blog - Stretch Suite"

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")

# Inline markdown: **bold** and [text](url)
LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def _add_hyperlink(paragraph, url, text):
    """Add a clickable hyperlink run to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
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


def _tokenize_inline(text):
    """Split a line into (kind, payload) tokens for links and bold text."""
    tokens = []
    pos = 0
    # Find links first, then handle bold within the non-link segments.
    for m in LINK_RE.finditer(text):
        if m.start() > pos:
            tokens.extend(_split_bold(text[pos:m.start()]))
        tokens.append(("link", (m.group(1), m.group(2))))
        pos = m.end()
    if pos < len(text):
        tokens.extend(_split_bold(text[pos:]))
    return tokens


def _split_bold(text):
    tokens = []
    pos = 0
    for m in BOLD_RE.finditer(text):
        if m.start() > pos:
            tokens.append(("text", text[pos:m.start()]))
        tokens.append(("bold", m.group(1)))
        pos = m.end()
    if pos < len(text):
        tokens.append(("text", text[pos:]))
    return tokens


def _render_inline(paragraph, text):
    for kind, payload in _tokenize_inline(text):
        if kind == "link":
            label, url = payload
            _add_hyperlink(paragraph, url, label)
        elif kind == "bold":
            run = paragraph.add_run(payload)
            run.bold = True
        else:
            paragraph.add_run(payload)


def build_docx(title, meta_description, body, out_path):
    doc = Document()

    # Title
    heading = doc.add_heading(title, level=0)

    # Meta description as a subtle italic note for the SEO/editor hand-off.
    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run("Meta description: " + meta_description)
    meta_run.italic = True
    meta_run.font.size = Pt(9)
    meta_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    for raw in body.split("\n"):
        line = raw.rstrip()
        if not line.strip():
            continue
        stripped = line.strip()

        if stripped.startswith("# "):
            # Skip a duplicated H1 in the body; title is already rendered.
            continue
        elif stripped.startswith("### "):
            p = doc.add_heading(level=3)
            _render_inline(p, stripped[4:].strip())
        elif stripped.startswith("## "):
            p = doc.add_heading(level=2)
            _render_inline(p, stripped[3:].strip())
        elif stripped.startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            _render_inline(p, stripped[2:].strip())
        elif re.match(r"^\d+\.\s", stripped):
            p = doc.add_paragraph(style="List Number")
            _render_inline(p, re.sub(r"^\d+\.\s", "", stripped))
        else:
            p = doc.add_paragraph()
            _render_inline(p, stripped)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path


def send_email(title, meta_description, docx_path, filename, ps_line=""):
    with open(docx_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")

    html_body = (
        f"<h1>{title}</h1>"
        f"<p><strong>Meta description:</strong> {meta_description}</p>"
        f"<p>The full blog post is attached as a .docx file "
        f"(<em>{filename}</em>).</p>"
        f"{ps_line}"
    )

    payload = {
        "from": EMAIL_FROM,
        "to": EMAIL_TO,
        "subject": EMAIL_SUBJECT,
        "html": html_body,
        "attachments": [
            {"filename": filename, "content": encoded}
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
    return resp


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Blog post JSON payload")
    parser.add_argument("--date", default=None, help="YYYY-MM-DD (defaults to today)")
    parser.add_argument("--ps", default="", help="Optional PS line appended to email HTML")
    parser.add_argument("--no-email", action="store_true", help="Only build the .docx")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    date = args.date or datetime.date.today().isoformat()
    filename = f"{date}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    build_docx(title, meta_description, body, out_path)
    print(f"Saved .docx to {out_path}")

    if args.no_email:
        return

    ps_line = f"<p>{args.ps}</p>" if args.ps else ""
    resp = send_email(title, meta_description, out_path, filename, ps_line)
    print(f"Resend status: {resp.status_code}")
    print(f"Resend response: {resp.text}")
    if resp.status_code >= 300:
        sys.exit(1)


if __name__ == "__main__":
    main()
