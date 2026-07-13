#!/usr/bin/env python3
"""Stretch Suite weekly blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file saved
under blogs/, and emails it as a base64-encoded attachment via the Resend API.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...",
        "meta_description": "...", "body": "...markdown..."}'
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
from docx.shared import Pt, RGBColor

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_URL = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
RECIPIENTS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")

# Inline markdown link pattern: [text](url)
LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def _add_inline(paragraph, text):
    """Render inline markdown (links + bold) into a docx paragraph."""
    pos = 0
    # Combine link and bold handling by scanning tokens left to right.
    tokens = []
    for m in LINK_RE.finditer(text):
        tokens.append((m.start(), m.end(), "link", m.group(1), m.group(2)))
    for m in BOLD_RE.finditer(text):
        tokens.append((m.start(), m.end(), "bold", m.group(1), None))
    tokens.sort(key=lambda t: t[0])

    for start, end, kind, content, url in tokens:
        if start < pos:
            continue  # overlapping, skip
        if start > pos:
            paragraph.add_run(text[pos:start])
        run = paragraph.add_run(content)
        if kind == "link":
            run.font.color.rgb = RGBColor(0x1A, 0x73, 0xE8)
            run.font.underline = True
        elif kind == "bold":
            run.bold = True
        pos = end
    if pos < len(text):
        paragraph.add_run(text[pos:])


def markdown_to_docx(title, meta_description, body):
    doc = Document()

    normal = doc.styles["Normal"].font
    normal.name = "Calibri"
    normal.size = Pt(11)

    for line in body.splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        if stripped.startswith("# "):
            doc.add_heading(stripped[2:].strip(), level=0)
        elif stripped.startswith("### "):
            doc.add_heading(stripped[4:].strip(), level=2)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:].strip(), level=1)
        elif stripped.startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            _add_inline(p, stripped[2:].strip())
        elif re.match(r"^\d+\.\s", stripped):
            p = doc.add_paragraph(style="List Number")
            _add_inline(p, re.sub(r"^\d+\.\s", "", stripped))
        elif stripped.startswith("> "):
            p = doc.add_paragraph()
            p.add_run(stripped[2:].strip()).italic = True
        else:
            p = doc.add_paragraph()
            _add_inline(p, stripped)

    return doc


def send_email(title, meta_description, filename, docx_bytes):
    b64 = base64.b64encode(docx_bytes).decode("ascii")
    html = (
        f"<h2>{title}</h2>"
        f"<p><strong>Meta description:</strong> {meta_description}</p>"
        f"<p>This week's Stretch Suite blog post is attached as a .docx file "
        f"(<em>{filename}</em>), ready for review and publishing.</p>"
        f"<p>Move Better. Feel Better. Stay Independent.</p>"
    )
    payload = {
        "from": FROM_ADDRESS,
        "to": RECIPIENTS,
        "subject": "New blog - Stretch Suite",
        "html": html,
        "attachments": [{"filename": filename, "content": b64}],
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
    parser.add_argument("--json", required=True, help="Blog post JSON payload")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    os.makedirs(BLOGS_DIR, exist_ok=True)
    date = datetime.date.today().isoformat()
    filename = f"{date}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    doc = markdown_to_docx(title, meta_description, body)
    doc.save(out_path)
    print(f"Saved blog to {out_path}")

    with open(out_path, "rb") as f:
        docx_bytes = f.read()

    result = send_email(title, meta_description, filename, docx_bytes)
    print(f"Email sent via Resend: {json.dumps(result)}")


if __name__ == "__main__":
    main()
