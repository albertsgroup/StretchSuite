#!/usr/bin/env python3
"""
Stretch Suite Blog Generator
Generates a .docx from blog JSON and emails it via Resend.

Usage:
    python blog_generator.py --json '{"slug":"...", "title":"...", "meta_description":"...", "body":"..."}'
"""

import argparse
import base64
import json
import os
import re
import sys
import requests

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_FROM    = "Stretch Suite Blog <onboarding@resend.dev>"
RESEND_TO      = "hammad@albertsgroup.net"
BLOGS_DIR      = os.path.join(os.path.dirname(__file__), "blogs")


def md_to_docx(title, meta_description, body_md, out_path):
    doc = Document()

    # --- styles ---
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # H1 title
    h1 = doc.add_heading(title, level=1)
    h1.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Meta description block
    meta_para = doc.add_paragraph()
    meta_run = meta_para.add_run(f"Meta Description: {meta_description}")
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    doc.add_paragraph()  # spacer

    lines = body_md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()

        # skip the H1 in the body (already added above)
        if line.startswith("# "):
            i += 1
            continue

        if line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)

        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)

        elif re.match(r"^[-*] ", line):
            # bullet list
            text = _strip_inline_md(line[2:].strip())
            doc.add_paragraph(text, style="List Bullet")

        elif line.strip() == "":
            # blank line -- paragraph break handled by surrounding context
            pass

        else:
            text = _strip_inline_md(line)
            if text:
                doc.add_paragraph(text)

        i += 1

    doc.save(out_path)


def _strip_inline_md(text):
    """Remove inline markdown (links, bold, italic) leaving plain text."""
    # [label](url) -> label
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    # **bold** or __bold__
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"__(.+?)__", r"\1", text)
    # *italic* or _italic_
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"_(.+?)_", r"\1", text)
    return text


def send_email(slug, title, meta_description, docx_path, date_prefix):
    with open(docx_path, "rb") as f:
        docx_b64 = base64.b64encode(f.read()).decode("utf-8")

    filename = f"{date_prefix}-{slug}.docx"

    html_body = f"""
<h2>{title}</h2>
<p><em>Meta description:</em> {meta_description}</p>
<p>The full blog post is attached as a .docx file (<strong>{filename}</strong>).</p>
<p>Ready to publish on <a href="https://www.stretchsuite.com">stretchsuite.com</a>.</p>
"""

    payload = {
        "from": RESEND_FROM,
        "to": [RESEND_TO],
        "subject": "New blog is ready to view",
        "html": html_body,
        "attachments": [
            {
                "filename": filename,
                "content": docx_b64,
            }
        ],
    }

    resp = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if resp.status_code in (200, 201):
        print(f"Email sent successfully. Response: {resp.json()}")
    else:
        print(f"Email send failed [{resp.status_code}]: {resp.text}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, dest="json_str", help="JSON string with blog fields")
    args = parser.parse_args()

    try:
        data = json.loads(args.json_str)
    except json.JSONDecodeError as e:
        print(f"Invalid JSON: {e}", file=sys.stderr)
        sys.exit(1)

    slug            = data["slug"]
    title           = data["title"]
    meta_description = data["meta_description"]
    body            = data["body"]

    # derive date prefix from today (YYYY-MM-DD)
    from datetime import date
    date_prefix = date.today().isoformat()

    os.makedirs(BLOGS_DIR, exist_ok=True)
    filename = f"{date_prefix}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    print(f"Generating {filename} ...")
    md_to_docx(title, meta_description, body, out_path)
    print(f"Saved to {out_path}")

    print("Sending email via Resend ...")
    send_email(slug, title, meta_description, out_path, date_prefix)


if __name__ == "__main__":
    main()
