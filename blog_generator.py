#!/usr/bin/env python3
"""Stretch Suite blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file saved
in ./blogs/, then emails it as a base64 attachment via the Resend API.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...",
        "meta_description": "...", "body": "full markdown body"}'
"""

import argparse
import base64
import json
import re
import sys
from datetime import date
from pathlib import Path

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
RECIPIENTS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]

BLOGS_DIR = Path(__file__).resolve().parent / "blogs"

LINK_RE = re.compile(r"\[([^\]]+)\]\((https?://[^\)]+)\)")
BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def add_hyperlink(paragraph, url, text):
    """Append a clickable hyperlink run to a paragraph."""
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


def add_formatted_runs(paragraph, text):
    """Parse inline markdown (links + bold) into runs on a paragraph."""
    pos = 0
    tokens = []
    for m in LINK_RE.finditer(text):
        tokens.append((m.start(), m.end(), "link", m.group(1), m.group(2)))
    # Only handle bold in segments that are not links.
    combined = sorted(tokens, key=lambda t: t[0])

    idx = 0
    while idx < len(text):
        link = next((t for t in combined if t[0] == idx), None)
        if link:
            _, end, _, label, url = link
            add_hyperlink(paragraph, url, label)
            idx = end
            continue
        # find next link start
        next_link_start = min(
            [t[0] for t in combined if t[0] > idx] + [len(text)]
        )
        segment = text[idx:next_link_start]
        _add_bold_aware(paragraph, segment)
        idx = next_link_start


def _add_bold_aware(paragraph, segment):
    last = 0
    for m in BOLD_RE.finditer(segment):
        if m.start() > last:
            paragraph.add_run(segment[last:m.start()])
        run = paragraph.add_run(m.group(1))
        run.bold = True
        last = m.end()
    if last < len(segment):
        paragraph.add_run(segment[last:])


def build_docx(title, meta_description, body, out_path):
    doc = Document()

    h1 = doc.add_heading(title, level=0)

    meta = doc.add_paragraph()
    meta_run = meta.add_run("Meta description: ")
    meta_run.bold = True
    meta_desc_run = meta.add_run(meta_description)
    meta_desc_run.italic = True
    meta_desc_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()

    for raw_line in body.split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.lstrip().startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            add_formatted_runs(p, line.lstrip()[2:].strip())
        else:
            p = doc.add_paragraph()
            add_formatted_runs(p, line.strip())

    doc.save(out_path)


def send_email(title, meta_description, out_path):
    with open(out_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")

    html_body = (
        f"<h1>{title}</h1>"
        f"<p><strong>Meta description:</strong> {meta_description}</p>"
        f"<p>This week's Stretch Suite blog post is ready. "
        f"The full post is attached as a Word document "
        f"(<code>{out_path.name}</code>).</p>"
        f"<p>Move Better. Feel Better. Stay Independent.</p>"
    )

    payload = {
        "from": FROM_ADDRESS,
        "to": RECIPIENTS,
        "subject": "New blog - Stretch Suite",
        "html": html_body,
        "attachments": [
            {"filename": out_path.name, "content": encoded}
        ],
    }

    resp = requests.post(
        RESEND_ENDPOINT,
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        data=json.dumps(payload),
        timeout=60,
    )
    return resp


def main():
    parser = argparse.ArgumentParser(description="Generate and email a Stretch Suite blog post.")
    parser.add_argument("--json", required=True, help="JSON payload for the blog post")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    BLOGS_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    filename = f"{today}-{slug}.docx"
    out_path = BLOGS_DIR / filename

    build_docx(title, meta_description, body, out_path)
    print(f"Saved blog to: {out_path}")

    resp = send_email(title, meta_description, out_path)
    if resp.status_code in (200, 201):
        print(f"Email sent successfully: {resp.json()}")
    else:
        print(f"Email send FAILED ({resp.status_code}): {resp.text}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
