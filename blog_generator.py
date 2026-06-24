#!/usr/bin/env python3
"""
Stretch Suite blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file saved
under blogs/, and emails it as a base64 attachment via the Resend API.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...",
                                       "meta_description": "...", "body": "..."}'
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

# --- Configuration ---------------------------------------------------------

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
TO_ADDRESSES = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]
SUBJECT = "New blog - Stretch Suite"

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")

LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


# --- Markdown -> docx helpers ----------------------------------------------

def add_hyperlink(paragraph, url, text):
    """Add a real clickable hyperlink run to a paragraph."""
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
    return hyperlink


def add_rich_text(paragraph, text):
    """Render a line of markdown text, turning [label](url) into hyperlinks."""
    pos = 0
    for match in LINK_RE.finditer(text):
        if match.start() > pos:
            paragraph.add_run(text[pos:match.start()])
        add_hyperlink(paragraph, match.group(2), match.group(1))
        pos = match.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def build_docx(title, meta_description, body, out_path):
    """Render the post to a .docx file at out_path."""
    doc = Document()

    # Meta description as a small grey note at the top.
    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run("Meta description: " + meta_description)
    meta_run.italic = True
    meta_run.font.size = Pt(9)
    meta_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    for raw_line in body.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.lstrip().startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            add_rich_text(p, line.lstrip()[2:].strip())
        else:
            p = doc.add_paragraph()
            add_rich_text(p, line.strip())

    doc.save(out_path)
    return out_path


# --- Email -----------------------------------------------------------------

def send_email(title, meta_description, docx_path):
    """POST the post to Resend with the .docx attached as base64."""
    with open(docx_path, "rb") as fh:
        encoded = base64.b64encode(fh.read()).decode("ascii")

    html = (
        f"<h1>{title}</h1>"
        f"<p><strong>Meta description:</strong> {meta_description}</p>"
        f"<p>This week's Stretch Suite blog post is attached as a .docx file: "
        f"<strong>{os.path.basename(docx_path)}</strong>.</p>"
        f"<p>Move Better. Feel Better. Stay Independent.</p>"
    )

    payload = {
        "from": FROM_ADDRESS,
        "to": TO_ADDRESSES,
        "subject": SUBJECT,
        "html": html,
        "attachments": [
            {
                "filename": os.path.basename(docx_path),
                "content": encoded,
            }
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
    parser.add_argument("--json", required=True, help="JSON payload for the blog post.")
    parser.add_argument("--no-email", action="store_true", help="Build the .docx but skip sending.")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    os.makedirs(BLOGS_DIR, exist_ok=True)
    date_str = datetime.date.today().isoformat()
    filename = f"{date_str}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    build_docx(title, meta_description, body, out_path)
    print(f"Saved .docx to {out_path}")

    if args.no_email:
        print("Skipping email (--no-email).")
        return

    resp = send_email(title, meta_description, out_path)
    if resp.status_code in (200, 201):
        print(f"Email sent. Resend response: {resp.text}")
    else:
        print(f"Email failed ({resp.status_code}): {resp.text}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
