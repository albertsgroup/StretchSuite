#!/usr/bin/env python3
"""
Stretch Suite blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file in the
/blogs/ folder, and emails it as a base64-encoded attachment via the Resend API.

Usage:
    python blog_generator.py --json '<json>'

Where <json> is:
{
  "slug": "url-friendly-slug",
  "title": "H1 title",
  "meta_description": "150-160 char meta description",
  "body": "full markdown body"
}
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
from docx.shared import Pt

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
RECIPIENTS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")


def slugify(value):
    value = value.lower().strip()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    return value.strip("-")


def add_markdown_to_doc(doc, markdown_body):
    """Render a (lightweight) markdown body into the python-docx document."""
    lines = markdown_body.split("\n")
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.lstrip().startswith(("- ", "* ")):
            text = line.lstrip()[2:].strip()
            _add_rich_paragraph(doc, text, style="List Bullet")
        else:
            _add_rich_paragraph(doc, line.strip())


def _add_rich_paragraph(doc, text, style=None):
    """Add a paragraph, rendering **bold** and [link](url) inline."""
    paragraph = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    # Split on markdown links and bold markers, preserving the delimiters.
    token_pattern = re.compile(r"(\[[^\]]+\]\([^)]+\)|\*\*[^*]+\*\*)")
    for token in token_pattern.split(text):
        if not token:
            continue
        link_match = re.match(r"\[([^\]]+)\]\(([^)]+)\)", token)
        bold_match = re.match(r"\*\*([^*]+)\*\*", token)
        if link_match:
            run = paragraph.add_run(link_match.group(1))
            run.italic = True
        elif bold_match:
            run = paragraph.add_run(bold_match.group(1))
            run.bold = True
        else:
            paragraph.add_run(token)


def build_docx(payload, out_path):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_heading(payload["title"], level=0)

    meta = doc.add_paragraph()
    meta_label = meta.add_run("Meta description: ")
    meta_label.bold = True
    meta.add_run(payload["meta_description"])

    doc.add_paragraph("")

    add_markdown_to_doc(doc, payload["body"])
    doc.save(out_path)
    return out_path


def send_email(payload, docx_path, filename):
    with open(docx_path, "rb") as fh:
        encoded = base64.b64encode(fh.read()).decode("utf-8")

    html = (
        f"<h1>{payload['title']}</h1>"
        f"<p><strong>Meta description:</strong> {payload['meta_description']}</p>"
        f"<p>The full blog post is attached as a .docx file "
        f"(<strong>{filename}</strong>).</p>"
        f"<p>Move Better. Feel Better. Stay Independent.</p>"
    )

    body = {
        "from": FROM_ADDRESS,
        "to": RECIPIENTS,
        "subject": "New blog - Stretch Suite",
        "html": html,
        "attachments": [
            {
                "filename": filename,
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
        json=body,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


def main():
    parser = argparse.ArgumentParser(description="Generate and email a Stretch Suite blog post.")
    parser.add_argument("--json", required=True, help="JSON payload with slug, title, meta_description, body")
    args = parser.parse_args()

    payload = json.loads(args.json)
    for key in ("slug", "title", "meta_description", "body"):
        if key not in payload or not str(payload[key]).strip():
            print(f"ERROR: missing required field '{key}'", file=sys.stderr)
            sys.exit(1)

    slug = slugify(payload["slug"])
    date_str = datetime.date.today().isoformat()
    filename = f"{date_str}-{slug}.docx"

    os.makedirs(BLOGS_DIR, exist_ok=True)
    out_path = os.path.join(BLOGS_DIR, filename)

    build_docx(payload, out_path)
    print(f"Saved .docx to {out_path}")

    result = send_email(payload, out_path, filename)
    print(f"Email sent via Resend. Response: {json.dumps(result)}")


if __name__ == "__main__":
    main()
