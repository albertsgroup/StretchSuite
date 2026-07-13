#!/usr/bin/env python3
"""
Stretch Suite weekly blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file saved
in the ./blogs/ folder, and emails it as a base64-encoded attachment via the
Resend API.

Usage:
    python blog_generator.py --json '{"slug": "...", "title": "...",
        "meta_description": "...", "body": "full markdown body"}'
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
RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
RECIPIENTS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")


def build_docx(title, meta_description, body, out_path):
    """Render the markdown-ish body into a clean .docx document."""
    doc = Document()

    # Title (H1)
    heading = doc.add_heading(title, level=0)

    # Meta description block
    meta_label = doc.add_paragraph()
    run = meta_label.add_run("Meta description: ")
    run.bold = True
    meta_label.add_run(meta_description)
    for r in meta_label.runs:
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()  # spacer

    for raw_line in body.split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            continue

        # Skip a duplicated H1 if the body repeats the title
        if line.startswith("# "):
            text = line[2:].strip()
            if text.strip().lower() == title.strip().lower():
                continue
            doc.add_heading(_clean(text), level=1)
        elif line.startswith("### "):
            doc.add_heading(_clean(line[4:].strip()), level=3)
        elif line.startswith("## "):
            doc.add_heading(_clean(line[3:].strip()), level=2)
        elif line.startswith("- ") or line.startswith("* "):
            _add_formatted(doc.add_paragraph(style="List Bullet"), line[2:].strip())
        else:
            _add_formatted(doc.add_paragraph(), line.strip())

    doc.save(out_path)
    return out_path


_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
_BOLD_RE = re.compile(r"\*\*([^*]+)\*\*")


def _clean(text):
    """Strip markdown link/bold syntax for headings."""
    text = _LINK_RE.sub(r"\1", text)
    text = _BOLD_RE.sub(r"\1", text)
    return text


def _add_formatted(paragraph, text):
    """Add inline text to a paragraph, resolving **bold** and [links](url)."""
    # First resolve links to their display text (docx plain runs), keeping URL
    # inline in parentheses so it survives in the document.
    text = _LINK_RE.sub(lambda m: f"{m.group(1)} ({m.group(2)})", text)

    pos = 0
    for match in _BOLD_RE.finditer(text):
        if match.start() > pos:
            paragraph.add_run(text[pos:match.start()])
        run = paragraph.add_run(match.group(1))
        run.bold = True
        pos = match.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def send_email(title, meta_description, docx_path):
    """Send the blog as a .docx attachment via Resend."""
    with open(docx_path, "rb") as fh:
        encoded = base64.b64encode(fh.read()).decode("utf-8")

    html = f"""
    <h2>{title}</h2>
    <p><strong>Meta description:</strong> {meta_description}</p>
    <p>This week's Stretch Suite blog post is ready. The full post is attached
    as a Word document ({os.path.basename(docx_path)}).</p>
    <p>Move Better. Feel Better. Stay Independent.</p>
    """

    payload = {
        "from": FROM_ADDRESS,
        "to": RECIPIENTS,
        "subject": "New blog - Stretch Suite",
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
        json=payload,
        timeout=60,
    )
    return resp


def main():
    parser = argparse.ArgumentParser(description="Generate and email a blog post.")
    parser.add_argument("--json", required=True, help="JSON payload for the blog post")
    args = parser.parse_args()

    try:
        data = json.loads(args.json)
    except json.JSONDecodeError as exc:
        print(f"ERROR: invalid JSON payload: {exc}", file=sys.stderr)
        sys.exit(1)

    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    os.makedirs(BLOGS_DIR, exist_ok=True)
    date = datetime.date.today().isoformat()
    filename = f"{date}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    build_docx(title, meta_description, body, out_path)
    print(f"Saved: {out_path}")

    # Also save the markdown alongside for the repo record.
    md_path = os.path.join(BLOGS_DIR, f"{date}-{slug}.md")
    with open(md_path, "w") as fh:
        fh.write(f"<!-- meta description: {meta_description} -->\n\n{body}\n")
    print(f"Saved: {md_path}")

    resp = send_email(title, meta_description, out_path)
    if resp.status_code in (200, 201):
        print(f"Email sent successfully: {resp.json()}")
    else:
        print(f"ERROR sending email: {resp.status_code} {resp.text}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
