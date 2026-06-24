"""
Stretch Suite Blog Generator
Usage: python blog_generator.py --json '<json_string>'

JSON fields:
  slug             - url-friendly slug
  title            - H1 title
  meta_description - 150-160 char meta description
  body             - full markdown body

Saves .docx to /blogs/ and sends via Resend API.
"""

import argparse
import base64
import json
import os
import re
import sys
from datetime import date

import requests
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
FROM_EMAIL = "Stretch Suite <blogs@update.albertsgroup.net>"
TO_EMAILS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]

BLOGS_DIR = os.path.join(os.path.dirname(__file__), "blogs")


# ---------------------------------------------------------------------------
# Markdown -> python-docx converter (handles H1/H2/H3, paragraphs, lists,
# inline bold/italic/links)
# ---------------------------------------------------------------------------

def _apply_inline(run_adder, text):
    """Parse inline **bold**, *italic*, and [link](url) within a paragraph."""
    pattern = re.compile(r'\*\*(.+?)\*\*|\*(.+?)\*|\[([^\]]+)\]\(([^)]+)\)')
    last = 0
    segments = []
    for m in pattern.finditer(text):
        if m.start() > last:
            segments.append(('text', text[last:m.start()]))
        if m.group(1) is not None:
            segments.append(('bold', m.group(1)))
        elif m.group(2) is not None:
            segments.append(('italic', m.group(2)))
        else:
            segments.append(('link', m.group(3), m.group(4)))
        last = m.end()
    if last < len(text):
        segments.append(('text', text[last:]))

    for seg in segments:
        if seg[0] == 'bold':
            run = run_adder(seg[1])
            run.bold = True
        elif seg[0] == 'italic':
            run = run_adder(seg[1])
            run.italic = True
        elif seg[0] == 'link':
            run = run_adder(seg[1])
            run.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
            run.underline = True
        else:
            run_adder(seg[0])


def markdown_to_docx(md_text, title, meta_description):
    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # Meta description block at top (italic, grey)
    meta_para = doc.add_paragraph()
    meta_run = meta_para.add_run(f"Meta description: {meta_description}")
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_paragraph()

    lines = md_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()

        # H1
        if line.startswith('# '):
            heading = line[2:].strip()
            para = doc.add_heading(level=1)
            para.clear()
            _apply_inline(para.add_run, heading)

        # H2
        elif line.startswith('## '):
            heading = line[3:].strip()
            para = doc.add_heading(level=2)
            para.clear()
            _apply_inline(para.add_run, heading)

        # H3
        elif line.startswith('### '):
            heading = line[4:].strip()
            para = doc.add_heading(level=3)
            para.clear()
            _apply_inline(para.add_run, heading)

        # Unordered list item
        elif line.startswith('- ') or line.startswith('* '):
            item_text = line[2:].strip()
            para = doc.add_paragraph(style='List Bullet')
            para.clear()
            _apply_inline(para.add_run, item_text)

        # Blank line — skip
        elif line.strip() == '':
            pass

        # Normal paragraph
        else:
            para = doc.add_paragraph()
            _apply_inline(para.add_run, line)

        i += 1

    return doc


def main():
    parser = argparse.ArgumentParser(description='Generate a blog .docx and email it.')
    parser.add_argument('--json', required=True, help='JSON string with blog data')
    args = parser.parse_args()

    try:
        data = json.loads(args.json)
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}", file=sys.stderr)
        sys.exit(1)

    slug = data['slug']
    title = data['title']
    meta_description = data['meta_description']
    body = data['body']

    today = date.today().strftime('%Y-%m-%d')
    filename = f"{today}-{slug}.docx"
    filepath = os.path.join(BLOGS_DIR, filename)

    os.makedirs(BLOGS_DIR, exist_ok=True)

    print(f"Generating .docx: {filename}")
    doc = markdown_to_docx(body, title, meta_description)
    doc.save(filepath)
    print(f"Saved to {filepath}")

    # Base64 encode for email attachment
    with open(filepath, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('utf-8')

    html_body = f"""
<html>
<body>
<h2>{title}</h2>
<p><strong>Meta description:</strong> {meta_description}</p>
<p>The latest Stretch Suite blog post is attached as a Word document (<code>{filename}</code>).</p>
<p>Please review, format for the website, and publish when ready.</p>
</body>
</html>
"""

    payload = {
        "from": FROM_EMAIL,
        "to": TO_EMAILS,
        "subject": "New blog - Stretch Suite",
        "html": html_body,
        "attachments": [
            {
                "filename": filename,
                "content": encoded,
            }
        ]
    }

    print("Sending email via Resend...")
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
        print(f"Email send failed ({resp.status_code}): {resp.text}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
