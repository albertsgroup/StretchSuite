"""
Stretch Suite blog generator.

Takes a JSON payload describing a blog post, renders it to a .docx file saved
under ./blogs/, and emails it as a base64 attachment via the Resend API.

Usage:
    python blog_generator.py --json '<json>'

Where <json> is:
    {
      "slug": "url-friendly-slug",
      "title": "H1 title",
      "meta_description": "150-160 char meta description",
      "body": "full markdown body"
    }

Optional flags:
    --ps            Append "PS: list completed!" to the email HTML body.
    --no-email      Only build the .docx, skip sending the email.
    --date YYYY-MM-DD  Override the date used in the filename (defaults to today).
"""

import argparse
import base64
import datetime
import json
import os
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.shared import Pt, RGBColor

RESEND_API_KEY = "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb"
RESEND_URL = "https://api.resend.com/emails"
FROM_ADDRESS = "Stretch Suite <hello@stretchsuite.com>"
RECIPIENTS = ["hammad@albertsgroup.net", "michelle@albertsgroup.net"]
SUBJECT = "New blog - Stretch Suite"

BLOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blogs")

INLINE_LINK = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD = re.compile(r"\*\*([^*]+)\*\*")


def _add_inline(paragraph, text):
    """Add text to a paragraph, rendering [label](url) links and **bold**."""
    # Split on links first, keeping the label as visible text.
    pos = 0
    for m in INLINE_LINK.finditer(text):
        _add_plain(paragraph, text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.font.underline = True
        run.font.color.rgb = RGBColor(0x1A, 0x5F, 0x8A)
        pos = m.end()
    _add_plain(paragraph, text[pos:])


def _add_plain(paragraph, text):
    if not text:
        return
    pos = 0
    for m in BOLD.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.bold = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def build_docx(title, meta_description, body, out_path):
    doc = Document()

    for raw in body.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            h = doc.add_heading(level=0)
            _add_inline(h, line[2:].strip())
        elif line.lstrip().startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            _add_inline(p, line.lstrip()[2:].strip())
        else:
            p = doc.add_paragraph()
            _add_inline(p, line.strip())

    doc.save(out_path)
    return out_path


def send_email(title, meta_description, docx_path, extra_html=""):
    with open(docx_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")

    html = (
        f"<h1>{title}</h1>"
        f"<p><strong>Meta description:</strong> {meta_description}</p>"
        f"<p>The full blog post is attached as a .docx file: "
        f"<strong>{os.path.basename(docx_path)}</strong>.</p>"
        f"{extra_html}"
    )

    payload = {
        "from": FROM_ADDRESS,
        "to": RECIPIENTS,
        "subject": SUBJECT,
        "html": html,
        "attachments": [
            {"filename": os.path.basename(docx_path), "content": encoded}
        ],
    }

    # Send via curl, which is already configured for this environment's
    # outbound proxy and CA bundle. The payload is passed through a temp file
    # to avoid argv size limits from the base64 attachment.
    with tempfile.NamedTemporaryFile(
        "w", suffix=".json", delete=False
    ) as tmp:
        json.dump(payload, tmp)
        tmp_path = tmp.name

    try:
        result = subprocess.run(
            [
                "curl", "-sS", "-w", "\n%{http_code}",
                "-X", "POST", RESEND_URL,
                "-H", f"Authorization: Bearer {RESEND_API_KEY}",
                "-H", "Content-Type: application/json",
                "--data", f"@{tmp_path}",
            ],
            capture_output=True, text=True, timeout=120,
        )
    finally:
        os.unlink(tmp_path)

    out = result.stdout.rsplit("\n", 1)
    body_resp = out[0] if len(out) == 2 else result.stdout
    status = out[1].strip() if len(out) == 2 else "000"
    try:
        status = int(status)
    except ValueError:
        status = 0
    if result.stderr:
        body_resp = f"{body_resp} | stderr: {result.stderr.strip()}"
    return status, body_resp


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="JSON payload for the post")
    parser.add_argument("--ps", action="store_true", help="Append list-completed PS")
    parser.add_argument("--no-email", action="store_true", help="Build docx only")
    parser.add_argument("--date", default=None, help="Override filename date")
    args = parser.parse_args()

    data = json.loads(args.json)
    slug = data["slug"]
    title = data["title"]
    meta_description = data["meta_description"]
    body = data["body"]

    date_str = args.date or datetime.date.today().isoformat()
    os.makedirs(BLOGS_DIR, exist_ok=True)
    filename = f"{date_str}-{slug}.docx"
    out_path = os.path.join(BLOGS_DIR, filename)

    build_docx(title, meta_description, body, out_path)
    print(f"Saved docx: {out_path}")

    if args.no_email:
        print("Skipping email (--no-email).")
        return

    extra_html = "<p>PS: list completed!</p>" if args.ps else ""
    status, resp = send_email(title, meta_description, out_path, extra_html)
    print(f"Resend status: {status}")
    print(f"Resend response: {resp}")
    if status not in (200, 201):
        sys.exit(1)


if __name__ == "__main__":
    main()
