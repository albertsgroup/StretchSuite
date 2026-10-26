"""
StretchSuite Blog Save & Notify
--------------------------------
Accepts a blog post as JSON, saves it as a .docx file, and sends a Resend
email notification with the file attached.

Usage:
  python blog_generator.py --json '<json>'

JSON schema:
  {
    "slug":             "url-friendly-slug",
    "title":            "Post title",
    "meta_description": "150-160 char description",
    "body":             "Full markdown body"
  }

Required env var:
  RESEND_API_KEY  (no fallback — must be set in the environment, never in code)
"""

import os
import re
import sys
import json
import base64
import datetime
import pathlib
import requests
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

RESEND_API_KEY = os.environ.get("RESEND_API_KEY")
NOTIFY_EMAILS  = os.environ.get("NOTIFY_EMAILS", "hammad@albertsgroup.net,michelle@albertsgroup.net").split(",")
FROM_EMAIL     = "Stretch Suite <hello@stretchsuite.com>"

BLOG_DIR = pathlib.Path(__file__).parent / "blogs"
BLOG_DIR.mkdir(exist_ok=True)


def markdown_to_docx(post: dict) -> pathlib.Path:
    """Convert a markdown blog post to a formatted .docx file."""
    today    = datetime.date.today().isoformat()
    slug     = re.sub(r"[^a-z0-9-]", "", post["slug"].lower().replace(" ", "-"))
    filepath = BLOG_DIR / f"{today}-{slug}.docx"

    doc = Document()

    # ── Document title style ─────────────────────────────────────────────────
    title_para = doc.add_heading(post["title"], level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Meta description block
    meta = doc.add_paragraph()
    meta_run = meta.add_run(f"Meta description: {post['meta_description']}")
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    meta_run.font.size = Pt(10)

    doc.add_paragraph(f"Published: {today}").runs[0].font.size = Pt(10)
    doc.add_paragraph()  # spacer

    # ── Parse and render markdown body ───────────────────────────────────────
    for line in post["body"].split("\n"):
        stripped = line.strip()

        if stripped.startswith("### "):
            doc.add_heading(stripped[4:], level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:], level=1)
        elif stripped.startswith("- ") or stripped.startswith("* "):
            doc.add_paragraph(stripped[2:], style="List Bullet")
        elif re.match(r"^\d+\.", stripped):
            doc.add_paragraph(re.sub(r"^\d+\.\s*", "", stripped), style="List Number")
        elif stripped == "---" or stripped == "***":
            doc.add_paragraph("─" * 60)
        elif stripped == "":
            doc.add_paragraph()
        else:
            # Render inline bold/italic/links as plain text (Word handles links separately)
            clean = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", stripped)  # strip markdown links
            para = doc.add_paragraph()
            # Handle **bold** and *italic* inline
            parts = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", clean)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = para.add_run(part[2:-2])
                    run.bold = True
                elif part.startswith("*") and part.endswith("*"):
                    run = para.add_run(part[1:-1])
                    run.italic = True
                else:
                    para.add_run(part)

    doc.save(filepath)
    print(f"Saved → {filepath}")
    return filepath


def send_notification(post: dict, filepath: pathlib.Path) -> None:
    """Email the .docx as an attachment via Resend."""
    if not RESEND_API_KEY:
        print("Email failed: RESEND_API_KEY is not set in the environment", file=sys.stderr)
        sys.exit(1)

    print(f"Sending email → {', '.join(NOTIFY_EMAILS)}")

    with open(filepath, "rb") as f:
        attachment_b64 = base64.b64encode(f.read()).decode()

    html = f"""
<h2>New blog is ready to view</h2>
<p><strong>{post['title']}</strong></p>
<p>{post['meta_description']}</p>
<p>The blog post is attached as a Word document.</p>
<p><em>Generated {datetime.date.today()} — Stretch Suite Blog Automation</em></p>
"""

    resp = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type":  "application/json",
        },
        json={
            "from":    FROM_EMAIL,
            "to":      NOTIFY_EMAILS,
            "subject": "New blog - Stretch Suite",
            "html":    html,
            "attachments": [{
                "filename": filepath.name,
                "content":  attachment_b64,
            }],
        },
        timeout=15,
    )

    if resp.status_code in (200, 201):
        print(f"Email sent ✓")
    else:
        print(f"Email failed: {resp.status_code} {resp.text}", file=sys.stderr)


def main():
    if "--json" in sys.argv:
        idx  = sys.argv.index("--json")
        data = sys.argv[idx + 1]
    else:
        data = sys.stdin.read()

    post     = json.loads(data)
    filepath = markdown_to_docx(post)
    send_notification(post, filepath)


if __name__ == "__main__":
    main()
