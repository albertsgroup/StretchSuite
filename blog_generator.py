"""
StretchSuite Blog Save & Notify
--------------------------------
Called by Claude Code after it has researched keywords and written the blog.
Accepts the post as JSON on stdin or via --json flag.

Usage:
  echo '<json>' | python blog_generator.py
  python blog_generator.py --json '<json>'

Required env var:
  RESEND_API_KEY  — Resend API key (defaults to the configured key below)

JSON schema expected:
  {
    "slug":             "url-friendly-slug",
    "title":            "Post title",
    "meta_description": "150-160 char description",
    "body":             "Full markdown body"
  }
"""

import os
import re
import sys
import json
import datetime
import pathlib
import requests

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb")
NOTIFY_EMAIL   = os.environ.get("NOTIFY_EMAIL", "hammad@albertsgroup.net")
FROM_EMAIL     = "onboarding@resend.dev"

BLOG_DIR = pathlib.Path(__file__).parent / "blogs"
BLOG_DIR.mkdir(exist_ok=True)


def save_blog(post: dict) -> pathlib.Path:
    today    = datetime.date.today().isoformat()
    slug     = re.sub(r"[^a-z0-9-]", "", post["slug"].lower().replace(" ", "-"))
    filename = BLOG_DIR / f"{today}-{slug}.md"

    frontmatter = f"""---
title: "{post['title']}"
date: "{today}"
meta_description: "{post['meta_description']}"
---

"""
    filename.write_text(frontmatter + post["body"], encoding="utf-8")
    print(f"Saved → {filename}")
    return filename


def send_notification(post: dict, filepath: pathlib.Path) -> None:
    html = f"""
<h2>New blog is ready to view</h2>
<p><strong>{post['title']}</strong></p>
<p>{post['meta_description']}</p>
<p>File: <code>{filepath.name}</code></p>
<p><em>Generated {datetime.date.today()} by Stretch Suite Blog Automation</em></p>
"""
    resp = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type":  "application/json",
        },
        json={
            "from":    FROM_EMAIL,
            "to":      NOTIFY_EMAIL,
            "subject": "New blog is ready to view",
            "html":    html,
        },
        timeout=15,
    )
    if resp.status_code in (200, 201):
        print(f"Email sent → {NOTIFY_EMAIL}")
    else:
        print(f"Email failed: {resp.status_code} {resp.text}", file=sys.stderr)


def main():
    if "--json" in sys.argv:
        idx  = sys.argv.index("--json")
        data = sys.argv[idx + 1]
    else:
        data = sys.stdin.read()

    post     = json.loads(data)
    filepath = save_blog(post)
    send_notification(post, filepath)


if __name__ == "__main__":
    main()
