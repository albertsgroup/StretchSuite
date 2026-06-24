"""
StretchSuite Blog Automation
----------------------------
1. Researches high-intent keywords for assisted stretching in Oswego, NY
2. Plans and writes a full SEO + AEO optimised blog post using Claude
3. Saves the post to /blogs/ as a Markdown file
4. Sends a Resend email notification: "new blog is ready to view"

Required env vars (set in .env or your shell):
  ANTHROPIC_API_KEY   — Claude API key
  RESEND_API_KEY      — Resend API key
  NOTIFY_EMAIL        — recipient address for the notification

Run manually:  python blog_generator.py
Scheduled:     every Monday at 9 AM via cron / Claude Code cron
"""

import os
import re
import json
import datetime
import pathlib
import requests
import anthropic

# ── Config ────────────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
RESEND_API_KEY    = os.environ.get("RESEND_API_KEY", "re_UZ6bX6yR_BLL4pdrww5KM1vJnJwQGi2tb")
NOTIFY_EMAIL      = os.environ.get("NOTIFY_EMAIL", "hammad@albertsgroup.net")
FROM_EMAIL        = "onboarding@resend.dev"

BLOG_DIR = pathlib.Path(__file__).parent / "blogs"
BLOG_DIR.mkdir(exist_ok=True)

BUSINESS = {
    "name":     "Stretch Suite",
    "type":     "assisted stretching studio",
    "location": "downtown Oswego, NY",
    "url":      "https://stretchsuite.com",
    "services": [
        "one-on-one assisted stretching",
        "flexibility coaching",
        "sports recovery stretching",
        "corporate wellness sessions",
        "senior mobility programs",
    ],
}

# Published blog slugs — used for internal linking suggestions
EXISTING_BLOGS = [
    {"title": "What Is Assisted Stretching?",          "slug": "/blog/what-is-assisted-stretching"},
    {"title": "Benefits of Regular Flexibility Work",  "slug": "/blog/benefits-of-flexibility"},
    {"title": "Stretch Suite vs DIY Stretching",       "slug": "/blog/stretch-suite-vs-diy"},
]
# ─────────────────────────────────────────────────────────────────────────────


def research_keywords(client: anthropic.Anthropic) -> dict:
    """Ask Claude (with web_search tool) for fresh high-intent keywords."""
    print("🔍  Researching keywords...")

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=2048,
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
        messages=[{
            "role": "user",
            "content": (
                "You are an SEO strategist for a local assisted-stretching studio called "
                f"'{BUSINESS['name']}' in {BUSINESS['location']}. "
                "Search the web for the most recent, high-intent keywords people use when "
                "looking for assisted stretching, sports recovery, flexibility coaching, and "
                "related services — especially in the Oswego / Central New York area. "
                "Also look for trending questions (People Also Ask style) and AEO-friendly "
                "long-tail phrases. Return a JSON object with keys: "
                "'primary_keyword', 'secondary_keywords' (list of 5), "
                "'longtail_questions' (list of 5), 'trending_topic' (one sentence)."
            ),
        }],
    )

    # Extract the last text block
    text = next(
        (b.text for b in reversed(response.content) if hasattr(b, "text")),
        ""
    )
    # Pull JSON from the response
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # Fallback if parsing fails
    return {
        "primary_keyword":     "assisted stretching Oswego NY",
        "secondary_keywords":  [
            "flexibility coaching near me",
            "sports recovery stretching",
            "stretch therapy Oswego",
            "one-on-one stretching session",
            "senior mobility stretching",
        ],
        "longtail_questions": [
            "What is assisted stretching and how does it work?",
            "How often should I get assisted stretching?",
            "Is assisted stretching worth it for back pain?",
            "What should I expect at my first stretch session?",
            "How is assisted stretching different from yoga?",
        ],
        "trending_topic": "Growing demand for recovery-focused wellness services post-pandemic.",
    }


def generate_blog(client: anthropic.Anthropic, keywords: dict) -> dict:
    """Generate a full blog post with SEO + AEO structure and internal links."""
    print("✍️   Writing blog post...")

    internal_links_str = "\n".join(
        f'  - [{b["title"]}]({BUSINESS["url"]}{b["slug"]})'
        for b in EXISTING_BLOGS
    )

    prompt = f"""
You are a content strategist and SEO copywriter for {BUSINESS['name']},
an {BUSINESS['type']} located in {BUSINESS['location']}.

Services offered: {", ".join(BUSINESS['services'])}.

Write a complete, publish-ready blog post using the following keyword data:
- Primary keyword: {keywords['primary_keyword']}
- Secondary keywords: {", ".join(keywords['secondary_keywords'])}
- Longtail / AEO questions to answer: {json.dumps(keywords['longtail_questions'], indent=2)}
- Trending angle: {keywords['trending_topic']}

SEO & AEO requirements:
1. Title tag (H1) — include primary keyword naturally, under 60 chars.
2. Meta description — 150–160 chars, compelling, includes primary keyword.
3. Introduction — hook in 2–3 sentences, state the problem/value.
4. Use H2/H3 headings that mirror the longtail questions exactly (this feeds Google's
   People Also Ask and AI answer boxes).
5. Answer each question concisely in the first 2 sentences under its heading
   (AEO "direct answer" pattern), then expand.
6. Sprinkle secondary keywords naturally; never stuff.
7. Internal links — weave in at least 3 of these existing pages naturally:
{internal_links_str}
8. Local SEO — mention Oswego, NY and Central New York at least twice each.
9. CTA at the end — invite readers to book a session at {BUSINESS['name']}.
10. Word count: 900–1200 words.
11. Tone: friendly, knowledgeable, not salesy.

Return a JSON object with exactly these keys:
{{
  "slug": "url-friendly-slug",
  "title": "H1 title",
  "meta_description": "150-160 char description",
  "body": "full markdown body starting with the H1"
}}
"""

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # Fallback structure if JSON parse fails
    return {
        "slug":             "assisted-stretching-oswego-ny",
        "title":            "Assisted Stretching in Oswego, NY | Stretch Suite",
        "meta_description": "Discover the benefits of assisted stretching at Stretch Suite "
                            "in downtown Oswego, NY. Book your session today.",
        "body":             text,
    }


def save_blog(post: dict, keywords: dict) -> pathlib.Path:
    """Write the blog post to a dated Markdown file."""
    today   = datetime.date.today().isoformat()
    slug    = re.sub(r"[^a-z0-9-]", "", post["slug"].lower().replace(" ", "-"))
    filename = BLOG_DIR / f"{today}-{slug}.md"

    frontmatter = f"""---
title: "{post['title']}"
date: "{today}"
meta_description: "{post['meta_description']}"
primary_keyword: "{keywords['primary_keyword']}"
tags: [{", ".join(f'"{k}"' for k in keywords['secondary_keywords'])}]
---

"""
    filename.write_text(frontmatter + post["body"], encoding="utf-8")
    print(f"💾  Saved → {filename}")
    return filename


def send_notification(post: dict, filepath: pathlib.Path) -> bool:
    """Send a Resend email notifying that a new blog is ready."""
    print("📧  Sending email notification...")

    html = f"""
<h2>New blog is ready to view</h2>
<p><strong>{post['title']}</strong></p>
<p>{post['meta_description']}</p>
<p>File saved at: <code>{filepath.name}</code></p>
<p><em>Generated by Stretch Suite Blog Automation — {datetime.date.today()}</em></p>
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
        print(f"✅  Email sent → {NOTIFY_EMAIL}")
        return True
    else:
        print(f"❌  Email failed: {resp.status_code} {resp.text}")
        return False


def run():
    if not ANTHROPIC_API_KEY:
        raise EnvironmentError("ANTHROPIC_API_KEY is not set.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    keywords = research_keywords(client)
    print(f"   Primary keyword: {keywords['primary_keyword']}")

    post     = generate_blog(client, keywords)
    filepath = save_blog(post, keywords)

    send_notification(post, filepath)
    print("\n🎉  Done! Blog generation complete.")


if __name__ == "__main__":
    run()
