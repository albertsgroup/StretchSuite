"""
Reddit r/Stretching - Assisted Stretching Research Scraper
Requires: pip install praw pandas

Setup:
  1. Go to https://www.reddit.com/prefs/apps
  2. Click "create another app" -> choose "script"
  3. Fill in any name/description, set redirect URI to http://localhost:8080
  4. Copy the client_id (under the app name) and client_secret
  5. Fill in the constants below and run: python reddit_assisted_stretching.py
"""

import praw
import pandas as pd
import json
import datetime

# ── Configuration ────────────────────────────────────────────────────────────
CLIENT_ID     = "YOUR_CLIENT_ID"
CLIENT_SECRET = "YOUR_CLIENT_SECRET"
USER_AGENT    = "assisted_stretching_research/1.0 by YOUR_REDDIT_USERNAME"

SUBREDDIT     = "Stretching"
SEARCH_TERMS  = [
    "assisted stretching",
    "assisted stretch",
    "stretching service",
    "stretch therapist",
    "StretchLab",
    "assisted flexibility",
]
POST_LIMIT    = 100   # per search term (max 100 per Reddit API call)
SORT          = "top" # top | new | relevance | hot | comments
# ─────────────────────────────────────────────────────────────────────────────


def scrape():
    reddit = praw.Reddit(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        user_agent=USER_AGENT,
    )

    posts_data = []
    comments_data = []
    seen_post_ids = set()

    sub = reddit.subreddit(SUBREDDIT)

    for term in SEARCH_TERMS:
        print(f"\nSearching: '{term}' ...")
        results = sub.search(term, limit=POST_LIMIT, sort=SORT)

        for post in results:
            if post.id in seen_post_ids:
                continue
            seen_post_ids.add(post.id)

            created = datetime.datetime.utcfromtimestamp(post.created_utc).isoformat()

            posts_data.append({
                "post_id":       post.id,
                "title":         post.title,
                "author":        str(post.author),
                "score":         post.score,
                "upvote_ratio":  post.upvote_ratio,
                "num_comments":  post.num_comments,
                "created_utc":   created,
                "url":           post.url,
                "permalink":     f"https://reddit.com{post.permalink}",
                "selftext":      post.selftext,
                "search_term":   term,
                "flair":         post.link_flair_text,
            })

            print(f"  [{post.score:>5}] {post.title[:80]}")

            # Fetch all comments
            try:
                post.comments.replace_more(limit=0)  # flatten MoreComments
                for comment in post.comments.list():
                    comment_created = datetime.datetime.utcfromtimestamp(
                        comment.created_utc
                    ).isoformat()
                    comments_data.append({
                        "comment_id":  comment.id,
                        "post_id":     post.id,
                        "post_title":  post.title,
                        "author":      str(comment.author),
                        "score":       comment.score,
                        "created_utc": comment_created,
                        "body":        comment.body,
                        "parent_id":   comment.parent_id,
                        "permalink":   f"https://reddit.com{comment.permalink}",
                    })
            except Exception as e:
                print(f"    Warning: could not fetch comments for {post.id}: {e}")

    # ── Save results ──────────────────────────────────────────────────────────
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    posts_file    = f"assisted_stretching_posts_{timestamp}.csv"
    comments_file = f"assisted_stretching_comments_{timestamp}.csv"
    json_file     = f"assisted_stretching_full_{timestamp}.json"

    posts_df    = pd.DataFrame(posts_data)
    comments_df = pd.DataFrame(comments_data)

    posts_df.to_csv(posts_file, index=False)
    comments_df.to_csv(comments_file, index=False)

    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(
            {"posts": posts_data, "comments": comments_data},
            f,
            indent=2,
            ensure_ascii=False,
        )

    # ── Summary ───────────────────────────────────────────────────────────────
    print("\n" + "="*60)
    print(f"Done!")
    print(f"  Unique posts:  {len(posts_data)}")
    print(f"  Total comments: {len(comments_data)}")
    print(f"\nFiles saved:")
    print(f"  {posts_file}")
    print(f"  {comments_file}")
    print(f"  {json_file}")

    if not posts_df.empty:
        print("\nTop 10 posts by score:")
        top = posts_df.nlargest(10, "score")[["title", "score", "num_comments"]]
        print(top.to_string(index=False))


if __name__ == "__main__":
    scrape()
