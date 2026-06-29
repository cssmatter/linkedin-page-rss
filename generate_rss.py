import pandas as pd
import json
import os
import time
from email.utils import formatdate
import html
import uuid
from openai import OpenAI
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "source.xlsx")
STATE_FILE = os.path.join(BASE_DIR, "state.json")
RSS_FILE = os.path.join(BASE_DIR, "feed.xml")
POSTS_PER_RUN = 1
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY")
NVIDIA_MODEL = "meta/llama-3.1-405b-instruct" # Using a strong model from NVIDIA API, fallback to openai/gpt-oss-120b if needed, but llama-3.1-405b or 70b is great. Let's use meta/llama-3.1-70b-instruct to be safe on limits, or the one from main.py: openai/gpt-oss-120b
NVIDIA_MODEL = "openai/gpt-oss-120b"

# Feed metadata
FEED_TITLE = "Certifications Learning Platform - Daily Updates"
FEED_LINK = "https://www.linkedin.com"
FEED_DESC = "Daily SEO-friendly blog posts about top certification courses."

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_index": 0, "history": []}

def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=4)

def generate_blog_post(client, title, link):
    prompt = f"""You are an expert content creator and social media manager for a certifications learning platform.
Write a highly engaging, SEO-friendly LinkedIn blog post promoting the following course:

Course Title: {title}
Course Link: {link}

Requirements:
- Make it professional, persuasive, and highlight the career benefits of taking this certification.
- Include the course link naturally in the post.
- Always add a call to action near the end to download the mobile application for certification preparation using this link: https://apps.apple.com/us/app/certification-preparation/id6776616024
- Add plenty of relevant, high-traffic hashtags at the bottom (e.g. #certifications, #learning, #career, etc.).
- Output ONLY the final post content. Do not include any introductory remarks, conversational filler, or markdown code blocks (like ```).
"""
    try:
        response = client.chat.completions.create(
            model=NVIDIA_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=1024,
        )
        content = response.choices[0].message.content.strip()
        return content
    except Exception as e:
        print(f"Error generating post for {title}: {e}")
        return None

def generate_rss_xml(history):
    rss_items = ""
    # Sort history to have the newest first in the RSS
    for item in reversed(history[-20:]): # keep last 20 in the feed
        # Escape HTML entities for the CDATA or just use CDATA and replace newlines with <br>
        desc_html = item['description'].replace('\n', '<br>\n')
        rss_items += f"""
    <item>
      <title>{html.escape(item['title'])}</title>
      <link>{html.escape(item['link'])}</link>
      <description><![CDATA[{desc_html}]]></description>
      <pubDate>{item['pubDate']}</pubDate>
      <guid isPermaLink="false">{item['guid']}</guid>
    </item>"""

    rss_template = f"""<?xml version="1.0" encoding="UTF-8" ?>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom">
  <channel>
    <title>{html.escape(FEED_TITLE)}</title>
    <atom:link href="https://cssmatter.github.io/linkedin-page-rss/feed.xml" rel="self" type="application/rss+xml" />
    <link>{html.escape(FEED_LINK)}</link>
    <description>{html.escape(FEED_DESC)}</description>
    <language>en-us</language>
    <lastBuildDate>{formatdate(localtime=False)}</lastBuildDate>{rss_items}
  </channel>
</rss>
"""
    with open(RSS_FILE, "w", encoding="utf-8") as f:
        f.write(rss_template)

def main():
    print("Loading Excel data...")
    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return

    # Check columns
    if 'Course Title' not in df.columns or 'Course Link' not in df.columns:
        print("Excel file must contain 'Course Title' and 'Course Link' columns.")
        return

    if 'Status' not in df.columns:
        df['Status'] = ''

    state = load_state()
    
    # Filter pending courses (Status != 'done' or nan/empty)
    # We treat any status that is not explicitly 'done' (case-insensitive) as pending
    df['Status'] = df['Status'].fillna('')
    pending_mask = df['Status'].astype(str).str.lower().str.strip() != 'done'
    pending_indices = df[pending_mask].index.tolist()

    if not pending_indices:
        print("All courses have been processed!")
        return
    
    indices_to_process = pending_indices[:POSTS_PER_RUN]
    
    client = OpenAI(
        base_url="https://integrate.api.nvidia.com/v1",
        api_key=NVIDIA_API_KEY
    )

    print(f"Processing {len(indices_to_process)} pending courses...")
    processed_any = False

    for idx in indices_to_process:
        row = df.loc[idx]
        title = str(row['Course Title']).strip()
        link = str(row['Course Link']).strip()

        if not title or title == 'nan':
            # Mark invalid rows as done to skip them in future
            df.at[idx, 'Status'] = 'done'
            processed_any = True
            continue

        print(f"Generating post for: {title}")
        blog_content = generate_blog_post(client, title, link)
        
        if blog_content:
            # Generate RSS item details
            pub_date = formatdate(localtime=False)
            item_guid = str(uuid.uuid4())
            
            state["history"].append({
                "title": f"New Course: {title}",
                "link": link,
                "description": blog_content,
                "pubDate": pub_date,
                "guid": item_guid
            })
            
            # Update Excel Status
            df.at[idx, 'Status'] = 'done'
            processed_any = True
            print("Successfully generated post and marked as done.")
            time.sleep(2) # Avoid rate limits
    
    save_state(state)

    if processed_any:
        print("Saving updated status to Excel...")
        try:
            df.to_excel(EXCEL_FILE, index=False)
        except Exception as e:
            print(f"Error saving Excel file: {e}")

    print("Generating RSS XML...")
    generate_rss_xml(state["history"])
    print(f"Done! Feed saved to {RSS_FILE}")

if __name__ == "__main__":
    main()
