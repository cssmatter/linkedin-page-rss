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
POSTS_PER_RUN = 1
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY")
NVIDIA_MODEL = "openai/gpt-oss-120b"

CONFIGS = [
    {
        "excel_file": os.path.join(BASE_DIR, "source.xlsx"),
        "state_file": os.path.join(BASE_DIR, "state.json"),
        "rss_file": os.path.join(BASE_DIR, "feed.xml"),
        "feed_title": "Certifications Learning Platform - Daily Updates",
        "feed_link": "https://www.linkedin.com",
        "feed_desc": "Daily SEO-friendly blog posts about top certification courses.",
        "rss_url": "https://cssmatter.github.io/linkedin-page-rss/feed.xml"
    },
    {
        "excel_file": os.path.join(BASE_DIR, "courses - interview practice test.xlsx"),
        "state_file": os.path.join(BASE_DIR, "state_interview.json"),
        "rss_file": os.path.join(BASE_DIR, "feed_interview.xml"),
        "feed_title": "Interview Practice Tests - Daily Updates",
        "feed_link": "https://www.linkedin.com",
        "feed_desc": "Daily SEO-friendly blog posts about top interview practice tests.",
        "rss_url": "https://cssmatter.github.io/linkedin-page-rss/feed_interview.xml"
    }
]

def load_state(state_file):
    if os.path.exists(state_file):
        with open(state_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_index": 0, "history": []}

def save_state(state, state_file):
    with open(state_file, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=4)

def generate_blog_post(client, title, link, is_interview=False):
    if is_interview:
        prompt = f"""You are an expert content creator and social media manager for a learning platform.
Write a highly engaging, SEO-friendly LinkedIn blog post promoting the following interview practice test:

Test Title: {title}
Test Link: {link}

Requirements:
- Make it professional, persuasive, and highlight the career benefits of preparing for interviews with this test.
- Include the test link naturally in the post.
- Always add a call to action near the end to download the mobile application for interview preparation using this link: https://apps.apple.com/us/app/interview-preparation-app/id123456789 (Use the correct app link if provided, otherwise you can use a generic CTA).
- Add plenty of relevant, high-traffic hashtags at the bottom (e.g. #interviewtips, #career, #hiring, etc.).
- Output ONLY the final post content. Do not include any introductory remarks, conversational filler, or markdown code blocks (like ```).
"""
    else:
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

def generate_rss_xml(history, rss_file, feed_title, feed_link, feed_desc, rss_url):
    rss_items = ""
    # Sort history to have the newest first in the RSS
    for item in reversed(history[-20:]): # keep last 20 in the feed
        # Provide plain text in description and HTML in content:encoded for better compatibility
        desc_plain = item['description']
        desc_html = item['description'].replace('\n', '<br>\n')
        rss_items += f"""
    <item>
      <title>{html.escape(item['title'])}</title>
      <link>{html.escape(item['link'])}</link>
      <description><![CDATA[{desc_plain}]]></description>
      <content:encoded><![CDATA[{desc_html}]]></content:encoded>
      <pubDate>{item['pubDate']}</pubDate>
      <guid isPermaLink="false">{item['guid']}</guid>
    </item>"""

    rss_template = f"""<?xml version="1.0" encoding="UTF-8" ?>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom" xmlns:content="http://purl.org/rss/1.0/modules/content/">
  <channel>
    <title>{html.escape(feed_title)}</title>
    <atom:link href="{html.escape(rss_url)}" rel="self" type="application/rss+xml" />
    <link>{html.escape(feed_link)}</link>
    <description>{html.escape(feed_desc)}</description>
    <language>en-us</language>
    <lastBuildDate>{formatdate(localtime=False)}</lastBuildDate>{rss_items}
  </channel>
</rss>
"""
    with open(rss_file, "w", encoding="utf-8") as f:
        f.write(rss_template)

def process_feed(config, client):
    print(f"Processing feed for: {os.path.basename(config['excel_file'])}")
    try:
        df = pd.read_excel(config["excel_file"])
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return

    # Check columns
    if 'Course Title' not in df.columns or 'Course Link' not in df.columns:
        print("Excel file must contain 'Course Title' and 'Course Link' columns.")
        return

    if 'Status' not in df.columns:
        df['Status'] = ''

    state = load_state(config["state_file"])
    
    # Filter pending courses (Status != 'done' or nan/empty)
    # We treat any status that is not explicitly 'done' (case-insensitive) as pending
    df['Status'] = df['Status'].fillna('')
    pending_mask = df['Status'].astype(str).str.lower().str.strip() != 'done'
    pending_indices = df[pending_mask].index.tolist()

    if not pending_indices:
        print("All courses have been processed!")
        return
    
    indices_to_process = pending_indices[:POSTS_PER_RUN]
    print(f"Processing {len(indices_to_process)} pending items...")
    processed_any = False
    
    is_interview = "interview" in config["excel_file"].lower()

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
        blog_content = generate_blog_post(client, title, link, is_interview=is_interview)
        
        if blog_content:
            # Generate RSS item details
            pub_date = formatdate(localtime=False)
            item_guid = str(uuid.uuid4())
            
            prefix = "New Interview Practice Test" if is_interview else "New Course"
            state["history"].append({
                "title": f"{prefix}: {title}",
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
    
    save_state(state, config["state_file"])

    if processed_any:
        print("Saving updated status to Excel...")
        try:
            df.to_excel(config["excel_file"], index=False)
        except Exception as e:
            print(f"Error saving Excel file: {e}")

    print("Generating RSS XML...")
    generate_rss_xml(
        state["history"], 
        config["rss_file"], 
        config["feed_title"], 
        config["feed_link"], 
        config["feed_desc"], 
        config["rss_url"]
    )
    print(f"Done! Feed saved to {config['rss_file']}\n")

def main():
    client = OpenAI(
        base_url="https://integrate.api.nvidia.com/v1",
        api_key=NVIDIA_API_KEY
    )
    
    for config in CONFIGS:
        process_feed(config, client)

if __name__ == "__main__":
    main()
