#!/usr/bin/env python3
"""
Mini Food Video Uploader — Step 5
Picks the next 'Created' video from Notion and posts it to Instagram + YouTube.

Usage:
  python uploader.py             # upload next pending video
  python uploader.py --dry-run   # print what would happen, no uploads
"""

import os
import sys
import argparse
import tempfile
import requests
from pathlib import Path

from agent.notion_tracker import get_next_created_entry, update_status
from agent.instagram_uploader import upload_reel
from agent.youtube_uploader import upload_short


def load_env() -> dict:
    return {
        "notion_key": os.environ["NOTION_API_KEY"],
        "notion_db_id": os.environ.get(
            "NOTION_DATABASE_ID", "38d48c07-8aa9-8029-8bd3-000b11e75d47"
        ),
        "ig_token": os.environ["INSTAGRAM_ACCESS_TOKEN"],
        "ig_account_id": os.environ["INSTAGRAM_ACCOUNT_ID"],
        "yt_client_id": os.environ["YOUTUBE_CLIENT_ID"],
        "yt_client_secret": os.environ["YOUTUBE_CLIENT_SECRET"],
        "yt_refresh_token": os.environ["YOUTUBE_REFRESH_TOKEN"],
    }


def extract_page_data(page: dict) -> dict:
    """Extract name, caption, and file URL from a Notion page result."""
    props = page.get("properties", {})

    name = ""
    title_list = props.get("Name", {}).get("title", [])
    if title_list:
        name = title_list[0].get("plain_text", "")

    caption = ""
    rt_list = props.get("Caption", {}).get("rich_text", [])
    if rt_list:
        caption = rt_list[0].get("plain_text", "")

    # Files & media — Notion returns file objects
    file_url = None
    files = props.get("Files & media", {}).get("files", [])
    if files:
        file_obj = files[0]
        if file_obj.get("type") == "file":
            file_url = file_obj["file"]["url"]
        elif file_obj.get("type") == "external":
            file_url = file_obj["external"]["url"]

    return {"name": name, "caption": caption, "file_url": file_url}


def download_video(url: str, dest_path: str) -> None:
    """Download a video from a URL to a local path."""
    resp = requests.get(url, timeout=120, stream=True)
    resp.raise_for_status()
    with open(dest_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)


def main(dry_run: bool = False) -> None:
    print("=== Mini Food Video Uploader ===")
    print(f"Mode: {'DRY RUN' if dry_run else 'LIVE'}\n")

    env = load_env()

    # --- Find the next video to upload ---
    print("Querying Notion for next 'Created' video...")
    if dry_run:
        page = {
            "id": "DRY_RUN_PAGE",
            "properties": {
                "Name": {"title": [{"plain_text": "Test Dish"}]},
                "Caption": {"rich_text": [{"plain_text": "Tiny perfection. #miniaturecooking #asmr"}]},
                "Files & media": {"files": []},
            },
        }
    else:
        page = get_next_created_entry(env["notion_key"], env["notion_db_id"])

    if page is None:
        print("No videos with status 'Created' found. Nothing to upload.")
        sys.exit(0)

    page_id = page["id"]
    data = extract_page_data(page)
    print(f"Found: '{data['name']}' (page {page_id})")

    # --- Download the video locally ---
    with tempfile.TemporaryDirectory() as tmpdir:
        video_path = str(Path(tmpdir) / "video.mp4")

        if dry_run:
            print("[DRY RUN] Skipping video download")
            # Create placeholder
            with open(video_path, "wb") as f:
                f.write(b"DRY_RUN")
        elif data["file_url"]:
            print(f"Downloading video from Notion...")
            download_video(data["file_url"], video_path)
            print(f"Video downloaded ({Path(video_path).stat().st_size // 1024} KB)")
        else:
            print("ERROR: No video file attached to this Notion entry.")
            sys.exit(1)

        caption = data["caption"]
        title = data["name"]

        # --- Upload to Instagram as Reel ---
        print("\nUploading to Instagram Reels...")
        # Instagram requires a public URL; we use the Notion file URL directly
        ig_video_url = data["file_url"] if not dry_run else "https://example.com/dry-run.mp4"
        ig_media_id = upload_reel(
            access_token=env["ig_token"],
            account_id=env["ig_account_id"],
            video_url=ig_video_url,
            caption=caption,
            dry_run=dry_run,
        )
        print(f"Instagram Reel ID: {ig_media_id}")

        # --- Upload to YouTube as Short ---
        print("\nUploading to YouTube Shorts...")
        yt_video_id = upload_short(
            client_id=env["yt_client_id"],
            client_secret=env["yt_client_secret"],
            refresh_token=env["yt_refresh_token"],
            video_path=video_path,
            title=f"Miniature {title} ASMR Cooking",
            description=caption,
            dry_run=dry_run,
        )
        print(f"YouTube Short ID: {yt_video_id}")

        # --- Update Notion status ---
        if not dry_run:
            update_status(env["notion_key"], page_id, "Uploaded")
            print(f"\nNotion status updated to 'Uploaded' for page {page_id}")
        else:
            print(f"\n[DRY RUN] Would mark Notion page {page_id} as 'Uploaded'")

    print("\n=== Upload complete ===")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Upload next pending video to Instagram + YouTube")
    parser.add_argument("--dry-run", action="store_true", help="Skip all uploads")
    args = parser.parse_args()
    main(dry_run=args.dry_run)
