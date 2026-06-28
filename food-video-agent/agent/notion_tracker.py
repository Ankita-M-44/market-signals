"""CRUD operations for the Notion Videos database."""

import os
import json
import requests
from datetime import date, timedelta

NOTION_API_BASE = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"

# Scheduled posting times offset from today (one video per day slot)
# Slot 0 = today, Slot 1 = today, Slot 2 = today (different upload times)
POST_DATE = date.today().isoformat()


def _headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def create_video_entry(
    api_key: str,
    database_id: str,
    dish: dict,
    caption: str,
    video_prompt: str,
    slot_index: int,
) -> str:
    """Create a new row in the Videos database. Returns the page ID."""
    url = f"{NOTION_API_BASE}/pages"
    payload = {
        "parent": {"database_id": database_id},
        "properties": {
            "Name": {"title": [{"text": {"content": dish["dish_name"]}}]},
            "Caption": {"rich_text": [{"text": {"content": caption}}]},
            "Date": {"date": {"start": POST_DATE}},
            "Status": {"status": {"name": "Not started"}},
        },
        # Store the prompt in the page body for reference
        "children": [
            {
                "object": "block",
                "type": "heading_3",
                "heading_3": {
                    "rich_text": [{"type": "text", "text": {"content": "Video Prompt"}}]
                },
            },
            {
                "object": "block",
                "type": "paragraph",
                "paragraph": {
                    "rich_text": [{"type": "text", "text": {"content": video_prompt}}]
                },
            },
            {
                "object": "block",
                "type": "heading_3",
                "heading_3": {
                    "rich_text": [{"type": "text", "text": {"content": "Dish Details"}}]
                },
            },
            {
                "object": "block",
                "type": "paragraph",
                "paragraph": {
                    "rich_text": [
                        {
                            "type": "text",
                            "text": {
                                "content": f"Cuisine: {dish['cuisine']}\nIngredients: {', '.join(dish['key_ingredients'])}\nSlot: {slot_index}"
                            },
                        }
                    ]
                },
            },
        ],
    }
    resp = requests.post(url, headers=_headers(api_key), json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def update_status(api_key: str, page_id: str, status: str) -> None:
    """Update the Status property of a Notion page. status: 'Not started' | 'Created' | 'Uploaded'"""
    url = f"{NOTION_API_BASE}/pages/{page_id}"
    payload = {"properties": {"Status": {"status": {"name": status}}}}
    resp = requests.patch(url, headers=_headers(api_key), json=payload, timeout=30)
    resp.raise_for_status()


def upload_video_file(api_key: str, page_id: str, video_path: str, filename: str) -> None:
    """Upload a local video file to Notion's Files & media property via the files upload API."""
    # Step 1: Create an upload session
    session_url = f"{NOTION_API_BASE}/file-uploads"
    session_resp = requests.post(
        session_url,
        headers=_headers(api_key),
        json={"filename": filename, "content_type": "video/mp4"},
        timeout=30,
    )
    session_resp.raise_for_status()
    session_data = session_resp.json()
    upload_url = session_data["upload_url"]
    file_upload_id = session_data["id"]

    # Step 2: Upload the file bytes
    with open(video_path, "rb") as f:
        upload_resp = requests.put(
            upload_url,
            data=f,
            headers={"Content-Type": "video/mp4"},
            timeout=120,
        )
        upload_resp.raise_for_status()

    # Step 3: Attach the uploaded file to the page property
    attach_url = f"{NOTION_API_BASE}/pages/{page_id}"
    payload = {
        "properties": {
            "Files & media": {
                "files": [
                    {
                        "type": "file",
                        "name": filename,
                        "file": {"id": file_upload_id},
                    }
                ]
            }
        }
    }
    attach_resp = requests.patch(
        attach_url, headers=_headers(api_key), json=payload, timeout=30
    )
    attach_resp.raise_for_status()


def get_next_created_entry(api_key: str, database_id: str) -> dict | None:
    """Query the database for the oldest entry with Status='Created'. Returns page properties dict or None."""
    url = f"{NOTION_API_BASE}/databases/{database_id}/query"
    payload = {
        "filter": {"property": "Status", "status": {"equals": "Created"}},
        "sorts": [{"property": "Date", "direction": "ascending"}],
        "page_size": 1,
    }
    resp = requests.post(url, headers=_headers(api_key), json=payload, timeout=30)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    return results[0] if results else None


def fetch_master_prompt_page(api_key: str, page_id: str) -> str:
    """Fetch the raw block content of the Mini food page and return plain text."""
    url = f"{NOTION_API_BASE}/blocks/{page_id}/children"
    resp = requests.get(url, headers=_headers(api_key), timeout=30)
    resp.raise_for_status()
    blocks = resp.json().get("results", [])
    lines = []
    for block in blocks:
        btype = block.get("type", "")
        content_obj = block.get(btype, {})
        rich_text = content_obj.get("rich_text", [])
        text = "".join(rt.get("plain_text", "") for rt in rich_text)
        if text:
            lines.append(text)
    return "\n".join(lines)
