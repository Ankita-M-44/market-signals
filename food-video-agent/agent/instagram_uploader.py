"""Upload videos to Instagram as Reels via the Instagram Graph API."""

import time
import requests

GRAPH_API_BASE = "https://graph.facebook.com/v22.0"
POLL_INTERVAL_SEC = 10
MAX_WAIT_SEC = 300


def upload_reel(
    access_token: str,
    account_id: str,
    video_url: str,
    caption: str,
    dry_run: bool = False,
) -> str:
    """
    Upload a video as an Instagram Reel.
    video_url must be a publicly accessible URL (e.g., a Notion file URL or CDN).
    Returns the media ID on success.
    """
    if dry_run:
        print(f"[DRY RUN] Would upload Instagram Reel: {caption[:60]}...")
        return "DRY_RUN_MEDIA_ID"

    # Step 1: Create a video media container
    container_url = f"{GRAPH_API_BASE}/{account_id}/media"
    container_resp = requests.post(
        container_url,
        data={
            "media_type": "REELS",
            "video_url": video_url,
            "caption": caption,
            "share_to_feed": "true",
            "access_token": access_token,
        },
        timeout=60,
    )
    container_resp.raise_for_status()
    container_id = container_resp.json()["id"]
    print(f"Instagram container created: {container_id}")

    # Step 2: Poll until the video is processed
    status_url = f"{GRAPH_API_BASE}/{container_id}"
    elapsed = 0
    while elapsed < MAX_WAIT_SEC:
        time.sleep(POLL_INTERVAL_SEC)
        elapsed += POLL_INTERVAL_SEC
        status_resp = requests.get(
            status_url,
            params={"fields": "status_code", "access_token": access_token},
            timeout=30,
        )
        status_resp.raise_for_status()
        status_code = status_resp.json().get("status_code")
        print(f"  Instagram processing status: {status_code} ({elapsed}s)")
        if status_code == "FINISHED":
            break
        if status_code == "ERROR":
            raise RuntimeError(f"Instagram video processing failed: {status_resp.json()}")
    else:
        raise TimeoutError("Instagram video processing timed out")

    # Step 3: Publish the container
    publish_url = f"{GRAPH_API_BASE}/{account_id}/media_publish"
    publish_resp = requests.post(
        publish_url,
        data={"creation_id": container_id, "access_token": access_token},
        timeout=30,
    )
    publish_resp.raise_for_status()
    media_id = publish_resp.json()["id"]
    print(f"Instagram Reel published: {media_id}")
    return media_id
