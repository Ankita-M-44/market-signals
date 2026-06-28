"""Upload videos to YouTube as Shorts via the YouTube Data API v3."""

import os
import requests
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


YOUTUBE_SCOPES = ["https://www.googleapis.com/auth/youtube.upload"]


def _get_credentials(client_id: str, client_secret: str, refresh_token: str) -> Credentials:
    """Build Google OAuth2 credentials from stored refresh token."""
    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=YOUTUBE_SCOPES,
    )
    creds.refresh(Request())
    return creds


def upload_short(
    client_id: str,
    client_secret: str,
    refresh_token: str,
    video_path: str,
    title: str,
    description: str,
    dry_run: bool = False,
) -> str:
    """
    Upload a local video file as a YouTube Short.
    Adds #Shorts to the title to trigger the Shorts player.
    Returns the video ID on success.
    """
    if dry_run:
        print(f"[DRY RUN] Would upload YouTube Short: {title}")
        return "DRY_RUN_VIDEO_ID"

    creds = _get_credentials(client_id, client_secret, refresh_token)
    youtube = build("youtube", "v3", credentials=creds)

    # YouTube Shorts: video must be vertical and ≤60s; #Shorts in title/description triggers the player
    shorts_title = f"{title} #Shorts"
    body = {
        "snippet": {
            "title": shorts_title[:100],
            "description": f"{description}\n\n#Shorts #miniaturecooking #asmrcooking",
            "tags": ["shorts", "miniature cooking", "asmr", "food", "cooking"],
            "categoryId": "26",  # Howto & Style
        },
        "status": {"privacyStatus": "public", "selfDeclaredMadeForKids": False},
    }

    media = MediaFileUpload(video_path, mimetype="video/mp4", resumable=True)
    request = youtube.videos().insert(part="snippet,status", body=body, media_body=media)

    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            print(f"  YouTube upload progress: {int(status.progress() * 100)}%")

    video_id = response["id"]
    print(f"YouTube Short uploaded: https://youtube.com/shorts/{video_id}")
    return video_id
