"""Generate miniature food videos using Google Veo 2 via the Gemini API."""

import time
import os
import requests

GEMINI_API_BASE = "https://generativelanguage.googleapis.com"
VEO_MODEL = "veo-2.0-generate-001"

# Poll interval and max wait time for video generation
POLL_INTERVAL_SEC = 15
MAX_WAIT_SEC = 600  # 10 minutes


def generate_video(
    api_key: str,
    prompt: str,
    output_path: str,
    aspect_ratio: str = "9:16",
    duration_seconds: int = 8,
    dry_run: bool = False,
) -> str:
    """
    Generate a video using Veo 2 and save it to output_path.
    Returns the output_path on success.
    """
    if dry_run:
        print(f"[DRY RUN] Would generate video for prompt: {prompt[:100]}...")
        # Create a placeholder file so the rest of the pipeline can proceed
        with open(output_path, "wb") as f:
            f.write(b"DRY_RUN_PLACEHOLDER")
        return output_path

    # Step 1: Submit the long-running generation request
    url = f"{GEMINI_API_BASE}/v1beta/models/{VEO_MODEL}:predictLongRunning"
    payload = {
        "instances": [{"prompt": prompt}],
        "parameters": {
            "aspectRatio": aspect_ratio,
            "durationSeconds": duration_seconds,
            "sampleCount": 1,
        },
    }
    resp = requests.post(
        url,
        params={"key": api_key},
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    operation_name = resp.json()["name"]
    print(f"Video generation started: {operation_name}")

    # Step 2: Poll until complete
    poll_url = f"{GEMINI_API_BASE}/v1beta/{operation_name}"
    elapsed = 0
    while elapsed < MAX_WAIT_SEC:
        time.sleep(POLL_INTERVAL_SEC)
        elapsed += POLL_INTERVAL_SEC
        poll_resp = requests.get(poll_url, params={"key": api_key}, timeout=30)
        poll_resp.raise_for_status()
        data = poll_resp.json()
        if data.get("done"):
            break
        print(f"  Waiting for video... ({elapsed}s elapsed)")
    else:
        raise TimeoutError(f"Video generation timed out after {MAX_WAIT_SEC}s")

    # Step 3: Extract the video URI and download it
    response_obj = data.get("response", {})
    predictions = response_obj.get("predictions", [])
    if not predictions:
        raise ValueError(f"No predictions returned: {data}")

    video_uri = predictions[0].get("bytesBase64Encoded") or predictions[0].get("video", {}).get("uri")

    if predictions[0].get("bytesBase64Encoded"):
        import base64
        video_bytes = base64.b64decode(predictions[0]["bytesBase64Encoded"])
        with open(output_path, "wb") as f:
            f.write(video_bytes)
    elif video_uri:
        dl_resp = requests.get(video_uri, timeout=120)
        dl_resp.raise_for_status()
        with open(output_path, "wb") as f:
            f.write(dl_resp.content)
    else:
        raise ValueError(f"Cannot find video data in response: predictions[0]")

    print(f"Video saved to {output_path}")
    return output_path
