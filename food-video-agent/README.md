# Mini Food Video Automation Agent

Automated pipeline for generating and posting miniature ASMR food cooking videos to Instagram Reels and YouTube Shorts.

## How It Works

```
Daily 8am UTC
    ↓
[pipeline.py]
  1. Claude Haiku picks 3 trending dishes
  2. Fetches master prompt from Notion
  3. Generates dish-specific Veo 2 video prompts
  4. Generates social media captions
  5. Creates entries in Notion Videos DB (Status: Not started)
  6. Calls Google Veo 2 API → 9:16 video, 8 seconds
  7. Uploads video to Notion Files & media
  8. Updates Status → Created

9am UTC   → [uploader.py] → Video 1 → Instagram Reel + YouTube Short → Status: Uploaded
1pm UTC   → [uploader.py] → Video 2 → Instagram Reel + YouTube Short → Status: Uploaded
6pm UTC   → [uploader.py] → Video 3 → Instagram Reel + YouTube Short → Status: Uploaded
```

## Notion Database

- **Page**: Mini food (https://app.notion.com/p/38d48c078aa980d99c39ff0a6e5e8f50)
- **Database**: Videos (ID: `38d48c07-8aa9-8029-8bd3-000b11e75d47`)
- **Schema**: Name | Caption | Date | Files & media | Status

The master video prompt template lives on the Notion "Mini food" page. Edit it there to change the visual style for all future videos.

## Required GitHub Secrets

Set these in your repo Settings → Secrets and variables → Actions:

| Secret | Where to get it |
|---|---|
| `ANTHROPIC_API_KEY` | https://console.anthropic.com |
| `GEMINI_API_KEY` | https://aistudio.google.com/app/apikey |
| `NOTION_API_KEY` | https://www.notion.so/my-integrations → New integration |
| `NOTION_DATABASE_ID` | `38d48c07-8aa9-8029-8bd3-000b11e75d47` (hardcoded fallback) |
| `INSTAGRAM_ACCESS_TOKEN` | See Instagram setup below |
| `INSTAGRAM_ACCOUNT_ID` | See Instagram setup below |
| `YOUTUBE_CLIENT_ID` | See YouTube setup below |
| `YOUTUBE_CLIENT_SECRET` | See YouTube setup below |
| `YOUTUBE_REFRESH_TOKEN` | See YouTube setup below |

## One-Time Setup

### Notion Integration
1. Go to https://www.notion.so/my-integrations → New integration
2. Give it access to your "Mini food" page (Share → Invite integration)
3. Copy the Internal Integration Token → `NOTION_API_KEY`

### Instagram (Graph API)
1. Go to https://developers.facebook.com → Create App → Business
2. Add "Instagram Graph API" product
3. Connect your Instagram Business account to a Facebook Page
4. Generate a long-lived access token (valid 60 days, must be refreshed):
   ```
   GET https://graph.facebook.com/v22.0/oauth/access_token
     ?grant_type=fb_exchange_token
     &client_id={APP_ID}
     &client_secret={APP_SECRET}
     &fb_exchange_token={SHORT_LIVED_TOKEN}
   ```
5. Get your Instagram Business Account ID:
   ```
   GET https://graph.facebook.com/v22.0/me/accounts?access_token={TOKEN}
   # Then: GET /{page-id}?fields=instagram_business_account&access_token={TOKEN}
   ```

### YouTube (Data API v3)
1. Go to https://console.cloud.google.com → New project
2. Enable "YouTube Data API v3"
3. Create OAuth 2.0 credentials (Desktop app type)
4. Run the one-time auth flow to get a refresh token:
   ```bash
   pip install google-auth-oauthlib
   python -c "
   from google_auth_oauthlib.flow import InstalledAppFlow
   flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', ['https://www.googleapis.com/auth/youtube.upload'])
   creds = flow.run_local_server(port=0)
   print('Refresh token:', creds.refresh_token)
   "
   ```
5. Store the refresh token as `YOUTUBE_REFRESH_TOKEN`

## Local Development

```bash
cd food-video-agent
pip install -r requirements.txt

# Test the pipeline (no API calls, no Notion writes)
ANTHROPIC_API_KEY=xxx GEMINI_API_KEY=xxx NOTION_API_KEY=xxx python pipeline.py --dry-run

# Test the uploader (no uploads)
NOTION_API_KEY=xxx INSTAGRAM_ACCESS_TOKEN=xxx INSTAGRAM_ACCOUNT_ID=xxx \
YOUTUBE_CLIENT_ID=xxx YOUTUBE_CLIENT_SECRET=xxx YOUTUBE_REFRESH_TOKEN=xxx \
python uploader.py --dry-run
```

## Manual Trigger

Both workflows support `workflow_dispatch` with a `dry_run` toggle in GitHub Actions → Run workflow.
