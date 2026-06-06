# Market Signals — Elli Enterprise PM

Automated monthly market research pipeline for Elli's Enterprise PM team.

**What it does:** Searches EV charging industry news across 25 signal categories, extracts concrete signals using Groq LLM, synthesises a structured report, and delivers a 14-slide Elli-branded PPTX + a cited Markdown report to Slack.

**Cost per run:** ~$0 (Groq free tier + DuckDuckGo keyless search)

---

## How It Works

```
GitHub Actions (1st of month)
  → DuckDuckGo: 25 queries (Fleet × 13, Site × 12)
  → Groq llama-3.3-70b: extract signals (6 calls) + synthesise report (2 calls)
  → output_YYYY-MM.pptx  (14-slide Elli-branded deck)
  → reports/output_YYYY-MM.md  (full report with source citations)
  → Slack: PPTX + MD posted to channel
```

---

## Setup

### 1. Add GitHub Secrets

In **Settings → Secrets and variables → Actions → Repository secrets**, add:

| Secret | How to get it |
|---|---|
| `GROQ_API_KEY` | [console.groq.com](https://console.groq.com) → API Keys |
| `SERPER_API_KEY` | [serper.dev](https://serper.dev) → Sign up free (email only, no credit card) → Dashboard → API Key |
| `SLACK_BOT_TOKEN` | [api.slack.com/apps](https://api.slack.com/apps) → create app → OAuth & Permissions → add `files:write` + `chat:write` scopes → Install to workspace → copy the `xoxb-...` token |
| `SLACK_CHANNEL_ID` | In Slack: right-click your target channel → Copy link → the last path segment is the channel ID (e.g. `C1234567890`) |

`GITHUB_TOKEN` is provided automatically by GitHub Actions — no setup needed.

### 2. Run Manually

Trigger at any time: **Actions → Monthly Market Signals Report → Run workflow**

- `dry_run: true` (default) — generates PPTX + MD as downloadable artifacts; skips Slack post and git commit. Use this for testing.
- `dry_run: false` — full run: posts to Slack, commits MD to `reports/`.
- `days_back` — research window in days (default: 45).

### 3. Automatic Monthly Run

The workflow runs automatically at 07:00 UTC on the 1st of every month.

---

## Outputs

| File | Description |
|---|---|
| `output_YYYY-MM.pptx` | 14-slide Elli-branded presentation |
| `reports/output_YYYY-MM.md` | Full report with inline source citations — feeds the product intelligence report |

The MD file is committed to this repo and auto-rendered by GitHub at:
`https://github.com/Ankita-M-44/market-signals/blob/main/reports/output_YYYY-MM.md`

---

## Local Development

```bash
pip install -r requirements.txt

# Generate template (one-time, already committed)
python create_template.py

# Test pipeline (no Slack, no git push)
DRY_RUN=true GROQ_API_KEY=your-key python pipeline.py
```

---

## Signal Coverage

**Fleet Mobility Management** (13 categories)
Pricing · Network/Roaming · Mixed-fleet transition · Home charging compliance ·
Plug&Charge · Mobile payment · Employee benefit regulation · CO₂ reporting ·
Platform convergence · HGV electrification · Competitor moves (DKV, UTA, Shell, BP, EnBW, Aral, Octopus) ·
Regulatory obligations · Customer pain points

**Charging Site Management** (12 categories)
CPMS competition · Energy management/V2G · Hardware compatibility · Site planning ·
Public charging monetisation · Building regulation · CPO data obligations ·
Maintenance/uptime · AI site operations · Segment demand ·
Fleet-site convergence · Carbon reporting
