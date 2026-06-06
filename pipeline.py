"""
Market Signals pipeline — research, analyse, generate PPTX + MD, post to Slack.

Environment variables:
  GROQ_API_KEY       required
  SLACK_BOT_TOKEN    required unless DRY_RUN=true
  SLACK_CHANNEL_ID   required unless DRY_RUN=true
  GITHUB_TOKEN       required unless DRY_RUN=true (for committing MD report)
  DRY_RUN            set to 'true' to skip Slack + git commit (save artifacts locally)
  DAYS_BACK          optional, default 45 — research window in days
"""
import json
import os
import subprocess
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import requests
from duckduckgo_search import DDGS
from groq import Groq
from pptx import Presentation

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GROQ_MODEL = "llama-3.3-70b-versatile"

FLEET_PRODUCTS = (
    "Fleet Charging Console, Charging Card, Home Charging Reimbursement, "
    "Charge & Fuel Card, LOGPAY Card 4U"
)
FLEET_COMPETITORS = (
    "DKV Mobility, UTA Edenred, Shell Recharge, BP Pulse, EnBW, Aral, Octopus Energy"
)
FLEET_USERS = "Fleet Managers, Company Car Drivers, HR / Finance / Sustainability"

SITE_PRODUCTS = (
    "Site Management Console, Standortanalyse, Planning & Installation, "
    "Hardware Migration, Maintenance & Operations"
)
SITE_COMPETITORS = (
    "Driivz, AMPECO, Monta, Virta, ChargePoint, E.On, EnBW, "
    "The Mobility House, Moon"
)
SITE_USERS = "Commercial Property Owners, Dealers, Facilities Managers"

FLEET_CATEGORIES = [
    ("Pricing model and subscription evolution",
     "fleet EV charging card fee kWh tariff pricing {m}"),
    ("Network access and roaming coverage",
     "eRoaming IONITY fleet charging network HPC partnership {m}"),
    ("Mixed-fleet and ICE-to-EV transition",
     "fuel card electrification ICE EV combined billing fleet {m}"),
    ("Home charging reimbursement and compliance",
     "home charging reimbursement kWh mandate wallbox fleet {m}"),
    ("Plug&Charge and authentication technology",
     "ISO 15118 Plug&Charge RFID fleet authentication {m}"),
    ("Mobile payment and in-app fuelling",
     "mobile app payment fuelling fleet digital receipt {m}"),
    ("Employee benefit and non-cash incentive regulation",
     "fleet benefit card tax-free EV charging social security exemption {m}"),
    ("Data, reporting and CO2 transparency",
     "CSRD fleet Scope 3 CO2 EV charging reporting dashboard {m}"),
    ("Fleet software consolidation and platform convergence",
     "CPMS fleet card integration unified platform merger {m}"),
    ("Commercial vehicle and HGV electrification",
     "HGV truck LCV fleet EV charging infrastructure {m}"),
    ("Competitor card and MSP moves",
     "DKV UTA Shell BP EnBW Aral Octopus fleet EV charging news {m}"),
    ("Regulatory and compliance obligations",
     "AFIR fleet ZEV mandate GDPR fleet card regulation {m}"),
    ("Customer pain points",
     "fleet manager EV charging problems unmet needs survey {m}"),
]

SITE_CATEGORIES = [
    ("CPMS platform competition and consolidation",
     "Driivz Ampeco Monta Virta ChargePoint CPMS news contract {m}"),
    ("Energy management and grid integration",
     "EV charging load management V2G demand response site {m}"),
    ("Hardware ecosystem and multi-vendor compatibility",
     "OCPP 2.0.1 Alpitronic Compleo charger integration {m}"),
    ("Site planning and installation market",
     "EV charging site planning grid connection permit cost {m}"),
    ("Public charging monetisation",
     "CPO ad-hoc payment AFIR eRoaming revenue CPO {m}"),
    ("Building and property regulation",
     "EPBD pre-cabling commercial building EV charging permit {m}"),
    ("CPO data and reporting obligations",
     "DATEX II CPO static dynamic data compliance {m}"),
    ("Maintenance, uptime and operations",
     "EV charging uptime remote diagnostics firmware SLA {m}"),
    ("AI and software-driven site operations",
     "AI predictive maintenance EV charging autonomous network {m}"),
    ("Segment-specific site demand",
     "retail depot parking municipality EV charging site {m}"),
    ("Fleet and site convergence",
     "fleet site unified EV charging console integration {m}"),
    ("Sustainability and carbon reporting at site level",
     "CO2 kWh renewable energy CSRD site charging sustainability {m}"),
]

# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

MONTH_NAMES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def compute_date_range():
    days_back = int(os.getenv("DAYS_BACK", "45"))
    end = date.today()
    start = end - timedelta(days=days_back)

    current_month = MONTH_NAMES[end.month]
    prior_month = MONTH_NAMES[start.month]
    year = end.year

    if prior_month == current_month:
        month_range = f"{current_month} {year}"
    else:
        month_range = f"{prior_month} {year} OR {current_month} {year}"

    label = end.strftime("%Y-%m")
    generated_at = end.strftime("%B %Y")
    period_start = start.strftime("%-d %B %Y")
    period_end = end.strftime("%-d %B %Y")

    return {
        "start": start,
        "end": end,
        "month_range": month_range,
        "label": label,
        "generated_at": generated_at,
        "period_start": period_start,
        "period_end": period_end,
    }


# ---------------------------------------------------------------------------
# Phase 1 — Research
# ---------------------------------------------------------------------------

def search_ddg(query, max_results=4):
    """Single DuckDuckGo query; returns list of {idx,title,snippet,url}."""
    results = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(query, max_results=max_results):
                results.append({
                    "title": r.get("title", ""),
                    "snippet": (r.get("body", "") or "")[:400],
                    "url": r.get("href", ""),
                })
    except Exception as e:
        print(f"  search warning: {e}", file=sys.stderr)
    return results


def run_searches(categories, month_range):
    """Run all searches; return {category: [{idx,title,snippet,url}]}."""
    results = {}
    for i, (category, query_template) in enumerate(categories, 1):
        query = query_template.replace("{m}", month_range)
        print(f"  [{i}/{len(categories)}] {category[:50]}...")
        hits = search_ddg(query)
        indexed = [{"idx": j + 1, **h} for j, h in enumerate(hits)]
        results[category] = indexed
        time.sleep(1.5)
    return results


# ---------------------------------------------------------------------------
# Phase 2 — Extract signals (Groq × 6 calls)
# ---------------------------------------------------------------------------

_groq_client = None


def groq_client():
    global _groq_client
    if _groq_client is None:
        _groq_client = Groq(api_key=os.environ["GROQ_API_KEY"])
    return _groq_client


def call_groq(system_prompt, user_prompt):
    """Call Groq in JSON mode; return parsed dict."""
    resp = groq_client().chat.completions.create(
        model=GROQ_MODEL,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.3,
        max_tokens=2048,
    )
    return json.loads(resp.choices[0].message.content)


def _format_snippets_for_prompt(category, hits):
    lines = [f"--- {category} ---"]
    for h in hits:
        lines.append(f"[{h['idx']}] {h['title']}")
        lines.append(f"URL: {h['url']}")
        lines.append(h["snippet"])
        lines.append("")
    return "\n".join(lines)


EXTRACT_SYSTEM = """You extract concrete market signals from EV charging industry search results.
Be factual and specific — only report what the search results actually contain.
Do not invent signals not supported by the text."""


def extract_signals(results_by_category, batch_size=4):
    """
    Extract concrete signals per category using batched Groq calls.
    Returns {category: [{signal: str, sources: [int]}]}
    """
    categories = list(results_by_category.keys())
    batches = [categories[i:i + batch_size] for i in range(0, len(categories), batch_size)]
    extracted = {}

    for b_idx, batch in enumerate(batches, 1):
        print(f"  extract batch {b_idx}/{len(batches)} ({len(batch)} categories)...")

        snippets_text = "\n".join(
            _format_snippets_for_prompt(cat, results_by_category[cat])
            for cat in batch
        )

        schema_example = {cat: [{"signal": "...", "sources": [1]}] for cat in batch}
        user_prompt = (
            f"Extract 2–3 concrete signals for each of these categories: "
            f"{', '.join(batch)}.\n\n"
            f"A signal is a specific product launch, pricing move, regulatory change, "
            f"partnership announcement, or verified customer pain point.\n\n"
            f"Return JSON matching this schema exactly:\n"
            f"{json.dumps(schema_example, indent=2)}\n\n"
            f"Search results:\n{snippets_text}"
        )

        try:
            data = call_groq(EXTRACT_SYSTEM, user_prompt)
            for cat in batch:
                extracted[cat] = data.get(cat, [])
        except Exception as e:
            print(f"  extract error batch {b_idx}: {e}", file=sys.stderr)
            for cat in batch:
                extracted[cat] = []

        time.sleep(1.0)

    return extracted


def build_url_map(results_by_category):
    """Build {(category, idx): url} lookup."""
    url_map = {}
    for cat, hits in results_by_category.items():
        for h in hits:
            url_map[(cat, h["idx"])] = h["url"]
    return url_map


def resolve_urls(signals_with_sources, category, url_map):
    """Replace source index numbers with actual URLs."""
    resolved = []
    for item in signals_with_sources:
        urls = [
            url_map.get((category, src), "")
            for src in item.get("sources", [])
            if url_map.get((category, src), "")
        ]
        resolved.append({"text": item["signal"], "urls": urls})
    return resolved


# ---------------------------------------------------------------------------
# Phase 3 — Synthesise report (Groq × 2 calls)
# ---------------------------------------------------------------------------

SYNTH_SYSTEM_TEMPLATE = """You are a strategic analyst for Elli, a Volkswagen Group EV charging company.

{area} Products: {products}
Key Competitors: {competitors}
Primary Users: {users}

Analyse the provided market signals (collected over the past 30–45 days) and produce
the monthly Market Signals report section for {area}.

Return ONLY valid JSON with this exact schema:
{{
  "trends": [
    {{"title": "...", "description": "2–3 sentences max 180 chars", "sources": [1,2]}},
    ...
  ],
  "competitor_moves": [{{"text": "...", "sources": [1]}}],
  "unmet_needs": [{{"text": "...", "sources": [1]}}],
  "active_regulations": [{{"text": "...", "sources": [1]}}],
  "strategic_implications": [{{"text": "..."}}],
  "open_questions": [{{"text": "..."}}]
}}

Produce exactly 5 trends, 3–5 items in each other list."""


def _signals_to_text(extracted, categories, url_map):
    """Format extracted signals as a numbered reference block for the synthesis prompt."""
    lines = []
    ref_num = 1
    ref_map = {}

    for cat in categories:
        for item in extracted.get(cat, []):
            for src_idx in item.get("sources", []):
                url = url_map.get((cat, src_idx), "")
                if url and (cat, src_idx) not in ref_map:
                    ref_map[(cat, src_idx)] = ref_num
                    ref_num += 1

    for cat in categories:
        items = extracted.get(cat, [])
        if not items:
            continue
        lines.append(f"### {cat}")
        for item in items:
            refs = [ref_map[(cat, s)] for s in item.get("sources", [])
                    if (cat, s) in ref_map]
            ref_str = " ".join(f"[{r}]" for r in refs)
            lines.append(f"- {item['signal']} {ref_str}")
        lines.append("")

    ref_list = ["**Sources:**"]
    for (cat, idx), num in sorted(ref_map.items(), key=lambda x: x[1]):
        url = url_map.get((cat, idx), "")
        ref_list.append(f"[{num}] {url}")

    return "\n".join(lines), ref_map


def synthesise_report(extracted, url_map, area, products, competitors, users, categories):
    system = SYNTH_SYSTEM_TEMPLATE.format(
        area=area, products=products, competitors=competitors, users=users
    )
    signals_text, ref_map = _signals_to_text(extracted, categories, url_map)
    user_prompt = f"Market signals:\n\n{signals_text}"

    try:
        data = call_groq(system, user_prompt)
    except Exception as e:
        print(f"  synthesis error for {area}: {e}", file=sys.stderr)
        data = {
            "trends": [],
            "competitor_moves": [],
            "unmet_needs": [],
            "active_regulations": [],
            "strategic_implications": [],
            "open_questions": [],
        }

    # Resolve source indices to URLs using ref_map
    inv_ref_map = {v: k for k, v in ref_map.items()}

    def resolve_item(item):
        urls = [
            url_map.get(inv_ref_map[s], "")
            for s in item.get("sources", [])
            if s in inv_ref_map and url_map.get(inv_ref_map[s], "")
        ]
        return {"text": item.get("text", item.get("title", item.get("description", ""))), "urls": urls}

    def resolve_trend(item):
        urls = [
            url_map.get(inv_ref_map[s], "")
            for s in item.get("sources", [])
            if s in inv_ref_map and url_map.get(inv_ref_map[s], "")
        ]
        return {
            "title": item.get("title", ""),
            "description": item.get("description", ""),
            "urls": urls,
        }

    return {
        "trends": [resolve_trend(t) for t in data.get("trends", [])[:5]],
        "competitor_moves": [resolve_item(x) for x in data.get("competitor_moves", [])],
        "unmet_needs": [resolve_item(x) for x in data.get("unmet_needs", [])],
        "active_regulations": [resolve_item(x) for x in data.get("active_regulations", [])],
        "strategic_implications": [resolve_item(x) for x in data.get("strategic_implications", [])],
        "open_questions": [resolve_item(x) for x in data.get("open_questions", [])],
    }


# ---------------------------------------------------------------------------
# Phase 4a — Build PPTX replacements
# ---------------------------------------------------------------------------

def _bullet_list_plain(items):
    return "\n".join(f"• {item['text']}" for item in items) if items else "No signals found."


def _trends_plain(trends):
    lines = []
    for i, t in enumerate(trends, 1):
        lines.append(f"{i}. {t['title']}")
        if t["description"]:
            lines.append(f"   {t['description']}")
        lines.append("")
    return "\n".join(lines).strip() if lines else "No trends found."


def build_replacements(fleet_report, site_report, dates):
    return {
        "{{GENERATED_AT}}": dates["generated_at"],
        "{{PERIOD_START}}": dates["period_start"],
        "{{PERIOD_END}}": dates["period_end"],
        "{{FLEET_TRENDS}}": _trends_plain(fleet_report["trends"]),
        "{{FLEET_COMPETITOR_MOVES}}": _bullet_list_plain(fleet_report["competitor_moves"]),
        "{{FLEET_UNMET_NEEDS}}": _bullet_list_plain(fleet_report["unmet_needs"]),
        "{{FLEET_ACTIVE_REGULATIONS}}": _bullet_list_plain(fleet_report["active_regulations"]),
        "{{FLEET_STRATEGIC_IMPLICATIONS}}": _bullet_list_plain(fleet_report["strategic_implications"]),
        "{{SITE_TRENDS}}": _trends_plain(site_report["trends"]),
        "{{SITE_COMPETITOR_MOVES}}": _bullet_list_plain(site_report["competitor_moves"]),
        "{{SITE_UNMET_NEEDS}}": _bullet_list_plain(site_report["unmet_needs"]),
        "{{SITE_ACTIVE_REGULATIONS}}": _bullet_list_plain(site_report["active_regulations"]),
        "{{SITE_STRATEGIC_IMPLICATIONS}}": _bullet_list_plain(site_report["strategic_implications"]),
        "{{OPEN_QUESTIONS}}": _bullet_list_plain(
            fleet_report["open_questions"] + site_report["open_questions"]
        ),
    }


# ---------------------------------------------------------------------------
# Phase 4b — Fill PPTX
# ---------------------------------------------------------------------------

def consolidate_runs(para):
    """Merge all runs in a paragraph into the first run to fix split placeholders."""
    runs = para.runs
    if len(runs) <= 1:
        return
    full_text = "".join(r.text for r in runs)
    runs[0].text = full_text
    for run in runs[1:]:
        run.text = ""


def fill_pptx(template_path, output_path, replacements):
    prs = Presentation(template_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                consolidate_runs(para)
                for run in para.runs:
                    for key, value in replacements.items():
                        if key in run.text:
                            run.text = run.text.replace(key, value)
    prs.save(output_path)


# ---------------------------------------------------------------------------
# Phase 4c — Generate Markdown
# ---------------------------------------------------------------------------

def _cite(urls):
    if not urls:
        return ""
    return " " + " ".join(f"[[{i+1}]]({u})" for i, u in enumerate(urls) if u)


def _bullet_list_md(items):
    if not items:
        return "- No signals found.\n"
    return "".join(f"- {item['text']}{_cite(item['urls'])}\n" for item in items)


def _trends_md(trends):
    lines = []
    for i, t in enumerate(trends, 1):
        lines.append(f"{i}. **{t['title']}**{_cite(t['urls'])}")
        if t["description"]:
            lines.append(f"   {t['description']}")
        lines.append("")
    return "\n".join(lines) if lines else "- No trends found.\n"


def _sources_section(report, area_label):
    seen = set()
    sources = []
    for section in ["trends", "competitor_moves", "unmet_needs",
                     "active_regulations", "strategic_implications"]:
        for item in report.get(section, []):
            for url in item.get("urls", []):
                if url and url not in seen:
                    seen.add(url)
                    sources.append(url)
    if not sources:
        return ""
    lines = [f"### {area_label}"]
    for i, url in enumerate(sources, 1):
        lines.append(f"{i}. {url}")
    return "\n".join(lines) + "\n"


def generate_markdown(fleet_report, site_report, dates):
    label = dates["generated_at"]
    period = f"{dates['period_start']} – {dates['period_end']}"
    today = dates["end"].strftime("%Y-%m-%d")

    sections = [
        f"# Elli Market Signals — {label}",
        f"*Research period: {period} | Generated: {today}*",
        "",
        "---",
        "",
        "## Fleet Mobility Management",
        "",
        "### Top 5 Trends",
        _trends_md(fleet_report["trends"]),
        "### Competitor Moves",
        _bullet_list_md(fleet_report["competitor_moves"]),
        "",
        "### Unmet Customer Needs",
        _bullet_list_md(fleet_report["unmet_needs"]),
        "",
        "### Active Regulations",
        _bullet_list_md(fleet_report["active_regulations"]),
        "",
        "### Strategic Implications for Elli",
        _bullet_list_md(fleet_report["strategic_implications"]),
        "",
        "---",
        "",
        "## Charging Site Management",
        "",
        "### Top 5 Trends",
        _trends_md(site_report["trends"]),
        "### Competitor Moves",
        _bullet_list_md(site_report["competitor_moves"]),
        "",
        "### Unmet Customer Needs",
        _bullet_list_md(site_report["unmet_needs"]),
        "",
        "### Active Regulations",
        _bullet_list_md(site_report["active_regulations"]),
        "",
        "### Strategic Implications for Elli",
        _bullet_list_md(site_report["strategic_implications"]),
        "",
        "---",
        "",
        "## Open Questions",
        _bullet_list_md(fleet_report["open_questions"] + site_report["open_questions"]),
        "",
        "---",
        "",
        "## Sources",
        "",
        _sources_section(fleet_report, "Fleet Mobility Management"),
        "",
        _sources_section(site_report, "Charging Site Management"),
    ]
    return "\n".join(sections)


# ---------------------------------------------------------------------------
# Phase 5 — Commit MD + post to Slack
# ---------------------------------------------------------------------------

def commit_markdown(md_path, label):
    """Commit the MD report to the repo and push; returns the GitHub blob URL."""
    repo_root = Path(__file__).parent
    token = os.getenv("GITHUB_TOKEN", "")

    try:
        subprocess.run(["git", "config", "user.email", "pipeline@elli-market-signals"], check=True, cwd=repo_root)
        subprocess.run(["git", "config", "user.name", "Market Signals Pipeline"], check=True, cwd=repo_root)
        subprocess.run(["git", "add", str(md_path)], check=True, cwd=repo_root)
        subprocess.run(
            ["git", "commit", "-m", f"report: add market signals {label}"],
            check=True, cwd=repo_root
        )
        # Inject token into remote URL for push
        result = subprocess.run(["git", "remote", "get-url", "origin"],
                                 capture_output=True, text=True, cwd=repo_root)
        remote_url = result.stdout.strip()
        if token and "github.com" in remote_url and "@" not in remote_url:
            auth_url = remote_url.replace("https://", f"https://x-access-token:{token}@")
            subprocess.run(["git", "push", auth_url, "HEAD"], check=True, cwd=repo_root)
        else:
            subprocess.run(["git", "push", "-u", "origin", "HEAD"], check=True, cwd=repo_root)

        # Derive GitHub URL from remote
        clean_url = remote_url.replace(".git", "")
        branch = subprocess.run(["git", "branch", "--show-current"],
                                  capture_output=True, text=True, cwd=repo_root).stdout.strip()
        return f"{clean_url}/blob/{branch}/{md_path.name}"
    except subprocess.CalledProcessError as e:
        print(f"  git error: {e}", file=sys.stderr)
        return ""


def post_to_slack(pptx_path, md_path, github_url):
    token = os.environ["SLACK_BOT_TOKEN"]
    channel = os.environ["SLACK_CHANNEL_ID"]
    headers = {"Authorization": f"Bearer {token}"}

    message = (
        f":bar_chart: *Elli Market Signals report is ready!*\n"
        f"Research period: see attached files.\n"
    )
    if github_url:
        message += f":link: <{github_url}|View full report with citations on GitHub>"

    # Post text message
    resp = requests.post(
        "https://slack.com/api/chat.postMessage",
        headers=headers,
        json={"channel": channel, "text": message},
    )
    resp.raise_for_status()
    if not resp.json().get("ok"):
        print(f"  Slack message error: {resp.json()}", file=sys.stderr)

    # Upload PPTX
    for fpath, fname in [(pptx_path, "Market Signals Report.pptx"),
                          (md_path, f"{md_path.name}")]:
        if not fpath.exists():
            continue
        with open(fpath, "rb") as f:
            resp = requests.post(
                "https://slack.com/api/files.upload",
                headers=headers,
                data={"channels": channel, "filename": fname},
                files={"file": (fname, f)},
            )
        resp.raise_for_status()
        if not resp.json().get("ok"):
            print(f"  Slack upload error for {fname}: {resp.json()}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    dry_run = os.getenv("DRY_RUN", "false").lower() == "true"
    dates = compute_date_range()
    label = dates["label"]

    print(f"=== Market Signals Pipeline — {dates['generated_at']} ===")
    print(f"Research window: {dates['period_start']} to {dates['period_end']}")
    print(f"DRY_RUN={dry_run}")

    # --- Phase 1: Research ---
    print("\n[Phase 1] Running searches...")
    print("  Fleet categories:")
    fleet_raw = run_searches(FLEET_CATEGORIES, dates["month_range"])
    print("  Site categories:")
    site_raw = run_searches(SITE_CATEGORIES, dates["month_range"])

    fleet_url_map = build_url_map(fleet_raw)
    site_url_map = build_url_map(site_raw)

    # --- Phase 2: Extract ---
    print("\n[Phase 2] Extracting signals...")
    print("  Fleet:")
    fleet_extracted = extract_signals(fleet_raw, batch_size=4)
    print("  Site:")
    site_extracted = extract_signals(site_raw, batch_size=4)

    # --- Phase 3: Synthesise ---
    print("\n[Phase 3] Synthesising report...")
    print("  Fleet report...")
    fleet_report = synthesise_report(
        fleet_extracted, fleet_url_map,
        area="Fleet Mobility Management",
        products=FLEET_PRODUCTS,
        competitors=FLEET_COMPETITORS,
        users=FLEET_USERS,
        categories=[c for c, _ in FLEET_CATEGORIES],
    )
    print("  Site report...")
    site_report = synthesise_report(
        site_extracted, site_url_map,
        area="Charging Site Management",
        products=SITE_PRODUCTS,
        competitors=SITE_COMPETITORS,
        users=SITE_USERS,
        categories=[c for c, _ in SITE_CATEGORIES],
    )

    # --- Phase 4: Generate outputs ---
    print("\n[Phase 4] Generating outputs...")
    pptx_path = Path(f"output_{label}.pptx")
    md_path = Path("reports") / f"output_{label}.md"
    md_path.parent.mkdir(exist_ok=True)

    replacements = build_replacements(fleet_report, site_report, dates)
    fill_pptx("template.pptx", str(pptx_path), replacements)
    print(f"  ✓ PPTX: {pptx_path}")

    md_content = generate_markdown(fleet_report, site_report, dates)
    md_path.write_text(md_content, encoding="utf-8")
    print(f"  ✓ MD: {md_path}")

    if dry_run:
        print("\n[DRY RUN] Skipping git commit and Slack post.")
        print(f"  Artifacts: {pptx_path}, {md_path}")
        return

    # --- Phase 5: Commit + notify ---
    print("\n[Phase 5] Committing MD and posting to Slack...")
    github_url = commit_markdown(md_path, label)
    print(f"  GitHub URL: {github_url or '(push failed)'}")

    post_to_slack(pptx_path, md_path, github_url)
    print("  ✓ Posted to Slack")

    print(f"\n=== Done — {dates['generated_at']} ===")


if __name__ == "__main__":
    main()
