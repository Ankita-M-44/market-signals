#!/usr/bin/env python3
"""
Mini Food Video Pipeline — Steps 1-4
Selects dishes → generates prompts → creates Notion entries → generates videos

Usage:
  python pipeline.py             # full run
  python pipeline.py --dry-run   # skip video gen + Notion writes, print output only
"""

import os
import sys
import argparse
import tempfile
from pathlib import Path

import anthropic

from agent.dish_selector import select_dishes
from agent.prompt_generator import generate_video_prompt, extract_master_prompt
from agent.caption_generator import generate_caption
from agent.notion_tracker import (
    create_video_entry,
    update_status,
    upload_video_file,
    fetch_master_prompt_page,
)
from agent.video_generator import generate_video

NOTION_MINI_FOOD_PAGE_ID = "38d48c078aa980d99c39ff0a6e5e8f50"


def load_env() -> dict:
    return {
        "anthropic_key": os.environ["ANTHROPIC_API_KEY"],
        "gemini_key": os.environ["GEMINI_API_KEY"],
        "notion_key": os.environ["NOTION_API_KEY"],
        "notion_db_id": os.environ.get(
            "NOTION_DATABASE_ID", "38d48c07-8aa9-8029-8bd3-000b11e75d47"
        ),
    }


def main(dry_run: bool = False) -> None:
    print("=== Mini Food Video Pipeline ===")
    print(f"Mode: {'DRY RUN' if dry_run else 'LIVE'}\n")

    env = load_env()
    claude = anthropic.Anthropic(api_key=env["anthropic_key"])

    # --- Step 1: Select 3 trending dishes ---
    print("Step 1: Selecting trending dishes...")
    dishes = select_dishes(claude)
    for i, d in enumerate(dishes):
        print(f"  [{i+1}] {d['dish_name']} ({d['cuisine']})")

    # --- Step 2: Fetch master prompt from Notion ---
    print("\nStep 2: Fetching master prompt from Notion...")
    if dry_run:
        master_raw = "[DRY RUN] Macro cinematic, 8k resolution, ultra-realistic video of a miniature ASMR cooking sequence.\n[FOOD_ITEM] sizzles on a tiny pan."
    else:
        master_raw = fetch_master_prompt_page(env["notion_key"], NOTION_MINI_FOOD_PAGE_ID)
    master_prompt = extract_master_prompt(master_raw)
    print(f"  Master prompt extracted ({len(master_prompt)} chars)")

    # --- Steps 3-4: Per dish: generate prompt, caption, Notion entry, video ---
    output_dir = Path("videos")
    output_dir.mkdir(exist_ok=True)

    for slot_index, dish in enumerate(dishes):
        dish_name = dish["dish_name"]
        print(f"\n--- Processing [{slot_index+1}/3]: {dish_name} ---")

        # Generate video prompt
        print("  Generating video prompt...")
        video_prompt = generate_video_prompt(claude, master_prompt, dish)
        print(f"  Prompt: {video_prompt[:120]}...")

        # Generate social caption
        print("  Generating caption...")
        caption = generate_caption(claude, dish)
        print(f"  Caption: {caption[:80]}...")

        # Create Notion entry
        print("  Creating Notion entry...")
        if dry_run:
            page_id = f"DRY_RUN_PAGE_{slot_index}"
            print(f"  [DRY RUN] Notion page ID: {page_id}")
        else:
            page_id = create_video_entry(
                env["notion_key"],
                env["notion_db_id"],
                dish,
                caption,
                video_prompt,
                slot_index,
            )
            print(f"  Notion page created: {page_id}")

        # Generate video
        video_filename = f"video_{slot_index}_{dish_name.replace(' ', '_').lower()}.mp4"
        video_path = str(output_dir / video_filename)
        print(f"  Generating video → {video_path}")
        generate_video(
            api_key=env["gemini_key"],
            prompt=video_prompt,
            output_path=video_path,
            aspect_ratio="9:16",
            duration_seconds=8,
            dry_run=dry_run,
        )

        # Upload video to Notion and update status
        if not dry_run:
            print("  Uploading video to Notion...")
            upload_video_file(env["notion_key"], page_id, video_path, video_filename)
            update_status(env["notion_key"], page_id, "Created")
            print(f"  Status updated to 'Created'")
        else:
            print(f"  [DRY RUN] Would upload {video_filename} to Notion page {page_id}")

    print("\n=== Pipeline complete ===")
    print(f"3 videos {'prepared (dry run)' if dry_run else 'generated and stored in Notion'}")
    print("Next step: upload_videos.yml will post them at 9am / 1pm / 6pm UTC")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Mini food video generation pipeline")
    parser.add_argument("--dry-run", action="store_true", help="Skip API calls and Notion writes")
    args = parser.parse_args()
    main(dry_run=args.dry_run)
