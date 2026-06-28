"""Select 3 trending dishes for miniature ASMR food video content."""

import json
import anthropic


SYSTEM_PROMPT = """You are a social media content strategist specializing in ASMR and miniature cooking videos.
Your job is to pick dishes that will perform exceptionally well as short-form vertical video content on Instagram Reels and YouTube Shorts.
Focus on dishes with visually interesting textures, satisfying cooking processes, and wide cross-cultural appeal.
Avoid dishes that are too complex or have too many steps to show in an 8-second video."""

DISH_SELECTION_PROMPT = """Pick exactly 3 dishes for miniature ASMR cooking videos today.

Criteria:
- Currently trending on food social media (comfort food, fusion, street food)
- Visually satisfying to cook: sizzle, pour, fold, or spreading motions
- Clear distinct ingredients and garnishes visible at miniature scale
- 9:16 vertical video friendly (tall food, layered presentation, or dramatic pour)

Return ONLY a valid JSON array with exactly 3 objects, no markdown, no explanation:
[
  {
    "dish_name": "Name of the dish",
    "cuisine": "Origin cuisine",
    "key_ingredients": ["ingredient1", "ingredient2", "ingredient3"],
    "cooking_action": "Main visual cooking action (e.g., fold, pour, sizzle, spread)",
    "visual_highlight": "The most visually satisfying moment in preparing this dish",
    "garnish": "Microscopic garnish or sauce placed alongside when plated"
  }
]"""


def select_dishes(client: anthropic.Anthropic) -> list[dict]:
    """Ask Claude Haiku to pick 3 trending dishes for today's videos."""
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": DISH_SELECTION_PROMPT}],
    )
    raw = message.content[0].text.strip()
    dishes = json.loads(raw)
    if len(dishes) != 3:
        raise ValueError(f"Expected 3 dishes, got {len(dishes)}")
    return dishes
