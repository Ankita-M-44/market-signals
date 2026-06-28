"""Generate dish-specific Veo video prompts from the master template."""

import re
import anthropic


EXPAND_PROMPT = """You are given a master miniature ASMR cooking video prompt template and details about a specific dish.
Your job is to produce a final video generation prompt by:
1. Replacing [FOOD_ITEM] with the dish name throughout
2. Expanding the ACTION SEQUENCE steps with dish-specific details (ingredients, garnishes, textures)
3. Keeping the exact scene setup, lighting, and aesthetics section unchanged
4. Keeping the total prompt under 500 words

Return ONLY the final prompt text with no explanation or markdown."""


def generate_video_prompt(
    client: anthropic.Anthropic,
    master_prompt: str,
    dish: dict,
) -> str:
    """Expand the master prompt template for a specific dish."""
    dish_context = f"""Dish: {dish['dish_name']} ({dish['cuisine']})
Key ingredients: {', '.join(dish['key_ingredients'])}
Main cooking action: {dish['cooking_action']}
Visual highlight: {dish['visual_highlight']}
Garnish/plating: {dish['garnish']}"""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[
            {
                "role": "user",
                "content": f"Master prompt template:\n\n{master_prompt}\n\nDish details:\n{dish_context}\n\n{EXPAND_PROMPT}",
            }
        ],
    )
    return message.content[0].text.strip()


def extract_master_prompt(notion_page_text: str) -> str:
    """Extract the master prompt block from the Notion page content."""
    # The prompt starts after the page title and before the database embed
    lines = notion_page_text.split("\n")
    prompt_lines = []
    capturing = False
    for line in lines:
        stripped = line.strip()
        # Start capturing after the first meaningful non-title line
        if not capturing and stripped.startswith("Macro cinematic"):
            capturing = True
        if capturing:
            # Stop at the database embed line
            if stripped.startswith("<database") or stripped.startswith("\\[database"):
                break
            # Clean escaped brackets from Notion markdown
            clean = re.sub(r"\\\\?(\[|\])", lambda m: m.group(1), stripped)
            prompt_lines.append(clean)
    return "\n".join(prompt_lines).strip()
