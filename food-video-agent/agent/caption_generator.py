"""Generate Instagram/YouTube captions for each video."""

import anthropic


CAPTION_PROMPT = """Write a social media caption for a miniature ASMR cooking video of {dish_name}.

Requirements:
- Opens with an engaging hook (1 sentence, no emoji at start)
- Tone: satisfying, calming, ASMR-friendly
- 2-3 sentences total
- Ends with 5-8 relevant hashtags on a new line
- Include #miniaturecooking #asmrcooking #shorts #reels
- Keep it under 200 characters before the hashtags

Return ONLY the caption text, no explanation."""


def generate_caption(client: anthropic.Anthropic, dish: dict) -> str:
    """Generate a social media caption for the given dish video."""
    prompt = CAPTION_PROMPT.format(dish_name=dish["dish_name"])
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text.strip()
