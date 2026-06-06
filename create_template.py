"""
One-time script: generates template.pptx from the Elli corporate template.
Run: python create_template.py
Then commit the resulting template.pptx to the repo.
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

SOURCE = "Elli_Mobility_PP_Template_2.0.pptx"
OUTPUT = "template.pptx"

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x00, 0xD4, 0x8A)


def remove_all_slides(prs):
    """Remove all existing slides from the presentation, keeping layouts/masters."""
    slide_id_list = prs.slides._sldIdLst
    for sld_id in list(slide_id_list):
        rId = sld_id.get(qn("r:id"))
        prs.part.drop_rel(rId)
        slide_id_list.remove(sld_id)


def get_layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"Layout '{name}' not found")


def set_placeholder_text(slide, idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            return
    raise ValueError(f"Placeholder idx={idx} not found on slide")


def add_text_box(slide, text, left_in, top_in, width_in, height_in,
                 font_size_pt, bold=True, color=WHITE):
    txBox = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_content_slide(prs, layout, title, body, subtitle=None):
    slide = prs.slides.add_slide(layout)
    set_placeholder_text(slide, 0, title)
    set_placeholder_text(slide, 15, body)
    if subtitle is not None:
        try:
            set_placeholder_text(slide, 4, subtitle)
        except ValueError:
            pass
    return slide


def add_section_divider(prs, layout, section_title, section_number):
    """Section break slide with large title TextBox, matching slide 3 in the source."""
    slide = prs.slides.add_slide(layout)
    # Large section title (matches slide 3: 72pt bold white)
    add_text_box(slide, section_title,
                 left_in=1.92, top_in=2.29, width_in=9.52, height_in=2.41,
                 font_size_pt=72, bold=True, color=WHITE)
    # Section number (matches slide 3: huge bold white)
    add_text_box(slide, section_number,
                 left_in=0.57, top_in=2.06, width_in=1.40, height_in=2.86,
                 font_size_pt=120, bold=True, color=WHITE)
    return slide


def main():
    prs = Presentation(SOURCE)

    remove_all_slides(prs)

    dark_content = get_layout(prs, "Dark background content")
    dark_empty = get_layout(prs, "Dark background empty")
    chapter = get_layout(prs, "6_Chapter Page / Break")
    light_content = get_layout(prs, "Light background content")

    # --- Slide 1: Cover ---
    slide = prs.slides.add_slide(chapter)
    set_placeholder_text(slide, 0, "Market Signals — {{GENERATED_AT}}")
    set_placeholder_text(slide, 15, "Research period: {{PERIOD_START}} – {{PERIOD_END}}")

    # --- Slide 2: Fleet section divider ---
    add_section_divider(prs, dark_empty, "Fleet Mobility Management", "1")

    # --- Slide 3: Fleet Trends ---
    add_content_slide(prs, dark_content,
                      title="Fleet Mobility — Top 5 Trends",
                      subtitle="{{PERIOD_START}} – {{PERIOD_END}}",
                      body="{{FLEET_TRENDS}}")

    # --- Slide 4: Fleet Competitor Moves ---
    add_content_slide(prs, dark_content,
                      title="Fleet Mobility — Competitor Moves",
                      body="{{FLEET_COMPETITOR_MOVES}}")

    # --- Slide 5: Fleet Unmet Needs ---
    add_content_slide(prs, dark_content,
                      title="Fleet Mobility — Unmet Customer Needs",
                      body="{{FLEET_UNMET_NEEDS}}")

    # --- Slide 6: Fleet Regulations ---
    add_content_slide(prs, dark_content,
                      title="Fleet Mobility — Active Regulations",
                      body="{{FLEET_ACTIVE_REGULATIONS}}")

    # --- Slide 7: Fleet Strategic Implications ---
    add_content_slide(prs, dark_content,
                      title="Fleet Mobility — Strategic Implications for Elli",
                      body="{{FLEET_STRATEGIC_IMPLICATIONS}}")

    # --- Slide 8: Site section divider ---
    add_section_divider(prs, dark_empty, "Charging Site Management", "2")

    # --- Slide 9: Site Trends ---
    add_content_slide(prs, dark_content,
                      title="Charging Site — Top 5 Trends",
                      subtitle="{{PERIOD_START}} – {{PERIOD_END}}",
                      body="{{SITE_TRENDS}}")

    # --- Slide 10: Site Competitor Moves ---
    add_content_slide(prs, dark_content,
                      title="Charging Site — Competitor Moves",
                      body="{{SITE_COMPETITOR_MOVES}}")

    # --- Slide 11: Site Unmet Needs ---
    add_content_slide(prs, dark_content,
                      title="Charging Site — Unmet Customer Needs",
                      body="{{SITE_UNMET_NEEDS}}")

    # --- Slide 12: Site Regulations ---
    add_content_slide(prs, dark_content,
                      title="Charging Site — Active Regulations",
                      body="{{SITE_ACTIVE_REGULATIONS}}")

    # --- Slide 13: Site Strategic Implications ---
    add_content_slide(prs, dark_content,
                      title="Charging Site — Strategic Implications for Elli",
                      body="{{SITE_STRATEGIC_IMPLICATIONS}}")

    # --- Slide 14: Open Questions ---
    add_content_slide(prs, light_content,
                      title="Open Questions",
                      body="{{OPEN_QUESTIONS}}")

    prs.save(OUTPUT)
    print(f"✓ Saved {OUTPUT} with {len(prs.slides)} slides")
    print("Slides:")
    for i, slide in enumerate(prs.slides, 1):
        texts = [s.text_frame.text[:60].replace('\n', ' ')
                 for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        print(f"  {i}: {' | '.join(texts[:2])}")


if __name__ == "__main__":
    main()
