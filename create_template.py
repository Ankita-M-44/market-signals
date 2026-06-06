"""
One-time script: generates template.pptx (14 slides) from Elli corporate template.
Run: python create_template.py
Then commit the resulting template.pptx to the repo.

Slide index map (must match fill_pptx() in pipeline.py):
  0  Cover              6_Chapter Page / Break
  1  Fleet divider      Dark background empty
  2  Fleet Trends       Dark background empty  <- drawn programmatically
  3  Fleet Competitors  Dark background empty  <- drawn programmatically
  4  Fleet Needs        Dark background empty  <- drawn programmatically
  5  Fleet Regulations  Dark background empty  <- drawn programmatically
  6  Fleet Implications Dark background empty  <- drawn programmatically
  7  Site divider       Dark background empty
  8  Site Trends        Dark background empty  <- drawn programmatically
  9  Site Competitors   Dark background empty  <- drawn programmatically
  10 Site Needs         Dark background empty  <- drawn programmatically
  11 Site Regulations   Dark background empty  <- drawn programmatically
  12 Site Implications  Dark background empty  <- drawn programmatically
  13 Open Questions     Dark background empty  <- drawn programmatically
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

SOURCE = "Elli_Mobility_PP_Template_2.0.pptx"
OUTPUT = "template.pptx"

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ELLI_GREEN = RGBColor(0x00, 0xD4, 0x8A)


def remove_all_slides(prs):
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


def set_ph(slide, idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            return


def add_section_divider(prs, layout, section_title, section_number):
    """Dark section break slide with large title and section number."""
    slide = prs.slides.add_slide(layout)
    tb = slide.shapes.add_textbox(
        Inches(1.92), Inches(2.29), Inches(9.52), Inches(2.41)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = section_title
    run.font.size = Pt(72)
    run.font.bold = True
    run.font.color.rgb = WHITE
    nb = slide.shapes.add_textbox(
        Inches(0.57), Inches(2.06), Inches(1.40), Inches(2.86)
    )
    nf = nb.text_frame
    para2 = nf.paragraphs[0]
    run2 = para2.add_run()
    run2.text = section_number
    run2.font.size = Pt(120)
    run2.font.bold = True
    run2.font.color.rgb = WHITE
    return slide


def main():
    prs = Presentation(SOURCE)
    remove_all_slides(prs)

    dark_empty = get_layout(prs, "Dark background empty")
    chapter    = get_layout(prs, "6_Chapter Page / Break")

    # Slide 0: Cover
    slide = prs.slides.add_slide(chapter)
    set_ph(slide, 0, "Market Signals — {{GENERATED_AT}}")
    set_ph(slide, 15, "Research period: {{PERIOD_START}} – {{PERIOD_END}}")

    # Slide 1: Fleet section divider
    add_section_divider(prs, dark_empty, "Fleet Mobility\nManagement", "1")

    # Slides 2-6: Fleet content (drawn programmatically by pipeline)
    for _ in range(5):
        prs.slides.add_slide(dark_empty)

    # Slide 7: Site section divider
    add_section_divider(prs, dark_empty, "Charging Site\nManagement", "2")

    # Slides 8-13: Site content (drawn programmatically by pipeline)
    for _ in range(6):
        prs.slides.add_slide(dark_empty)

    prs.save(OUTPUT)
    print(f"Saved {OUTPUT} with {len(prs.slides)} slides")
    for i, slide in enumerate(prs.slides):
        texts = [
            s.text_frame.text[:60].replace('\n', ' ')
            for s in slide.shapes
            if s.has_text_frame and s.text_frame.text.strip()
        ]
        print(f"  {i:2d}: {' | '.join(texts[:2]) or '(empty canvas)'}")


if __name__ == "__main__":
    main()
