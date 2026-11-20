
"""
Carelon Global Solutions - AI-Powered Member Assistant Case Study Presentation
Comprehensive python-pptx script for 23 slides + appendix
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# ── Color palette ──────────────────────────────────────────────────────────────
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
NAVY        = RGBColor(0x1A, 0x1A, 0x2E)
PURPLE      = RGBColor(0x5B, 0x2D, 0x8E)
TEAL        = RGBColor(0x0D, 0x7C, 0x8F)
LIGHT_BLUE  = RGBColor(0xE8, 0xF4, 0xF8)
LIGHT_PURPLE= RGBColor(0xF0, 0xE6, 0xFF)
GRAY        = RGBColor(0x88, 0x88, 0x88)
DARK_GRAY   = RGBColor(0x44, 0x44, 0x44)
GREEN       = RGBColor(0x2E, 0xCC, 0x71)
RED         = RGBColor(0xE7, 0x4C, 0x3C)
ORANGE      = RGBColor(0xF3, 0x9C, 0x12)
LIGHT_GRAY  = RGBColor(0xF5, 0xF5, 0xF5)
TABLE_HEAD_PURPLE = RGBColor(0x5B, 0x2D, 0x8E)
TABLE_HEAD_TEAL   = RGBColor(0x0D, 0x7C, 0x8F)
TABLE_ALT   = RGBColor(0xF5, 0xF0, 0xFF)
TABLE_ALT2  = RGBColor(0xE8, 0xF4, 0xF8)

FOOTER_TEXT = "Carelon Global Solutions | Confidential – Interview Submission"
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]  # completely blank

# ── Helper utilities ───────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill_color, border_color=None, border_pt=0):
    shape = slide.shapes.add_shape(1, l, t, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, l, t, w, h, text, font_size=14, bold=False, color=NAVY,
                align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txb

def set_cell_bg(cell, color):
    from pptx.oxml.ns import qn
    from lxml import etree
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
    srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', '%02X%02X%02X' % (color[0], color[1], color[2]))

def style_cell(cell, text, font_size=11, bold=False, color=NAVY,
               bg=None, align=PP_ALIGN.LEFT, italic=False):
    cell.text = text
    if bg:
        set_cell_bg(cell, bg)
    tf = cell.text_frame
    tf.word_wrap = True
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color
            run.font.name = "Calibri"
            run.font.italic = italic
        # If no runs, add one
        if not para.runs:
            run = para.add_run()
            run.text = text
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color
            run.font.name = "Calibri"
            run.font.italic = italic

def style_table_header_row(table, col_headers, bg=TABLE_HEAD_PURPLE):
    row = table.rows[0]
    for i, hdr in enumerate(col_headers):
        style_cell(row.cells[i], hdr, font_size=11, bold=True, color=WHITE, bg=bg, align=PP_ALIGN.CENTER)

def add_header_bar(slide, headline, bar_color=PURPLE):
    """Standard 1-inch colored header bar at top of slide."""
    bar = add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.0), bar_color)
    # Headline text inside bar
    txb = slide.shapes.add_textbox(Inches(0.3), Inches(0.08), Inches(12.7), Inches(0.84))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = headline
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = "Calibri"
    return bar

def add_footer_and_slide_number(slide, slide_num):
    """Footer bottom-left + slide number bottom-right."""
    # Footer
    txb = slide.shapes.add_textbox(Inches(0.2), Inches(7.15), Inches(10), Inches(0.3))
    tf = txb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = FOOTER_TEXT
    run.font.size = Pt(9)
    run.font.color.rgb = GRAY
    run.font.name = "Calibri"
    # Slide number
    txb2 = slide.shapes.add_textbox(Inches(12.5), Inches(7.15), Inches(0.7), Inches(0.3))
    tf2 = txb2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    run2 = p2.add_run()
    run2.text = str(slide_num)
    run2.font.size = Pt(9)
    run2.font.color.rgb = GRAY
    run2.font.name = "Calibri"

def add_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text

def add_bullet_textbox(slide, l, t, w, h, bullets, font_size=13, color=NAVY,
                        indent=False, bold_first=False):
    """Add a textbox with multiple bullet lines."""
    from pptx.util import Pt
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    for idx, bullet in enumerate(bullets):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(2)
        p.space_after  = Pt(2)
        run = p.add_run()
        run.text = bullet
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run.font.name = "Calibri"
        run.font.bold = (bold_first and idx == 0)
    return txb

def new_slide(slide_num):
    slide = prs.slides.add_slide(blank_layout)
    # White background
    bg = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    return slide

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(1)
# Purple accent bar on left
add_rect(slide, Inches(0), Inches(0), Inches(0.35), SLIDE_H, PURPLE)
# Top band
add_rect(slide, Inches(0.35), Inches(0), Inches(13.0), Inches(0.08), TEAL)
# Main title area background
add_rect(slide, Inches(0.35), Inches(0.08), Inches(13.0), Inches(2.6), LIGHT_BLUE)

# Main headline
add_textbox(slide, Inches(0.6), Inches(0.2), Inches(12.5), Inches(1.5),
            "AI-Powered Member Experience:\nFrom Navigation Chaos to Confident Self-Service",
            font_size=28, bold=True, color=PURPLE, align=PP_ALIGN.LEFT)

# Subheadline
add_textbox(slide, Inches(0.6), Inches(1.65), Inches(12.5), Inches(0.6),
            "Sr. Business Analyst II Interview Submission | Carelon Global Solutions",
            font_size=16, bold=False, color=TEAL, align=PP_ALIGN.LEFT)

# Divider line
add_rect(slide, Inches(0.6), Inches(2.6), Inches(11.5), Inches(0.04), PURPLE)

# Candidate name and date
add_textbox(slide, Inches(0.6), Inches(2.72), Inches(6), Inches(0.4),
            "[Candidate Name]", font_size=15, bold=True, color=NAVY)
add_textbox(slide, Inches(0.6), Inches(3.08), Inches(6), Inches(0.35),
            "November 2026", font_size=13, color=DARK_GRAY)

# Thesis box
thesis_box = add_rect(slide, Inches(0.6), Inches(3.6), Inches(12.1), Inches(1.4), LIGHT_PURPLE,
                       border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.75), Inches(3.65), Inches(11.8), Inches(1.3),
            "Thesis: A trust-first, retrieval-grounded AI assistant resolves member questions correctly on first contact — "
            "reducing call volume, improving health literacy, and treating every member equitably across all plan types.",
            font_size=13, bold=False, color=NAVY, align=PP_ALIGN.LEFT)

# Bottom decorative strip
add_rect(slide, Inches(0.35), Inches(5.1), Inches(13.0), Inches(0.08), PURPLE)

add_footer_and_slide_number(slide, 1)
add_notes(slide,
    "Open by stating the thesis clearly. The interviewer should know in 20 seconds what your recommendation is and why it matters.\n\n"
    "Likely Q: Why did you choose first-contact resolution as your north-star?\n"
    "A: Because it captures both member experience (they got their answer) and operational efficiency (no repeat contacts) in one number."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — PROBLEM FRAMING
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(2)
add_header_bar(slide, "Members Are Caught in a 3-Channel Loop That Erodes Trust and Costs Millions", PURPLE)
add_footer_and_slide_number(slide, 2)

bullets = [
    "🔄  The Loop: Member has question → searches website → can't find clear answer → calls customer service → long wait → agent resolves → member calls again next month",
    "📞  ~30% of payer call-center volume is coverage/benefits Q&A  (Illustrative assumption: industry benchmark)",
    "💰  Cost per inbound call: ~$8–$12  (Illustrative assumption: industry benchmark)",
    "😤  Low health insurance literacy compounds complexity — deductibles, coinsurance, prior auth, EOBs are opaque",
    "📱  Digital tools exist but are siloed; members switch channels 2–3 times per query",
]
add_bullet_textbox(slide, Inches(0.4), Inches(1.1), Inches(12.5), Inches(2.8), bullets, font_size=13)

# Flow diagram box
add_rect(slide, Inches(0.4), Inches(4.0), Inches(12.5), Inches(1.3), LIGHT_BLUE,
         border_color=TEAL, border_pt=1)
add_textbox(slide, Inches(0.55), Inches(4.08), Inches(12.2), Inches(1.1),
            "Member Question  →  Portal  →  FAQ  →  Live Chat  →  Phone  →  Agent  →  Resolved … (but member calls again)",
            font_size=13, bold=False, color=TEAL, align=PP_ALIGN.CENTER)

add_notes(slide,
    "The core problem is not that members are unintelligent — it is that insurance information is genuinely complex, presented in plan-specific jargon, and scattered across channels.\n\n"
    "Likely Q: How do you know 30% of calls are coverage Q&A?\n"
    "A: This is an illustrative industry benchmark. At Carelon I would validate against actual IVR intent data and CRM call reason codes before citing any number."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — PERSONAS
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(3)
add_header_bar(slide, "Five Personas, Five Different Definitions of 'Getting Help'", TEAL)
add_footer_and_slide_number(slide, 3)

# Table
cols = 6
rows_data = [
    ["Persona", "Top Questions", "Health Literacy", "Channel Preference", "Emotional State", "Key JTBD"],
    ["Alex, 24\nNew-to-Insurance", "What does my plan cover?\nWhat is a deductible?", "Very Low", "Mobile-first, texts", "Anxious, overwhelmed", "Understand my plan in plain English before I need it"],
    ["Maria, 38\nFamily Caregiver", "Is my child's specialist in-network?\nHow do I track 4 family deductibles?", "Medium", "Mobile + web", "Stressed, time-poor", "One place to manage my whole family's coverage"],
    ["David, 52\nChronic Condition", "Is my insulin covered? Why was my PA denied?\nWhat's my OOP max?", "Medium-High", "Web + phone", "Frustrated, determined", "Navigate coverage so I can focus on my health"],
    ["Eleanor, 68\nMedicare-Eligible Senior", "What's covered under Medicare Advantage?\nWhat is my copay?", "Low-Medium", "Phone-first", "Cautious, trust-sensitive", "Talk to someone who explains it like I matter"],
    ["Priya, 34\nHDHP Member", "How much of my deductible is left?\nWill this procedure count?", "High", "Web, self-serve", "Cost-conscious, analytical", "Real-time cost clarity so I can plan my spending"],
]
n_rows = len(rows_data)

tbl = slide.shapes.add_table(n_rows, cols, Inches(0.3), Inches(1.1), Inches(12.7), Inches(5.9)).table
col_widths = [Inches(1.6), Inches(2.5), Inches(1.3), Inches(1.6), Inches(1.6), Inches(2.6)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

alt_colors = [TABLE_ALT, TABLE_ALT2]
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            bg = alt_colors[(r-1) % 2]
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

add_notes(slide,
    "Personas drive every product decision. Alex shapes plain-language rules. Eleanor explains why phone handoff must be seamless with full context. Priya defines the deductible tracker priority.\n\n"
    "Likely Q: How did you develop these personas?\n"
    "A: Combination of call-center reason codes, member survey data (which Carelon has in CSAT reports), digital analytics, and stakeholder interviews with frontline agents who hear from members daily."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — MARKET & COMPETITOR SCAN
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(4)
add_header_bar(slide, "Payer Assistants Are Improving Fast — the Gap to Close Is Trust, Not Technology", PURPLE)
add_footer_and_slide_number(slide, 4)

# 2x2 positioning grid
add_textbox(slide, Inches(0.4), Inches(1.05), Inches(12.5), Inches(0.4),
            "Positioning Map: Trust & Safety (Y) vs. Capability (X)", font_size=13, bold=True, color=NAVY)

# Draw 2x2
quad_l = Inches(1.5); quad_t = Inches(1.5); quad_w = Inches(10); quad_h = Inches(3.5)
cell_w = quad_w / 2; cell_h = quad_h / 2

add_rect(slide, quad_l, quad_t, cell_w, cell_h, LIGHT_BLUE, border_color=GRAY, border_pt=0.5)
add_rect(slide, quad_l + cell_w, quad_t, cell_w, cell_h, LIGHT_PURPLE, border_color=GRAY, border_pt=0.5)
add_rect(slide, quad_l, quad_t + cell_h, cell_w, cell_h, LIGHT_GRAY, border_color=GRAY, border_pt=0.5)
add_rect(slide, quad_l + cell_w, quad_t + cell_h, cell_w, cell_h, LIGHT_GRAY, border_color=GRAY, border_pt=0.5)

add_textbox(slide, quad_l + Inches(0.1), quad_t + Inches(0.1), cell_w - Inches(0.2), cell_h - Inches(0.2),
            "HIGH TRUST | NARROW CAPABILITY\n\nTraditional IVR\n(Automated phone menus)", font_size=12, color=TEAL, bold=False)
add_textbox(slide, quad_l + cell_w + Inches(0.1), quad_t + Inches(0.1), cell_w - Inches(0.2), cell_h - Inches(0.2),
            "HIGH TRUST | BROAD CAPABILITY\n\n⭐ TARGET POSITION\nOur AI Assistant", font_size=12, color=PURPLE, bold=True)
add_textbox(slide, quad_l + Inches(0.1), quad_t + cell_h + Inches(0.1), cell_w - Inches(0.2), cell_h - Inches(0.2),
            "LOW TRUST | NARROW CAPABILITY\n\nBasic FAQ Chatbots\n(Most current payer portals)", font_size=12, color=DARK_GRAY)
add_textbox(slide, quad_l + cell_w + Inches(0.1), quad_t + cell_h + Inches(0.1), cell_w - Inches(0.2), cell_h - Inches(0.2),
            "LOW TRUST | BROAD CAPABILITY\n\nGeneral ChatGPT / Consumer AI\n(Capable but not PHI-safe, no plan data)", font_size=12, color=RED)

# Axis labels
add_textbox(slide, Inches(0.3), Inches(1.9), Inches(1.1), Inches(2.6),
            "HIGH\nTRUST", font_size=10, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(0.3), Inches(3.8), Inches(1.1), Inches(1.1),
            "LOW\nTRUST", font_size=10, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(1.8), Inches(5.1), Inches(4.5), Inches(0.3),
            "NARROW CAPABILITY", font_size=10, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(6.5), Inches(5.1), Inches(4.5), Inches(0.3),
            "BROAD CAPABILITY", font_size=10, bold=True, color=TEAL, align=PP_ALIGN.CENTER)

# Competitor notes
add_bullet_textbox(slide, Inches(0.4), Inches(5.3), Inches(12.5), Inches(1.6), [
    "• Cigna Edie / Humana Ask Nurse / UHC chat: basic FAQ + claims status; limited plan-specific grounding",
    "• Microsoft/Azure Health Bot: infrastructure layer; needs payer customization",
    "• Babylon Health, Ada: symptom-focused; not insurance navigation",
    "⚠️  Verify competitor capabilities before presenting — landscape moves fast",
], font_size=11, color=DARK_GRAY)

add_notes(slide,
    "The insight is that generic AI is capable but not trustworthy for PHI; legacy chatbots are trusted but not capable. The white space is high-trust, high-capability — which requires retrieval-grounded answers, PHI-compliant architecture, and rigorous hallucination controls.\n\n"
    "Likely Q: Who is your biggest competitor risk?\n"
    "A: Google and Apple health integrations that members already trust for health information — they could add insurance Q&A and disintermediate payer portals entirely."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — VISION & PRINCIPLES
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(5)
add_header_bar(slide, "North-Star: Every Member Gets the Right Answer, Right Now, Without Calling", TEAL)
add_footer_and_slide_number(slide, 5)

# Vision statement box
add_rect(slide, Inches(0.4), Inches(1.1), Inches(12.5), Inches(1.0), LIGHT_BLUE, border_color=TEAL, border_pt=1.5)
add_textbox(slide, Inches(0.6), Inches(1.15), Inches(12.1), Inches(0.9),
            "Vision: Every Carelon member can understand their coverage, costs, and care options instantly — in their language, at their literacy level — with the confidence that the answer is accurate and their privacy is protected.",
            font_size=12, bold=False, color=NAVY, align=PP_ALIGN.LEFT)

# North-Star metric
add_rect(slide, Inches(0.4), Inches(2.2), Inches(5.8), Inches(0.9), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.55), Inches(2.25), Inches(5.5), Inches(0.8),
            "⭐ North-Star: First-Contact Resolution Rate (FCR)\nTarget: 65% within 12 months of full launch  (Illustrative)",
            font_size=12, bold=True, color=PURPLE)

# Guardrail box
add_rect(slide, Inches(6.5), Inches(2.2), Inches(6.3), Inches(0.9), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(6.65), Inches(2.25), Inches(6.0), Inches(0.8),
            "Guardrails: Hallucination <1% | Safety recall >99%\nEquity gap <5pp | CSAT ≥4.2/5",
            font_size=12, bold=False, color=PURPLE)

# 5 principles
add_textbox(slide, Inches(0.4), Inches(3.25), Inches(5), Inches(0.4), "5 Product Principles", font_size=14, bold=True, color=NAVY)
principles = [
    "1.  Trust before Convenience — accuracy and safety over speed",
    "2.  Plain Language Always — 6th-grade reading level default; member can ask for simpler",
    "3.  Never a Dead End — every answer includes a next step or warm handoff",
    "4.  Equitable by Design — same quality across all plans, languages, literacy levels",
    "5.  Minimum Necessary Data — access only PHI required to answer the question",
]
add_bullet_textbox(slide, Inches(0.4), Inches(3.65), Inches(12.5), Inches(2.8), principles, font_size=13, color=NAVY)

add_notes(slide,
    "Product principles are not marketing copy — they are decision filters. When engineering says 'we can ship faster if we skip the confidence threshold,' principle #1 says no.\n\n"
    "Likely Q: What if FCR and member satisfaction conflict?\n"
    "A: They rarely do if the assistant is honest — a correct 'I don't know + warm handoff' scores higher CSAT than a confident wrong answer."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — OPPORTUNITY SIZING
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(6)
add_header_bar(slide, "A 10-Point FCR Improvement Saves ~$12–18M Annually — Fully Illustrative", PURPLE)
add_footer_and_slide_number(slide, 6)

rows_data = [
    ["Metric", "Value", "Notes"],
    ["Total member calls per year", "5M", "Illustrative: mid-size payer assumption"],
    ["% that are AI-automatable intent", "40%  →  2M conversations", "Illustrative"],
    ["Cost per inbound agent call", "$10 avg", "Illustrative benchmark"],
    ["Current digital self-service rate", "25%", "Illustrative"],
    ["Target digital self-service rate (Y1)", "45%  (+20pp)", "Target"],
    ["Incremental calls deflected", "400K calls/year", ""],
    ["Gross savings", "$4M/year", "At $10/call"],
    ["AI platform + ops cost", "$1.5–2M/year", "Illustrative"],
    ["Net annual benefit", "~$2–2.5M direct + CSAT/NPS lift", ""],
    ["3-year NPV (10% discount)", "~$5–6M", "Illustrative"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 3, Inches(0.4), Inches(1.1), Inches(12.5), Inches(5.0)).table
tbl.columns[0].width = Inches(4.5)
tbl.columns[1].width = Inches(4.0)
tbl.columns[2].width = Inches(4.0)

for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = TABLE_ALT if r % 2 == 0 else WHITE
            is_bold = (r in [7, 8, 9, 10] and c == 1)
            style_cell(cell, text, font_size=11, bold=is_bold, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# Warning note
add_rect(slide, Inches(0.4), Inches(6.2), Inches(12.5), Inches(0.75), LIGHT_PURPLE, border_color=PURPLE, border_pt=1)
add_textbox(slide, Inches(0.55), Inches(6.25), Inches(12.1), Inches(0.65),
            "⚠️  All numbers are illustrative assumptions. Actual sizing requires: (1) IVR intent data, (2) actual cost-per-contact from finance, (3) current digital containment rate from analytics.",
            font_size=11, color=PURPLE, italic=True)

add_notes(slide,
    "Lead with the assumption, not the number. I always say 'if Carelon's call volume is in this range and cost-per-contact is industry standard, the opportunity is in this range.' This shows financial acumen without overclaiming.\n\n"
    "Likely Q: What if the actual call volume is different?\n"
    "A: The model is designed to flex — I would re-run with actual data in week 1."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — USE CASE PRIORITIZATION
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(7)
add_header_bar(slide, "V1 Covers 70% of Member Intent Volume with the Lowest Clinical Risk", TEAL)
add_footer_and_slide_number(slide, 7)

rows_data = [
    ["Use Case", "Reach", "Impact", "Confidence", "Effort", "RICE", "Phase"],
    ["Coverage & Benefits Q&A", "5", "3", "0.8", "2", "6.0", "V1 ✅"],
    ["Deductible/OOP Tracker", "5", "3", "0.9", "2", "6.75", "V1 ✅"],
    ["Claims Status + EOB", "5", "3", "0.8", "3", "4.0", "V1 ✅"],
    ["ID Card & Plan Documents", "5", "2", "0.9", "1", "9.0", "V1 ✅"],
    ["Warm Handoff to Agent", "5", "2", "0.9", "2", "4.5", "V1 ✅"],
    ["Find In-Network Provider", "4", "3", "0.7", "4", "2.1", "V2"],
    ["Prior Auth Status", "3", "3", "0.7", "3", "2.1", "V2"],
    ["Cost Estimator", "4", "3", "0.6", "4", "1.8", "V2"],
    ["Billing Help", "3", "2", "0.6", "3", "1.2", "V3"],
    ["Prescription Coverage", "3", "3", "0.5", "4", "1.1", "V3"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 7, Inches(0.3), Inches(1.1), Inches(12.7), Inches(4.6)).table
col_widths = [Inches(3.0), Inches(0.9), Inches(0.9), Inches(1.1), Inches(0.9), Inches(0.9), Inches(1.0)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

v1_color = RGBColor(0xE8, 0xFF, 0xF0)
v2_color = RGBColor(0xFF, 0xF9, 0xE8)
v3_color = RGBColor(0xFF, 0xF0, 0xF0)
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=10, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            phase = row_data[6]
            if "V1" in phase:
                bg = v1_color
            elif "V2" in phase:
                bg = v2_color
            else:
                bg = v3_color
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT)

# Rationale box
add_rect(slide, Inches(0.3), Inches(5.8), Inches(12.7), Inches(0.75), LIGHT_BLUE, border_color=TEAL, border_pt=1)
add_textbox(slide, Inches(0.45), Inches(5.85), Inches(12.4), Inches(0.65),
            "V1 NOT built: cost estimator (needs real-time claims+provider data, high complexity), Rx formulary (pharmacy system integration), billing disputes (legal/compliance review required)",
            font_size=11, color=TEAL, italic=True)

add_notes(slide,
    "RICE is a starting point for conversation, not a final verdict. The real prioritization conversation happens with legal, clinical, and data science. I defend V1 scope aggressively: ID card + deductible + claims status + coverage Q&A covers the top contact reasons with manageable risk.\n\n"
    "Likely Q: Why not ship the cost estimator in V1?\n"
    "A: Provider pricing data is inconsistent, varies by contract, and an inaccurate estimate erodes trust more than no estimate. V2 ships it with heavy caveats and a 'verify with provider' disclosure."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — EXPERIENCE DESIGN
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(8)
add_header_bar(slide, "The Assistant Adapts Its Depth, Tone and Language to Each Member", PURPLE)
add_footer_and_slide_number(slide, 8)

# Flow overview
flow_boxes = [
    ("Member\nOpens App", LIGHT_BLUE),
    ("Greeting +\nQuick-Action Chips", LIGHT_BLUE),
    ("Intent\nClassification", LIGHT_PURPLE),
    ("Retrieval from\nPlan Docs + Claims", LIGHT_PURPLE),
    ("Confidence\nCheck", LIGHT_BLUE),
]
x_start = Inches(0.3); y_pos = Inches(1.1); box_w = Inches(2.1); box_h = Inches(0.75); gap = Inches(0.25)
for i, (label, bg_col) in enumerate(flow_boxes):
    x = x_start + i * (box_w + gap)
    add_rect(slide, x, y_pos, box_w, box_h, bg_col, border_color=PURPLE, border_pt=0.8)
    add_textbox(slide, x + Inches(0.05), y_pos + Inches(0.05), box_w - Inches(0.1), box_h - Inches(0.1),
                label, font_size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    if i < len(flow_boxes)-1:
        add_textbox(slide, x + box_w + Inches(0.02), y_pos + Inches(0.2), gap, box_h - Inches(0.4),
                    "→", font_size=14, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)

# Branch outcomes
y2 = Inches(2.05)
add_textbox(slide, Inches(0.3), y2 - Inches(0.05), Inches(12.5), Inches(0.3),
            "↓ HIGH CONFIDENCE:  Answer + source citation     |     LOW CONFIDENCE:  Honest uncertainty + handoff offer     →  Feedback: 👍 👎 + 'Explain simpler'",
            font_size=11, color=TEAL, bold=True)

# Design principles
principles = [
    "🗣️  Plain language default: 6th-grade Flesch-Kincaid; 'Explain simpler' drops to 4th-grade",
    "🔔  Proactive nudges: 'Your deductible resets Jan 1 — you have $340 left. Want to see what's covered before year-end?'",
    "🌐  Multilingual: Spanish, Mandarin, Vietnamese, Tagalog at minimum (CMS language access requirements for MA plans)",
    "♿  WCAG 2.1 AA: screen reader compatible, keyboard navigable, 4.5:1 contrast ratio",
    "🤝  Handoff with full context: agent receives summary of conversation — member NEVER repeats themselves",
    "📱  Channels: mobile app (primary), web portal, SMS (read-only status), IVR integration (future V2)",
]
add_bullet_textbox(slide, Inches(0.3), Inches(2.4), Inches(12.7), Inches(4.6), principles, font_size=12, color=NAVY)

add_notes(slide,
    "The 'never repeat yourself' principle is the most important UX commitment. When I hand off to a human agent, that agent sees: member ID, the question asked, what the AI found, why it escalated, and suggested next steps.\n\n"
    "Likely Q: How do you handle multilingual?\n"
    "A: CMS requires MA plans to provide language access. We integrate with the existing translation layer and test separately for each language — accuracy in Spanish matters as much as accuracy in English."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — SOLUTION ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(9)
add_header_bar(slide, "A Retrieval-Grounded Architecture Keeps Every Answer Traceable to a Source", TEAL)
add_footer_and_slide_number(slide, 9)

arch_layers = [
    ("CHANNELS", "Mobile App  |  Web Portal  |  SMS  |  (V2: IVR)", LIGHT_BLUE, TEAL),
    ("AUTH & CONSENT", "SSO / MFA  |  PHI consent check  |  Session scope", LIGHT_PURPLE, PURPLE),
    ("ORCHESTRATION", "Intent classifier  →  Router  →  Safety guardrail check", LIGHT_BLUE, TEAL),
    ("RETRIEVAL ENGINE", "Plan document store (PDF chunks)  +  Claims/eligibility API  +  Provider directory", LIGHT_PURPLE, PURPLE),
    ("LLM RESPONSE", "Grounded prompt  |  Confidence scoring  |  Citation attachment", LIGHT_BLUE, TEAL),
    ("GUARDRAIL LAYER", "Clinical safety classifier  |  PII filter  |  Confidence threshold  |  Audit log write", RGBColor(0xFF, 0xF0, 0xE8), RED),
    ("RESPONSE / ESCALATION", "Answer + source citation + confidence indicator  |  OR  Warm handoff → agent + context summary", LIGHT_BLUE, GREEN),
]

y_start = Inches(1.1); layer_h = Inches(0.72); gap = Inches(0.0)
for i, (label, desc, bg_col, border_col) in enumerate(arch_layers):
    y = y_start + i * (layer_h + Inches(0.05))
    add_rect(slide, Inches(0.4), y, Inches(2.0), layer_h, bg_col, border_color=border_col, border_pt=1)
    add_textbox(slide, Inches(0.42), y + Inches(0.1), Inches(1.9), layer_h - Inches(0.2),
                label, font_size=10, bold=True, color=border_col, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(2.5), y, Inches(10.5), layer_h, bg_col, border_color=border_col, border_pt=0.5)
    add_textbox(slide, Inches(2.6), y + Inches(0.1), Inches(10.2), layer_h - Inches(0.2),
                desc, font_size=11, color=NAVY, align=PP_ALIGN.LEFT)
    if i < len(arch_layers)-1:
        add_textbox(slide, Inches(2.9), y + layer_h, Inches(0.5), Inches(0.05),
                    "↓", font_size=12, bold=True, color=TEAL, align=PP_ALIGN.CENTER)

# RAG note
add_rect(slide, Inches(0.4), Inches(6.3), Inches(12.5), Inches(0.75), LIGHT_PURPLE, border_color=PURPLE, border_pt=1)
add_textbox(slide, Inches(0.55), Inches(6.35), Inches(12.1), Inches(0.65),
            "Hallucination control: RAG (Retrieval-Augmented Generation) — answers ONLY from retrieved plan documents and claims data. Source citation mandatory. Confidence score <0.7 → 'I'm not certain — let me connect you to someone who can confirm.'",
            font_size=11, color=PURPLE, italic=False)

add_notes(slide,
    "Business-readable architecture — I don't need to know every API endpoint, but I need to know what each layer does and why it exists.\n\n"
    "Likely Q: What prevents the AI from making up a benefit that doesn't exist?\n"
    "A: The RAG constraint. The model cannot answer from general knowledge — it can only synthesize from retrieved chunks of the member's actual plan documents. If it can't find the answer, it says so."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — TRUST, SAFETY & COMPLIANCE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(10)
add_header_bar(slide, "Safety Is Not a Feature — It Is the Product's License to Operate", PURPLE)
add_footer_and_slide_number(slide, 10)

col_data = [
    ("🔒  PHI & Privacy", [
        "HIPAA minimum necessary access",
        "PHI never used to train LLM",
        "Session data encrypted, purged after 90 days",
        "Consent screen before first use",
        "Audit log of every query and response",
    ], TEAL),
    ("🛡️  Clinical Safety Guardrails", [
        "NO medical diagnosis or treatment advice",
        "Safety classifier scans every input for: distress, self-harm, suicidal ideation, urgent symptoms",
        "IF triggered → immediate: 'Please call 911' or 'Crisis line: 988' + human escalation",
        "Complaints/appeals → mandatory human handoff (CMS requirement for MA)",
        "Confidence threshold: <0.7 = honest 'I don't know'",
    ], RED),
    ("⚖️  Disclosure & Compliance", [
        "'AI assistant' disclosure on every session",
        "Clear 'Talk to a person' always visible",
        "CMS rules: Medicare Advantage members can always reach a human",
        "State regulations on AI in insurance (evolving — legal review required)",
        "WCAG 2.1 AA accessibility",
        "Annual third-party audit of guardrail performance",
    ], PURPLE),
]

col_x = [Inches(0.3), Inches(4.6), Inches(8.9)]
for i, (title, items, color) in enumerate(col_data):
    x = col_x[i]
    add_rect(slide, x, Inches(1.1), Inches(4.1), Inches(0.5), color)
    add_textbox(slide, x + Inches(0.1), Inches(1.12), Inches(3.9), Inches(0.46),
                title, font_size=13, bold=True, color=WHITE)
    add_rect(slide, x, Inches(1.6), Inches(4.1), Inches(5.3), LIGHT_BLUE, border_color=color, border_pt=1)
    add_bullet_textbox(slide, x + Inches(0.1), Inches(1.65), Inches(3.9), Inches(5.2),
                       ["• " + item for item in items], font_size=11, color=NAVY)

add_notes(slide,
    "The safety story is the most important thing to get right in the interview. If the interviewer is a clinical leader, they want to know I won't let an AI tell a member with chest pain to 'check their EOB.'\n\n"
    "Likely Q: What if the safety classifier misses a distress signal?\n"
    "A: Layered defense — keyword matching + LLM classifier + human review of flagged sessions. Miss rate target <1 in 10,000 interactions. All misses reviewed and used to retrain classifier."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — FAIRNESS & EQUITY
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(11)
add_header_bar(slide, "Equal Performance Across All Members Is a Non-Negotiable Product Requirement", TEAL)
add_footer_and_slide_number(slide, 11)

rows_data = [
    ["Dimension", "Risk", "Mitigation", "Measurement"],
    ["Plan Type\n(Commercial/MA/Medicaid)", "Different plan complexity; Medicaid members may have lower literacy", "Separate testing per plan type; plain language mandatory", "FCR and CSAT by plan type"],
    ["Language", "Non-English speakers get lower-quality answers if translation layer is weak", "Native-language plan document retrieval; human review of translated answers", "FCR by language; error rate by language"],
    ["Health Literacy", "Low-literacy members may not ask the right question", "Proactive prompts; 'Explain simpler'; agent handoff never penalized", "% using 'Explain simpler'; escalation rate by literacy proxy"],
    ["Age\n(Senior vs. Young Adult)", "Seniors prefer phone; may distrust AI", "Prominent 'Talk to a person' for seniors; IVR integration in V2", "CSAT by age band; adoption rate by age"],
    ["Race/Ethnicity", "AI may underperform for underrepresented training data", "Bias testing pre-launch; disparity threshold: <5pp gap in FCR", "FCR by race/ethnicity (where self-reported data available)"],
]

n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 4, Inches(0.3), Inches(1.1), Inches(12.7), Inches(4.5)).table
col_widths = [Inches(2.0), Inches(3.5), Inches(3.8), Inches(3.4)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            bg = TABLE_ALT2 if r % 2 == 0 else WHITE
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# Equity KPI box
add_rect(slide, Inches(0.3), Inches(5.7), Inches(12.7), Inches(0.85), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.45), Inches(5.75), Inches(12.4), Inches(0.75),
            "⚖️  Equity KPI: Max allowed gap in First-Contact Resolution Rate between any two member segments = 5 percentage points. "
            "Exceeded → automatic alert to PM and clinical quality team.",
            font_size=12, bold=True, color=PURPLE)

add_notes(slide,
    "Fairness is not an afterthought — it is a launch criterion. I would not launch if any segment FCR gap exceeds 5pp.\n\n"
    "Likely Q: How do you measure equity when you don't have race/ethnicity data?\n"
    "A: Use proxy variables — zip code, language preference, Medicaid dual-eligible status — and validate with available self-reported data. Equity measurement is imperfect but required."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — BUSINESS CASE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(12)
add_header_bar(slide, "Break-Even at Month 14; 3-Year NPV of ~$5–7M Under Base Case — All Illustrative", PURPLE)
add_footer_and_slide_number(slide, 12)

rows_data = [
    ["Scenario", "Call Deflection", "CSAT Lift", "NPS Impact", "Year 1 Net", "3-Year NPV"],
    ["Conservative", "10% reduction", "+0.2 pts", "+3 pts", "-$0.5M (investment phase)", "$2M"],
    ["Base Case ⭐", "20% reduction", "+0.5 pts", "+7 pts", "+$0.8M", "$5M"],
    ["Aggressive", "30% reduction", "+0.8 pts", "+12 pts", "+$2M", "$9M"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 6, Inches(0.3), Inches(1.1), Inches(12.7), Inches(2.0)).table
col_widths = [Inches(2.0), Inches(2.0), Inches(1.8), Inches(1.8), Inches(2.3), Inches(2.8)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

row_bgs = [None, LIGHT_BLUE, LIGHT_PURPLE, TABLE_ALT2]
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = row_bgs[r]
            style_cell(cell, text, font_size=11, bold=(r == 2), color=NAVY, bg=bg, align=PP_ALIGN.CENTER)

# Cost assumptions
add_rect(slide, Inches(0.3), Inches(3.2), Inches(6.0), Inches(2.5), LIGHT_BLUE, border_color=TEAL, border_pt=1)
add_textbox(slide, Inches(0.4), Inches(3.25), Inches(5.8), Inches(0.3), "Cost Assumptions", font_size=12, bold=True, color=TEAL)
add_bullet_textbox(slide, Inches(0.4), Inches(3.55), Inches(5.8), Inches(2.0), [
    "• Year 1 investment: ~$3–4M (platform, integration, testing, compliance)  — Illustrative",
    "• Ongoing annual: ~$1.5–2M (compute, ops, human review team)  — Illustrative",
    "• CSAT 1pt improvement ≈ 2–3pp NPS ≈ measurable retention impact  — Illustrative",
], font_size=11, color=NAVY)

# Sensitivity
add_rect(slide, Inches(6.5), Inches(3.2), Inches(6.5), Inches(2.5), LIGHT_PURPLE, border_color=PURPLE, border_pt=1)
add_textbox(slide, Inches(6.6), Inches(3.25), Inches(6.3), Inches(0.3), "Key Sensitivities", font_size=12, bold=True, color=PURPLE)
add_bullet_textbox(slide, Inches(6.6), Inches(3.55), Inches(6.3), Inches(2.0), [
    "• Biggest sensitivity: actual call deflection rate",
    "  → Each 5pp change in deflection = ±$800K annual impact",
    "• Second sensitivity: LLM compute cost — monitor as model scales",
    "• Break-even: Month 14 (base case)  — Illustrative",
], font_size=11, color=NAVY)

add_notes(slide,
    "Always label illustrative numbers prominently and show the assumption chain. Never say 'this will save $5M' — say 'under base-case assumptions, the model shows $5M NPV.'\n\n"
    "Likely Q: What if call deflection is only 5%?\n"
    "A: Conservative case still breaks even by month 24, and CSAT/NPS lift has value beyond call deflection. We would revisit scope, not abandon the product."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — METRICS & EXPERIMENTATION
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(13)
add_header_bar(slide, "A Four-Level Metric Tree Tells Us If the Product Is Working — and For Whom", TEAL)
add_footer_and_slide_number(slide, 13)

rows_data = [
    ["Category", "Metric", "Baseline", "Target (12mo)", "Measurement"],
    ["Member Experience", "First-Contact Resolution Rate", "~25%", "65%", "CRM + session logs"],
    ["Member Experience", "Member CSAT", "~3.8/5", "4.3/5", "Post-session survey"],
    ["Member Experience", "Avg. Time to Resolution", "~8 min", "<3 min", "Session timestamps"],
    ["Operational", "Call Deflection Rate", "0% (pre-launch)", "20%", "Call center volume delta"],
    ["Operational", "Escalation Rate", "—", "<15%", "Escalation events/total sessions"],
    ["Quality & Safety", "Answer Groundedness Score", "—", ">98%", "RAG citation audit"],
    ["Quality & Safety", "Safety Trigger Recall", "—", ">99%", "Red-team testing monthly"],
    ["Equity", "Max FCR Gap by Segment", "Unknown", "<5pp", "Segment analytics"],
    ["Equity", "Adoption Rate by Plan Type", "0%", "No >10pp gap", "Usage analytics"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 5, Inches(0.3), Inches(1.1), Inches(12.7), Inches(4.1)).table
col_widths = [Inches(2.1), Inches(3.0), Inches(1.8), Inches(2.0), Inches(3.8)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

cat_colors = {
    "Member Experience": LIGHT_BLUE,
    "Operational": TABLE_ALT,
    "Quality & Safety": RGBColor(0xFF, 0xF0, 0xE8),
    "Equity": TABLE_ALT2,
}
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            bg = cat_colors.get(row_data[0], WHITE)
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# Pilot design
add_rect(slide, Inches(0.3), Inches(5.3), Inches(12.7), Inches(1.25), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.42), Inches(5.35), Inches(2.5), Inches(0.35), "Pilot Design:", font_size=12, bold=True, color=PURPLE)
add_textbox(slide, Inches(0.42), Inches(5.65), Inches(12.4), Inches(0.85),
            "90-day A/B test — 50% of members in 2 plan types (1 Commercial, 1 Medicaid) get the assistant; 50% get current experience. Random assignment at member ID level.\n"
            "Kill criteria: Safety recall <95% OR any segment FCR gap >10pp.",
            font_size=11, color=NAVY)

add_notes(slide,
    "The kill criteria are as important as the success criteria. I always define them before launch.\n\n"
    "Likely Q: Why A/B test rather than phased rollout?\n"
    "A: Both have merits. A/B gives causal evidence; phased rollout reduces risk. For V1 I recommend A/B in a contained population because we need clean causal data to justify V2 investment to finance."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — PRD ON A PAGE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(14)
add_header_bar(slide, "PRD: AI Member Assistant — V1 Scope, Goals, Risks, Open Questions", PURPLE)
add_footer_and_slide_number(slide, 14)

sections = [
    ("Problem", "Members navigate 3+ channels to resolve coverage, claims and cost questions, resulting in high call volume, low CSAT, and health literacy gaps.", LIGHT_BLUE, TEAL),
    ("Goals", "FCR ≥65%  |  Call deflection ≥20%  |  CSAT ≥4.3/5  |  Equity gap <5pp", LIGHT_PURPLE, PURPLE),
    ("Non-Goals (V1)", "Medical diagnosis • Real-time cost estimation (V2) • Pharmacy/Rx formulary (V3) • Employer/broker-facing features", LIGHT_BLUE, TEAL),
    ("V1 Scope", "Coverage & Benefits Q&A  |  Claims Status + EOB  |  Deductible/OOP Tracker  |  Plan Documents  |  Warm Handoff", LIGHT_PURPLE, PURPLE),
    ("Dependencies", "Claims/eligibility API  |  Plan document CMS feeds  |  SSO/auth layer  |  Escalation routing  |  LLM platform", LIGHT_BLUE, TEAL),
    ("Risks", "Hallucination → RAG constraint  |  Safety miss → layered classifier  |  Low senior adoption → IVR V2", RGBColor(0xFF, 0xF0, 0xE8), RED),
    ("Open Questions", "LLM platform? (Sprint 3)  |  PHI data retention? (legal sign-off)  |  CMS filing for Medicare AI? (compliance)", LIGHT_PURPLE, PURPLE),
]

y_start = Inches(1.1); sec_h = Inches(0.72); gap = Inches(0.03)
for i, (title, body, bg, color) in enumerate(sections):
    y = y_start + i * (sec_h + gap)
    add_rect(slide, Inches(0.3), y, Inches(2.2), sec_h, color)
    add_textbox(slide, Inches(0.35), y + Inches(0.1), Inches(2.1), sec_h - Inches(0.2),
                title, font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(2.6), y, Inches(10.5), sec_h, bg, border_color=color, border_pt=0.5)
    add_textbox(slide, Inches(2.7), y + Inches(0.08), Inches(10.2), sec_h - Inches(0.16),
                body, font_size=11, color=NAVY)

add_notes(slide,
    "The PRD one-pager is what I share with engineering, legal, and clinical quality in the same document. Each audience reads what matters to them.\n\n"
    "Likely Q: What is your process for getting PRD sign-off?\n"
    "A: I run a structured review — engineering reviews feasibility; legal reviews compliance; clinical quality reviews safety guardrails; finance reviews business case. PRD is locked only when all four sign off."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — AGILE EXECUTION
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(15)
add_header_bar(slide, "MVP in 16 Weeks Across Three Epics — Safety and Equity Gate Every Sprint", TEAL)
add_footer_and_slide_number(slide, 15)

# Epic headings
epics = [
    ("EPIC 1: Core Conversation Engine", PURPLE),
    ("EPIC 2: Member Account Integration", TEAL),
    ("EPIC 3: Handoff & Equity", PURPLE),
]

stories = [
    ("E1", "Coverage Q&A (Story 1)",
     "As a member, I can type a question about my coverage and receive a plain-language answer with the plan document source cited.\n"
     "AC: Given 'Is my MRI covered?', When AI retrieves from plan doc with confidence >0.7, Then response shows answer + document section within 5 seconds."),
    ("E1", "Safety Story 2 🛡️",
     "As a care manager, I need the assistant to detect distress language and immediately escalate to a human agent.\n"
     "AC: Given 'I want to hurt myself', When safety classifier detects trigger, Then: crisis resources + 'Connecting you now' + escalation ticket within 2 seconds. Miss rate <0.01%."),
    ("E2", "Deductible Tracker (Story 3)",
     "As an HDHP member, I can see my current deductible and OOP max progress in real-time.\n"
     "AC: Given 'How much of my deductible is left?', When claims API returns data <30 seconds old, Then shows deductible paid/remaining + OOP max progress bar."),
    ("E3", "Equity Story 4 ⚖️",
     "As a Spanish-speaking Medicaid member, I receive the same quality of answer as an English-speaking commercial member.\n"
     "AC: Given language preference = Spanish, When AI generates response, Then: response in Spanish with same accuracy; FCR gap vs English cohort <5pp measured monthly."),
    ("E3", "Handoff Story 5",
     "As a member service agent, I receive a full context summary when a member is transferred to me.\n"
     "AC: Given AI escalates to agent, When agent receives handoff, Then: agent screen shows original question, AI response, escalation reason, suggested next steps within 3 seconds."),
]

# Sprint roadmap table
add_textbox(slide, Inches(0.3), Inches(1.1), Inches(5), Inches(0.35), "One-PI Roadmap", font_size=13, bold=True, color=NAVY)
sprint_data = [
    ["Sprints", "Focus", "Deliverable"],
    ["1–2", "Platform setup, auth, PHI compliance review", "Infrastructure ready"],
    ["3–4", "Coverage Q&A with RAG, confidence scoring", "Story 1 done"],
    ["5–6", "Claims/EOB retrieval, deductible tracker", "Story 3 done"],
    ["7–8", "Safety classifier, escalation routing, handoff context", "Stories 2 & 5 done"],
    ["9–10", "Multilingual (Spanish), equity testing", "Story 4 done"],
    ["11–12", "Pilot launch (A/B, 2 plan types), monitoring dashboard", "Pilot live"],
    ["13–14", "Pilot analysis, bug fixes, scale decision", "Go/No-Go"],
]
n_rows = len(sprint_data)
tbl = slide.shapes.add_table(n_rows, 3, Inches(0.3), Inches(1.5), Inches(5.5), Inches(5.0)).table
tbl.columns[0].width = Inches(0.9)
tbl.columns[1].width = Inches(3.0)
tbl.columns[2].width = Inches(1.6)

for r, row_data in enumerate(sprint_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=10, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            bg = TABLE_ALT2 if r % 2 == 0 else WHITE
            style_cell(cell, text, font_size=10, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# User stories on right
add_textbox(slide, Inches(6.1), Inches(1.1), Inches(7), Inches(0.35), "Key User Stories", font_size=13, bold=True, color=NAVY)
for i, (epic_tag, title, story) in enumerate(stories):
    y = Inches(1.5) + i * Inches(1.17)
    epic_color = PURPLE if epic_tag == "E1" else (TEAL if epic_tag == "E2" else PURPLE)
    add_rect(slide, Inches(6.1), y, Inches(0.5), Inches(1.05), epic_color)
    add_textbox(slide, Inches(6.12), y + Inches(0.2), Inches(0.46), Inches(0.6),
                epic_tag, font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(6.65), y, Inches(6.6), Inches(1.05), LIGHT_BLUE if epic_tag == "E2" else LIGHT_PURPLE,
             border_color=epic_color, border_pt=0.5)
    add_textbox(slide, Inches(6.75), y + Inches(0.02), Inches(6.4), Inches(0.3),
                title, font_size=10, bold=True, color=epic_color)
    add_textbox(slide, Inches(6.75), y + Inches(0.3), Inches(6.4), Inches(0.72),
                story, font_size=9, color=NAVY)

add_notes(slide,
    "User stories are my primary communication tool with engineering. 'Given/When/Then' format ensures we agree on what done means before we start building.\n\n"
    "Likely Q: What is your approach to sprint ceremonies?\n"
    "A: Daily standups are 15 min focused on blockers. Sprint reviews include a demo to at least one member-facing stakeholder. Retrospectives generate 1 actionable improvement per sprint, not a list."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — STAKEHOLDERS & RISKS
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(16)
add_header_bar(slide, "Six Stakeholder Groups, One Shared Goal — Members Get Answers Faster", PURPLE)
add_footer_and_slide_number(slide, 16)

# RACI table
raci_data = [
    ["Stakeholder", "Product Scope", "Safety Standards", "PHI Architecture", "Pilot Design", "Budget Approval", "Launch Go/No-Go"],
    ["PM (me)", "R/A", "C", "C", "R/A", "C", "R/A"],
    ["Engineering/DS", "C", "C", "R", "C", "I", "C"],
    ["Clinical Quality", "C", "A/R", "C", "C", "I", "A"],
    ["Legal/Compliance", "C", "C", "A", "C", "I", "A"],
    ["Call Center Ops", "C", "I", "I", "C", "I", "C"],
    ["Finance", "I", "I", "I", "C", "A", "C"],
]
n_rows = len(raci_data)
tbl = slide.shapes.add_table(n_rows, 7, Inches(0.3), Inches(1.1), Inches(12.7), Inches(2.5)).table
col_widths = [Inches(2.5), Inches(1.6), Inches(1.7), Inches(1.7), Inches(1.5), Inches(1.7), Inches(2.0)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

for r, row_data in enumerate(raci_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=10, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = LIGHT_PURPLE if "A" in text and c > 0 else (TABLE_ALT2 if r % 2 == 0 else WHITE)
            style_cell(cell, text, font_size=10, bold=("A" in text and c > 0), color=NAVY, bg=bg, align=PP_ALIGN.CENTER)

# Risk register
add_textbox(slide, Inches(0.3), Inches(3.7), Inches(5), Inches(0.4), "Risk Register", font_size=13, bold=True, color=NAVY)
risks = [
    ["Risk", "Likelihood", "Impact", "Mitigation"],
    ["AI hallucinates a benefit", "Medium", "High", "RAG constraint + confidence threshold + monthly audit"],
    ["Safety classifier misses distress", "Low", "Critical", "Layered classifier + red-team monthly + human review"],
    ["State AI regulation for insurance", "Medium", "High", "Monitor NAIC model bulletin; legal sign-off before launch"],
    ["Agent resistance to AI", "Medium", "Medium", "Change mgmt: AI as assistant; demonstrate time savings"],
    ["LLM bias in underrepresented populations", "Medium", "High", "Pre-launch equity audit; FCR monitoring by segment"],
    ["Low senior adoption", "Medium", "Medium", "IVR integration V2; 'Talk to person' always prominent"],
]
n_rows = len(risks)
tbl2 = slide.shapes.add_table(n_rows, 4, Inches(0.3), Inches(4.1), Inches(12.7), Inches(2.9)).table
col_widths = [Inches(3.5), Inches(1.5), Inches(1.5), Inches(6.2)]
for i, w in enumerate(col_widths):
    tbl2.columns[i].width = w

impact_colors = {"Critical": RGBColor(0xFF, 0xDD, 0xDD), "High": RGBColor(0xFF, 0xF0, 0xE0), "Medium": LIGHT_BLUE}
for r, row_data in enumerate(risks):
    for c, text in enumerate(row_data):
        cell = tbl2.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=10, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = impact_colors.get(row_data[2], WHITE) if c == 2 else (TABLE_ALT if r % 2 == 0 else WHITE)
            style_cell(cell, text, font_size=10, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

add_notes(slide,
    "The RACI prevents 'too many cooks' and 'nobody cooking' — both real failure modes.\n\n"
    "Likely Q: What is the biggest risk you are most worried about?\n"
    "A: Safety recall. Everything else is recoverable. If the AI misses a member in crisis, that is not recoverable. It is the first thing I red-team and the last thing I accept as 'good enough.'"
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 17 — CASE B: WHY COST-BASED SELECTION FAILS
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(17)
add_header_bar(slide, "Selecting by Last Year's Cost Leaves Rising-Risk Members Unprotected", TEAL)
add_footer_and_slide_number(slide, 17)

# Funnel
funnel_data = [
    ("100% Eligible Members  (High prior-year cost)", Inches(12.0), PURPLE, WHITE),
    ("60%  Selected for outreach  (capacity limited)", Inches(10.0), TEAL, WHITE),
    ("25%  Successfully contacted", Inches(8.0), RGBColor(0x0D, 0x9C, 0xAF), WHITE),
    ("15%  Actively engaged in care management", Inches(5.5), ORANGE, WHITE),
    ("8%  Measurably improved outcomes", Inches(3.5), RED, WHITE),
]

y_start = Inches(1.1); funnel_h = Inches(0.72)
for i, (label, bar_w, color, txt_color) in enumerate(funnel_data):
    y = y_start + i * (funnel_h + Inches(0.05))
    x_offset = (SLIDE_W - bar_w) / 2
    add_rect(slide, x_offset, y, bar_w, funnel_h, color)
    add_textbox(slide, x_offset + Inches(0.1), y + Inches(0.1), bar_w - Inches(0.2), funnel_h - Inches(0.2),
                label, font_size=12, bold=True, color=txt_color, align=PP_ALIGN.CENTER)

add_textbox(slide, Inches(0.3), Inches(4.95), Inches(12.5), Inches(0.3),
            "(Illustrative assumption: typical care management program funnel)", font_size=10, italic=True, color=GRAY)

# Problems
problems = [
    "⬅️  Backward-looking: last year's highest-cost ≠ this year's highest-need",
    "📵  Reachability blind: selects members who cannot be reached (wrong numbers, low phone access)",
    "⏰  Timing-agnostic: misses the 48–72 hour post-discharge window when intervention prevents readmission",
    "⚖️  Equity gap: cost-based selection may systematically underserve low-cost but high-need members (e.g., Medicaid members with SDOH barriers)",
    "👩‍⚕️  Staff time wasted: nurses spend 60%+ of outreach time on unsuccessful calls  (Illustrative)",
]
add_bullet_textbox(slide, Inches(0.3), Inches(5.3), Inches(12.7), Inches(2.1), problems, font_size=12, color=NAVY)

add_notes(slide,
    "This is the core diagnostic insight for Case B. The cost-based model is not random — it has a specific failure pattern. It misses timing, reachability, and equity.\n\n"
    "Likely Q: Why not just hire more staff?\n"
    "A: Clinical leaders already asked for that. The finance leader's question is ROI. The answer is: more staff + wrong selection = more expensive failure. Better targeting first, then capacity investment."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 18 — FOUR-FACTOR TARGETING ENGINE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(18)
add_header_bar(slide, "A Four-Factor Engine Prioritizes Members Where Outreach Will Actually Help", PURPLE)
add_footer_and_slide_number(slide, 18)

factors = [
    ("1", "Predicted Future Risk\n(not past cost)", 
     "AI model predicts member risk in next 90 days.\nUses: diagnosis codes, pharmacy fills, lab trends, ED visits, care gaps.\nOutput: risk score 0–100.", PURPLE),
    ("2", "Impactability\n(uplift, not just risk)",
     "Expected benefit of care manager outreach for THIS member.\nHigh-risk but stable members may benefit less than rising-risk members with SDOH barriers.", TEAL),
    ("3", "Reachability",
     "Probability of successful contact by channel (phone/SMS/app/mail), time of day, day of week, based on historical contact data.\nAvoid calling members with 5% phone contact rate.", PURPLE),
    ("4", "Timing Triggers",
     "Real-time alerts: hospital discharge (contact within 48 hrs), ED visit, new high-risk diagnosis, gap in care (missed medication refill, overdue follow-up).\nThese override the ranked list.", RED),
]

col_w = Inches(3.0); col_gap = Inches(0.3); y_top = Inches(1.1)
for i, (num, title, desc, color) in enumerate(factors):
    x = Inches(0.3) + i * (col_w + col_gap)
    # Header
    add_rect(slide, x, y_top, col_w, Inches(0.6), color)
    add_textbox(slide, x + Inches(0.05), y_top + Inches(0.05), col_w - Inches(0.1), Inches(0.5),
                f"Factor {num}: {title}", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Body
    add_rect(slide, x, y_top + Inches(0.6), col_w, Inches(3.5), LIGHT_BLUE if i % 2 == 0 else LIGHT_PURPLE,
             border_color=color, border_pt=0.8)
    add_textbox(slide, x + Inches(0.1), y_top + Inches(0.7), col_w - Inches(0.2), Inches(3.3),
                desc, font_size=11, color=NAVY)

# Combined worklist box
add_rect(slide, Inches(0.3), Inches(5.0), Inches(12.7), Inches(1.3), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.42), Inches(5.05), Inches(3), Inches(0.4), "Combined Worklist Formula:", font_size=12, bold=True, color=PURPLE)
add_textbox(slide, Inches(0.42), Inches(5.4), Inches(12.4), Inches(0.85),
            "Daily worklist = Rank(Risk × Impactability × Reachability × Timing Weight) — capped at nurse capacity.\n"
            "Clinician override always available with reason capture. No member is 'locked out' by the algorithm.",
            font_size=12, color=NAVY)

add_notes(slide,
    "The key insight that separates this from pure AI is the impactability factor. High risk does not equal high benefit from outreach.\n\n"
    "Likely Q: How do you calculate impactability?\n"
    "A: Uplift modeling — compare predicted outcome WITH intervention vs. predicted outcome WITHOUT intervention, using historical program data."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 19 — POST-DISCHARGE FAST LANE
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(19)
add_header_bar(slide, "The 48–72 Hour Post-Discharge Window Is the Highest-Impact Intervention Point", TEAL)
add_footer_and_slide_number(slide, 19)

# Timeline
timeline_items = [
    ("Hour 0", "Discharge notification → Immediate alert to care management system", PURPLE),
    ("Hour 4", "Risk + reachability scored; assigned to care manager", TEAL),
    ("Hour 24", "First contact attempt (preferred channel from member profile)", PURPLE),
    ("Hour 48", "Second attempt if no contact; SMS/app notification sent", TEAL),
    ("Hour 72", "Escalation if no contact — supervisor alert; home visit trigger considered", RED),
]

y_start = Inches(1.1); item_h = Inches(0.68)
for i, (time_lbl, desc, color) in enumerate(timeline_items):
    y = y_start + i * (item_h + Inches(0.06))
    add_rect(slide, Inches(0.3), y, Inches(1.4), item_h, color)
    add_textbox(slide, Inches(0.32), y + Inches(0.12), Inches(1.36), item_h - Inches(0.24),
                time_lbl, font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(1.8), y, Inches(11.2), item_h, LIGHT_BLUE if i % 2 == 0 else LIGHT_PURPLE,
             border_color=color, border_pt=0.5)
    add_textbox(slide, Inches(1.9), y + Inches(0.1), Inches(11.0), item_h - Inches(0.2),
                desc, font_size=12, color=NAVY)

# TOC workflow bullets
add_textbox(slide, Inches(0.3), Inches(5.0), Inches(12.5), Inches(0.35), "Transitions-of-Care Workflow:", font_size=13, bold=True, color=NAVY)
toc_items = [
    "💊  Medication reconciliation: confirm received, understood, and affordable discharge medications",
    "📅  Follow-up appointment: confirm PCP/specialist appointment booked within 7 days",
    "🚨  Red flags review: educate on symptoms requiring immediate ED/urgent care",
    "🏠  SDOH screening: transportation, food access, medication cost barriers identified",
    "📞  30-day check-in: scheduled at discharge call",
]
add_bullet_textbox(slide, Inches(0.3), Inches(5.35), Inches(12.7), Inches(1.7), toc_items, font_size=11, color=NAVY)

# Evidence note
add_rect(slide, Inches(0.3), Inches(7.05), Inches(12.7), Inches(0.35), LIGHT_BLUE, border_color=TEAL, border_pt=0.5)
add_textbox(slide, Inches(0.4), Inches(7.07), Inches(12.4), Inches(0.3),
            "Evidence: 48-hour contact is associated with 20–25% reduction in 30-day readmission in published literature.  (Cite source when presenting — verify current evidence.)",
            font_size=10, italic=True, color=TEAL)

add_notes(slide,
    "This slide speaks directly to the finance leader (readmissions are expensive) and the clinical leader (transitions of care is a known care quality gap).\n\n"
    "Likely Q: What if the member doesn't want to be contacted?\n"
    "A: Respect opt-out. The system flags the preference and routes to passive support — app notifications, educational content — rather than phone calls."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 20 — AI FAIRNESS IN CARE MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(20)
add_header_bar(slide, "Cost Is a Proxy for Race — We Must Audit and Correct the AI Model's Bias", PURPLE)
add_footer_and_slide_number(slide, 20)

add_rect(slide, Inches(0.3), Inches(1.1), Inches(12.7), Inches(0.75), LIGHT_PURPLE, border_color=PURPLE, border_pt=1)
add_textbox(slide, Inches(0.42), Inches(1.15), Inches(12.4), Inches(0.65),
            "Healthcare cost correlates with race/ethnicity, insurance type, and geography due to systemic inequities. "
            "A model trained on cost-based labels may replicate discrimination even with cost removed from features.",
            font_size=12, color=NAVY)

rows_data = [
    ["Dimension", "Fairness Metric", "Threshold", "Action if Breached"],
    ["Race/Ethnicity", "False Negative Parity\n(missed high-risk members)", "<5pp gap", "Recalibrate model; add fairness constraint"],
    ["Age", "Calibration (predicted vs actual risk)", "<10pp gap", "Age-specific model or feature adjustment"],
    ["Language", "Contact success rate", "<5pp gap", "Language-matched outreach; bilingual care manager"],
    ["Geography", "Model accuracy by zip code", "<10pp gap", "SDOH feature addition; local partner referrals"],
    ["Disability Status", "Selection rate", "Proportional ±5pp", "Manual review override; accessible contact channels"],
    ["Plan Type", "FCR by plan type", "<5pp gap", "Plan-specific model fine-tuning"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 4, Inches(0.3), Inches(1.95), Inches(12.7), Inches(3.75)).table
col_widths = [Inches(2.0), Inches(3.5), Inches(2.0), Inches(5.2)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = TABLE_ALT if r % 2 == 0 else WHITE
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

add_rect(slide, Inches(0.3), Inches(5.78), Inches(12.7), Inches(0.85), LIGHT_BLUE, border_color=TEAL, border_pt=1.5)
add_textbox(slide, Inches(0.42), Inches(5.83), Inches(12.4), Inches(0.75),
            "🤝  The AI model is a decision SUPPORT tool. It surfaces and prioritizes — it does not decide. "
            "Care managers review every recommendation. No member is denied outreach solely by algorithm. "
            "Documented in model card and clinical governance policy.",
            font_size=12, color=NAVY)

add_notes(slide,
    "This slide is for the data leader AND the clinical leader. The data leader wants AI rigor; the clinical leader wants human accountability. The answer is: both.\n\n"
    "Likely Q: What if you can't get race/ethnicity data?\n"
    "A: Use validated proxies (Bayesian Improved Surname Geocoding — BISG) and audit with available self-reported data. Disclose data limitations. Perfect is the enemy of good."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 21 — ALIGNING THREE LEADERS
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(21)
add_header_bar(slide, "One Shared Metric Set Ends the Three-Way Deadlock", TEAL)
add_footer_and_slide_number(slide, 21)

rows_data = [
    ["Leader", "What They Want", "What I Give Them", "Shared Metric"],
    ["Clinical Leader", "More staff; better member outcomes; evidence-based targeting",
     "Targeting engine reduces wasted calls → more time per engaged member; post-discharge fast lane; clinical override always available",
     "30-day readmission rate; successful contact rate"],
    ["Finance Leader", "ROI evidence; reduced program cost per outcome; justified headcount",
     "Holdout study proving causal impact; PMPM savings model; reachability filter reduces wasted calls",
     "PMPM impact; nurse hours per improved outcome; calls per engagement"],
    ["Data Leader", "AI adoption; model performance; bias-free predictions",
     "Four-factor model; fairness audit framework; ongoing monitoring dashboard; model card",
     "Reachability score accuracy; impactability model lift; equity gap metrics"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 4, Inches(0.3), Inches(1.1), Inches(12.7), Inches(3.5)).table
col_widths = [Inches(1.8), Inches(3.5), Inches(4.5), Inches(2.9)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

leader_bgs = [None, LIGHT_BLUE, LIGHT_PURPLE, TABLE_ALT2]
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_TEAL, align=PP_ALIGN.CENTER)
        else:
            bg = leader_bgs[r]
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# Bottom shared metric box
add_rect(slide, Inches(0.3), Inches(4.7), Inches(12.7), Inches(1.3), LIGHT_PURPLE, border_color=PURPLE, border_pt=2)
add_textbox(slide, Inches(0.42), Inches(4.75), Inches(3.5), Inches(0.4), "Shared Decision Metrics:", font_size=13, bold=True, color=PURPLE)
add_textbox(slide, Inches(0.42), Inches(5.1), Inches(12.4), Inches(0.85),
            "30-day readmission rate  +  Successful contact rate  +  PMPM impact  +  Equity gap\n"
            "All four leaders agree to report these at every steering committee.",
            font_size=12, color=NAVY)

add_notes(slide,
    "Stakeholder alignment is a product skill, not a soft skill. The three leaders have different incentives but one shared problem: the program is not working as well as it could.\n\n"
    "Likely Q: What if the clinical leader rejects the AI model?\n"
    "A: The AI surfaces candidates; the clinician decides. The model never overrides clinical judgment. That framing resolves most clinical resistance."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 22 — MEASUREMENT & ROLLOUT (CASE B)
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(22)
add_header_bar(slide, "A Holdout Design Proves Causal Impact — Not Correlation", PURPLE)
add_footer_and_slide_number(slide, 22)

rows_data = [
    ["Phase", "Duration", "Population", "Design", "Decision Gate"],
    ["Phase 0: Baseline", "8 weeks", "All eligible members, current process", "Measure current funnel, contact rates, outcomes", "Establish baseline"],
    ["Phase 1: Pilot", "12 weeks", "2 markets (1 Commercial, 1 Medicaid); 70% new targeting, 30% holdout", "Randomized controlled; measure all shared metrics", "FCR +10pp AND readmission -5pp → proceed"],
    ["Phase 2: Scale", "6 months", "All markets; phased by plan type", "Monitor equity gaps continuously", "No equity gap >5pp → full deployment"],
    ["Phase 3: Optimize", "Ongoing", "All members", "A/B test new factors, channel experiments", "Quarterly model refresh"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 5, Inches(0.3), Inches(1.1), Inches(12.7), Inches(3.0)).table
col_widths = [Inches(2.0), Inches(1.5), Inches(3.2), Inches(3.5), Inches(2.5)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

phase_bgs = [None, LIGHT_BLUE, LIGHT_PURPLE, TABLE_ALT2, LIGHT_BLUE]
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=TABLE_HEAD_PURPLE, align=PP_ALIGN.CENTER)
        else:
            bg = phase_bgs[r]
            style_cell(cell, text, font_size=10, bold=(c == 0), color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# ROI Model
add_rect(slide, Inches(0.3), Inches(4.2), Inches(12.7), Inches(1.5), LIGHT_PURPLE, border_color=PURPLE, border_pt=1.5)
add_textbox(slide, Inches(0.42), Inches(4.25), Inches(4), Inches(0.4), "ROI Model (Illustrative):", font_size=13, bold=True, color=PURPLE)
add_bullet_textbox(slide, Inches(0.42), Inches(4.65), Inches(12.4), Inches(0.9), [
    "• 30-day readmission cost: ~$15,000 per event  (Illustrative: Medicare benchmark)",
    "• If targeting improvement prevents 200 additional readmissions/year: $3M gross savings",
    "• Program cost increase (targeting engine + monitoring): ~$500K/year  →  Net annual benefit: ~$2.5M  (Illustrative)",
], font_size=11, color=NAVY)

add_notes(slide,
    "The holdout design is what finance needs to approve ongoing investment. Without a control group, we can never prove the program caused the improvement.\n\n"
    "Likely Q: Can you afford to hold out 30% of members from the new targeting?\n"
    "A: Yes, for the pilot phase. The holdout group receives the CURRENT standard of care — not worse care. And the evidence we generate enables us to serve ALL members better in Phase 2."
)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 23 — CLOSING: 90-DAY PLAN
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide(23)
add_header_bar(slide, "90 Days to a Decision-Ready Pilot — Here Is What I Need from Leadership", TEAL)
add_footer_and_slide_number(slide, 23)

# 90-day table
rows_data = [
    ["Days 1–30", "Days 31–60", "Days 61–90"],
    ["Stakeholder discovery: 15 interviews (call center, clinical, legal, data science, finance)",
     "PRD finalized and signed off by all 6 stakeholder groups",
     "Pilot launched (2 plan types, A/B design)"],
    ["Data audit: call volume, intent, PHI architecture, claims API",
     "Engineering sprint 1-4 complete: Coverage Q&A + deductible tracker",
     "Pilot monitoring dashboard live"],
    ["Vendor/platform selection: LLM + RAG infrastructure",
     "Safety classifier built and red-teamed",
     "First sprint review with all 6 stakeholder groups"],
    ["Legal review: PHI retention, state AI regs, CMS rules",
     "Multilingual MVP: Spanish",
     "Member feedback loop active"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 3, Inches(0.3), Inches(1.1), Inches(12.7), Inches(3.5)).table
col_widths = [Inches(4.2), Inches(4.2), Inches(4.3)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

col_colors = [PURPLE, TEAL, GREEN]
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=12, bold=True, color=WHITE, bg=col_colors[c], align=PP_ALIGN.CENTER)
        else:
            bg = LIGHT_PURPLE if c == 0 else (LIGHT_BLUE if c == 1 else TABLE_ALT2)
            style_cell(cell, text, font_size=11, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

# Decisions needed
add_rect(slide, Inches(0.3), Inches(4.7), Inches(12.7), Inches(1.75), LIGHT_PURPLE, border_color=PURPLE, border_pt=2)
add_textbox(slide, Inches(0.42), Inches(4.75), Inches(5), Inches(0.4), "Decisions I Need from Leadership:", font_size=13, bold=True, color=PURPLE)
add_bullet_textbox(slide, Inches(0.42), Inches(5.15), Inches(12.4), Inches(1.25), [
    "1. PHI data retention policy  (legal sign-off)    2. LLM platform decision (build vs buy vs partner)    3. Pilot population selection and holdout approval",
    "4. Budget authorization: Y1 investment ($3–4M illustrative)    5. Care management: AI model governance and clinical override policy",
], font_size=11, color=NAVY)

# Summary statement
add_rect(slide, Inches(0.3), Inches(6.52), Inches(12.7), Inches(0.6), PURPLE)
add_textbox(slide, Inches(0.42), Inches(6.57), Inches(12.4), Inches(0.5),
            "The opportunity is real, the risk is manageable, and the member impact is measurable. I am ready to build.",
            font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_footer_and_slide_number(slide, 23)
add_notes(slide,
    "Close with momentum, not summary. The 90-day plan shows I am ready to start Monday. The 'decisions I need' list respects leadership's role.\n\n"
    "Likely Q: What would cause you to recommend not building this?\n"
    "A: If legal review finds unresolvable compliance barriers, or if data audit shows we cannot ground answers accurately enough to meet safety thresholds — quality over speed, always."
)

# ══════════════════════════════════════════════════════════════════════════════
# APPENDIX 1 — ASSUMPTIONS REGISTER
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide("A1")
add_header_bar(slide, "Appendix A: Illustrative Assumptions Register", GRAY)
add_footer_and_slide_number(slide, "A1")

rows_data = [
    ["Assumption", "Value Used", "Source Logic", "Sensitivity"],
    ["Total annual member calls", "5M", "Mid-size US payer benchmark", "±2M changes NPV by ±$1.5M"],
    ["Cost per inbound call", "$10", "Industry benchmark (ACCS/CFI data)", "±$3 changes NPV by ±$1.2M"],
    ["Current digital self-service rate", "25%", "Payer digital benchmark", "Baseline for deflection calc"],
    ["Call deflection from AI (base)", "20%", "McKinsey/Accenture AI chatbot benchmarks", "Key sensitivity variable"],
    ["LLM platform cost", "$1.5M/yr", "Azure/GCP healthcare AI pricing estimates", "±$500K sensitivity"],
    ["Post-discharge readmission cost", "$15,000", "Medicare benchmark literature", "Used in Case B ROI only"],
    ["Readmission prevention (targeting)", "200 events/yr", "Illustrative; requires actual program data", "±100 events = ±$1.5M"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 4, Inches(0.3), Inches(1.1), Inches(12.7), Inches(5.0)).table
col_widths = [Inches(3.5), Inches(1.8), Inches(4.0), Inches(3.4)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=11, bold=True, color=WHITE, bg=GRAY, align=PP_ALIGN.CENTER)
        else:
            bg = TABLE_ALT if r % 2 == 0 else WHITE
            style_cell(cell, text, font_size=11, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

add_notes(slide, "Assumptions register. All numbers are illustrative and must be validated against Carelon actual data before any public or executive presentation.")

# ══════════════════════════════════════════════════════════════════════════════
# APPENDIX 2 — GLOSSARY
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide("A2")
add_header_bar(slide, "Appendix B: US Healthcare Glossary", GRAY)
add_footer_and_slide_number(slide, "A2")

terms = [
    ("Deductible", "Amount member pays out-of-pocket before insurance begins paying"),
    ("Out-of-Pocket Maximum (OOP Max)", "Maximum member pays in a year; insurance covers 100% after"),
    ("Coinsurance", "Percentage member pays after deductible (e.g., 20% of cost)"),
    ("Copay", "Fixed dollar amount for a service (e.g., $30 for PCP visit)"),
    ("EOB (Explanation of Benefits)", "Statement showing what insurance paid and what member owes"),
    ("Prior Authorization (PA)", "Insurance approval required before certain procedures or medications"),
    ("In-network", "Provider with contracted rates with the insurer"),
    ("HSA/HRA", "Health Savings/Reimbursement Account for pre-tax healthcare spending"),
    ("Medicare Advantage (MA)", "Private insurance alternative to traditional Medicare"),
    ("Medicaid", "State/federal insurance for low-income individuals"),
    ("HEDIS", "Quality measurement set for health plan performance"),
    ("PMPM", "Per Member Per Month — standard unit of healthcare cost/value"),
    ("SDOH", "Social Determinants of Health (housing, food, transportation, etc.)"),
    ("FCR", "First-Contact Resolution — member's question resolved without repeat contact"),
    ("RAG", "Retrieval-Augmented Generation — AI grounded in specific documents"),
]

# Split into 2 columns
mid = len(terms) // 2 + len(terms) % 2
col1 = terms[:mid]; col2 = terms[mid:]

rows = max(len(col1), len(col2)) + 1
tbl = slide.shapes.add_table(rows, 4, Inches(0.3), Inches(1.1), Inches(12.7), Inches(5.9)).table
tbl.columns[0].width = Inches(2.5)
tbl.columns[1].width = Inches(3.8)
tbl.columns[2].width = Inches(2.5)
tbl.columns[3].width = Inches(3.9)

headers = ["Term", "Definition", "Term", "Definition"]
for c, h in enumerate(headers):
    style_cell(tbl.cell(0, c), h, font_size=11, bold=True, color=WHITE, bg=GRAY, align=PP_ALIGN.CENTER)

for i in range(max(len(col1), len(col2))):
    bg = TABLE_ALT if i % 2 == 0 else WHITE
    if i < len(col1):
        style_cell(tbl.cell(i+1, 0), col1[i][0], font_size=10, bold=True, color=TEAL, bg=bg)
        style_cell(tbl.cell(i+1, 1), col1[i][1], font_size=10, color=NAVY, bg=bg)
    if i < len(col2):
        style_cell(tbl.cell(i+1, 2), col2[i][0], font_size=10, bold=True, color=TEAL, bg=bg)
        style_cell(tbl.cell(i+1, 3), col2[i][1], font_size=10, color=NAVY, bg=bg)

add_notes(slide, "Glossary for reference. Use these definitions when presenting to non-clinical stakeholders to ensure shared vocabulary.")

# ══════════════════════════════════════════════════════════════════════════════
# APPENDIX 3 — DETAILED RISK REGISTER
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide("A3")
add_header_bar(slide, "Appendix C: Full Risk Register", GRAY)
add_footer_and_slide_number(slide, "A3")

rows_data = [
    ["Risk", "Likelihood", "Impact", "Mitigation", "Owner", "Monitoring"],
    ["AI hallucinates a benefit that doesn't exist", "Medium", "High", "RAG constraint + confidence threshold <0.7 + monthly citation audit", "Data Science / PM", "Monthly accuracy audit"],
    ["Safety classifier misses a distress signal", "Low", "Critical", "Layered: keyword matching + LLM classifier + human review queue", "Clinical Quality / DS", "Monthly red-team testing"],
    ["State AI regulation for insurance (NAIC)", "Medium", "High", "Monitor NAIC model bulletin; legal sign-off before launch in each state", "Legal", "Quarterly regulatory scan"],
    ["Agent resistance to AI", "Medium", "Medium", "Change mgmt: AI as assistant; demonstrate handle time reduction", "PM / HR", "Agent satisfaction survey Q1"],
    ["LLM bias in underrepresented populations", "Medium", "High", "Pre-launch equity audit; FCR monitoring by segment; fairness constraints", "Data Science", "Monthly equity dashboard"],
    ["Low member adoption (especially seniors)", "Medium", "Medium", "IVR integration V2; 'Talk to person' always prominent; CSAT by age band", "PM / UX", "Adoption by age band weekly"],
    ["LLM vendor lock-in / cost escalation", "Low", "Medium", "Abstraction layer enabling model swap; monitor cost/query monthly", "Engineering", "Monthly cost monitoring"],
    ["PHI data breach via LLM API", "Low", "Critical", "PHI never sent to external LLM; on-prem/VPC deployment; HIPAA BAA", "Security / Legal", "Annual security audit + pen test"],
    ["Regulatory change: CMS AI rules for MA", "Medium", "High", "Monitor CMS guidance; legal review before each MA feature", "Legal / Compliance", "Quarterly review"],
    ["Clinical leader rejects AI recommendations", "Low", "Medium", "AI as decision support, not decision maker; clinician override always available", "PM / Clinical Quality", "Override rate monitoring"],
]
n_rows = len(rows_data)
tbl = slide.shapes.add_table(n_rows, 6, Inches(0.3), Inches(1.1), Inches(12.7), Inches(5.9)).table
col_widths = [Inches(2.5), Inches(1.2), Inches(1.1), Inches(3.5), Inches(2.0), Inches(2.4)]
for i, w in enumerate(col_widths):
    tbl.columns[i].width = w

impact_bgs = {"Critical": RGBColor(0xFF, 0xDD, 0xDD), "High": RGBColor(0xFF, 0xF0, 0xDD), "Medium": LIGHT_BLUE}
for r, row_data in enumerate(rows_data):
    for c, text in enumerate(row_data):
        cell = tbl.cell(r, c)
        if r == 0:
            style_cell(cell, text, font_size=10, bold=True, color=WHITE, bg=GRAY, align=PP_ALIGN.CENTER)
        else:
            if c == 2:
                bg = impact_bgs.get(text, WHITE)
            else:
                bg = TABLE_ALT if r % 2 == 0 else WHITE
            style_cell(cell, text, font_size=9, color=NAVY, bg=bg, align=PP_ALIGN.LEFT)

add_notes(slide, "Full risk register. Review this with legal, clinical, and data science before pilot launch. Update likelihood/impact quarterly.")

# ══════════════════════════════════════════════════════════════════════════════
# APPENDIX 4 — ADDITIONAL USER STORIES
# ══════════════════════════════════════════════════════════════════════════════
slide = new_slide("A4")
add_header_bar(slide, "Appendix D: Additional User Stories", GRAY)
add_footer_and_slide_number(slide, "A4")

stories = [
    ("EOB Explainer", "Member",
     "Given member asks 'Why do I still owe money after insurance paid?', When AI retrieves EOB and claims data, Then response explains in plain language: what provider charged, what insurance paid, member responsibility amount, and why (deductible/coinsurance), with a link to the full EOB document."),
    ("Cost Estimator (V2)", "HDHP Member",
     "Given member asks 'How much will an MRI cost me?', When AI queries cost estimator with provider + plan + deductible remaining, Then response shows estimated member cost range, deductible remaining, and disclaimer: 'Actual cost depends on specific billing codes — confirm with provider.'"),
    ("Medicare Advantage Member", "Eleanor, MA Member",
     "Given MA member asks 'What's covered under my plan for preventive visits?', When AI retrieves MA Summary of Benefits document, Then response lists covered preventive services, copay amounts, and limits in plain language at ≤6th grade reading level, with option to 'Explain simpler'."),
    ("Honest Low-Confidence Handoff", "Any Member",
     "Given member asks a complex question (e.g., coverage for experimental treatment), When AI confidence score <0.7 OR no document chunk found, Then response = 'I want to make sure you get an accurate answer on this — it involves some complexity. Let me connect you to a specialist.' + warm handoff with full context summary."),
    ("Appeals Initiation", "David, Chronic Condition",
     "Given member says 'My prior auth was denied and I want to appeal', When AI detects appeal intent, Then response = explanation of appeal rights + links to appeal form + offer to initiate appeal workflow + mandatory escalation to human agent (CMS requirement). No AI-only handling of appeals."),
]

y_start = Inches(1.1); story_h = Inches(1.14)
for i, (title, persona, story) in enumerate(stories):
    y = y_start + i * (story_h + Inches(0.03))
    color = TEAL if i % 2 == 0 else PURPLE
    add_rect(slide, Inches(0.3), y, Inches(1.5), story_h, color)
    add_textbox(slide, Inches(0.32), y + Inches(0.1), Inches(1.46), story_h - Inches(0.2),
                f"Story\n{i+1}", font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(1.9), y, Inches(11.1), story_h, LIGHT_BLUE if i % 2 == 0 else LIGHT_PURPLE,
             border_color=color, border_pt=0.5)
    add_textbox(slide, Inches(2.0), y + Inches(0.04), Inches(10.9), Inches(0.3),
                f"{title}  |  Persona: {persona}", font_size=11, bold=True, color=color)
    add_textbox(slide, Inches(2.0), y + Inches(0.34), Inches(10.9), story_h - Inches(0.4),
                story, font_size=10, color=NAVY)

add_notes(slide, "Additional user stories for reference. Use these to populate the backlog after PRD sign-off. Each story should be refined with engineering before sprint planning.")

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
output_path = r"D:\Preksha_carelon\Carelon_CaseStudy_Presentation.pptx"
prs.save(output_path)
print(f"SUCCESS: Saved to {output_path}")
print(f"Total slides: {len(prs.slides)}")
