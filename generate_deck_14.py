import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Initialize Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Color Palette Definitions
C_NAVY = RGBColor(26, 26, 46)        # #1A1A2E - Deep Navy Primary Text
C_TEAL = RGBColor(13, 124, 143)      # #0D7C8F - Calm Healthcare Teal
C_PURPLE = RGBColor(91, 45, 142)     # #5B2D8E - Carelon Brand Purple Accent
C_BG_LIGHT = RGBColor(248, 249, 250) # #F8F9FA - Soft Canvas
C_CARD_BG = RGBColor(255, 255, 255)  # White Card
C_BORDER = RGBColor(226, 230, 239)   # Border Gray
C_MUTED = RGBColor(106, 115, 131)    # Muted Gray Body/Labels
C_TEAL_LIGHT = RGBColor(235, 246, 248) # Soft Teal Tint
C_PURPLE_LIGHT = RGBColor(246, 240, 255) # Soft Purple Tint
C_GREEN = RGBColor(46, 204, 113)     # #2ECC71 - Success Green
C_RED = RGBColor(231, 76, 60)        # #E74C3C - Red / Alert
C_AMBER = RGBColor(243, 156, 18)     # #F39C12 - Warning

def apply_background(slide):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_BG_LIGHT
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, tag_text, headline_text, max_w=12.0):
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.35))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.name = "Calibri"
    p_tag.font.color.rgb = C_TEAL

    hl_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(max_w), Inches(0.85))
    tf_hl = hl_box.text_frame
    tf_hl.word_wrap = True
    tf_hl.margin_left = tf_hl.margin_top = tf_hl.margin_right = tf_hl.margin_bottom = 0
    p_hl = tf_hl.paragraphs[0]
    p_hl.text = headline_text
    p_hl.font.size = Pt(20)
    p_hl.font.bold = True
    p_hl.font.name = "Calibri"
    p_hl.font.color.rgb = C_NAVY

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.62), Inches(1.5), Inches(0.04))
    line.fill.solid()
    line.fill.fore_color.rgb = C_PURPLE
    line.line.fill.background()

def add_footer(slide, current_idx, total_slides=18):
    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
    tf = footer_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"Carelon Global Solutions | Sr. BA II Submission | Member AI Assistant  •  Confidential (Illustrative Assumptions Labeled)                                          Slide {current_idx:02d} of {total_slides:02d}"
    p.font.size = Pt(9)
    p.font.color.rgb = C_MUTED
    p.font.name = "Calibri"

def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1)
    return card

def add_speaker_notes(slide, script_text, question_text, answer_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = f"TALK TRACK (45s):\n{script_text}\n\nLIKELY INTERVIEWER QUESTION:\n{question_text}\n\nSTRONG PM ANSWER:\n{answer_text}"

# ==============================================================================
# SLIDE 1: TITLE & THESIS
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
apply_background(s1)

# Large Hero Card
add_card(s1, 0.8, 0.8, 11.733, 5.9, bg_color=C_CARD_BG, border_color=C_BORDER)

# Decorative left accent bar on hero card
accent_bar = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(0.2), Inches(5.9))
accent_bar.fill.solid()
accent_bar.fill.fore_color.rgb = C_PURPLE
accent_bar.line.fill.background()

# Title text content
tb_title = s1.shapes.add_textbox(Inches(1.4), Inches(1.2), Inches(10.5), Inches(3.2))
tf1 = tb_title.text_frame
tf1.word_wrap = True

p0 = tf1.paragraphs[0]
p0.text = "CARELON GLOBAL SOLUTIONS | PRODUCT STRATEGY CASE SUBMISSION"
p0.font.size = Pt(11)
p0.font.bold = True
p0.font.color.rgb = C_TEAL

p1 = tf1.add_paragraph()
p1.text = "An AI Assistant for Health Insurance Members"
p1.font.size = Pt(30)
p1.font.bold = True
p1.font.color.rgb = C_NAVY
p1.space_before = Pt(8)

p2 = tf1.add_paragraph()
p2.text = "Sr. Business Analyst II (Product Owner / US Healthcare Agile) Interview Submission"
p2.font.size = Pt(14)
p2.font.color.rgb = C_MUTED
p2.space_before = Pt(6)

# One-Sentence Thesis Card
thesis_card = add_card(s1, 1.4, 3.4, 10.5, 1.4, bg_color=C_PURPLE_LIGHT, border_color=C_PURPLE)
tb_thesis = s1.shapes.add_textbox(Inches(1.6), Inches(3.5), Inches(10.1), Inches(1.2))
tft = tb_thesis.text_frame
tft.word_wrap = True
pt_label = tft.paragraphs[0]
pt_label.text = "STRATEGIC THESIS:"
pt_label.font.size = Pt(10)
pt_label.font.bold = True
pt_label.font.color.rgb = C_PURPLE

pt_body = tft.add_paragraph()
pt_body.text = "By grounding conversational AI strictly in verified member plan documents and accumulator APIs, we resolve 65% of repetitive coverage inquiries on first contact—deflecting $4M+ in live-call costs, improving member literacy, and ensuring strict health equity across commercial, Medicare, and Medicaid populations."
pt_body.font.size = Pt(13)
pt_body.font.color.rgb = C_NAVY
pt_body.font.bold = True
pt_body.space_before = Pt(4)

# Metadata bar
tb_meta = s1.shapes.add_textbox(Inches(1.4), Inches(5.3), Inches(10.5), Inches(0.8))
tf_meta = tb_meta.text_frame
p_meta = tf_meta.paragraphs[0]
p_meta.text = "Candidate: Product Management Aspirant  •  Elevance Health Subsidiary Context  •  November 2026  •  Methodology: Scaled Agile (SAFe) & Responsible AI"
p_meta.font.size = Pt(11)
p_meta.font.color.rgb = C_MUTED

add_footer(s1, 1)
add_speaker_notes(s1, 
    "Good morning. Today I will present an end-to-end product strategy for an AI Assistant designed specifically for US health plan members. My core thesis is that member confusion is not an intelligence deficit—it is a structural health insurance complexity problem. By using retrieval-augmented generation anchored strictly in individual benefit documents and claims accumulators, we can achieve 65% first-contact resolution while safeguarding clinical safety and cross-population equity.",
    "Why choose First-Contact Resolution (FCR) as your north-star metric over simple containment or deflection?",
    "Containment can be a vanity metric—a member who gets trapped in a chatbot and rage-quits is technically 'contained' but highly dissatisfied. First-Contact Resolution measures both effectiveness and satisfaction: did the member obtain the correct, verified answer to their need without having to repeat their journey or call back within 7 days?")

# ==============================================================================
# SLIDE 2: PROBLEM FRAMING & ESCALATION LOOP
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
apply_background(s2)
add_header(s2, "Problem & Operational Impact", "Members Are Trapped in a 3-Channel Loop That Erodes Trust and Costs ~$50M Annually")

# Left Column: The 3-Channel Escalation Loop Visual (Cards)
add_card(s2, 0.8, 1.85, 5.7, 4.9)
tb_loop_hdr = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(0.4))
tb_loop_hdr.text_frame.paragraphs[0].text = "The Friction-Heavy Multi-Channel Loop Today"
tb_loop_hdr.text_frame.paragraphs[0].font.size = Pt(13)
tb_loop_hdr.text_frame.paragraphs[0].font.bold = True
tb_loop_hdr.text_frame.paragraphs[0].font.color.rgb = C_NAVY

steps = [
    ("1. Digital Search", "Member logs into portal to check coverage or a recent claim; encounters 90-page dense PDF Summary of Benefits."),
    ("2. Web FAQ / IVR Dead End", "Keyword search fails on plan-specific nuances (e.g. out-of-network coinsurance vs copay)."),
    ("3. Live Call Escalation", "Member calls customer service. Average wait time 6-8 mins. Representative looks up exact same SBC."),
    ("4. Channel Switching Fatigue", "Member repeats their story 2-3 times across departments, leading to frustration and repeat calls next month.")
]

for i, (stitle, sdesc) in enumerate(steps):
    step_bg = C_TEAL_LIGHT if i % 2 == 0 else C_PURPLE_LIGHT
    card_step = add_card(s2, 1.0, 2.45 + (i * 1.0), 5.3, 0.9, bg_color=step_bg, border_color=C_BORDER)
    tb_s = s2.shapes.add_textbox(Inches(1.1), Inches(2.5 + (i * 1.0)), Inches(5.1), Inches(0.8))
    tfs = tb_s.text_frame
    tfs.word_wrap = True
    ps1 = tfs.paragraphs[0]
    ps1.text = stitle
    ps1.font.size = Pt(11)
    ps1.font.bold = True
    ps1.font.color.rgb = C_TEAL if i % 2 == 0 else C_PURPLE
    ps2 = tfs.add_paragraph()
    ps2.text = sdesc
    ps2.font.size = Pt(10)
    ps2.font.color.rgb = C_NAVY
    ps2.space_before = Pt(2)

# Right Column: Operational Sizing & Who is Hurting
add_card(s2, 6.7, 1.85, 5.8, 4.9)
tb_stake_hdr = s2.shapes.add_textbox(Inches(6.9), Inches(2.0), Inches(5.4), Inches(0.4))
tb_stake_hdr.text_frame.paragraphs[0].text = "Who is Hurting & The True Cost-to-Serve"
tb_stake_hdr.text_frame.paragraphs[0].font.size = Pt(13)
tb_stake_hdr.text_frame.paragraphs[0].font.bold = True
tb_stake_hdr.text_frame.paragraphs[0].font.color.rgb = C_NAVY

# 3 Stakeholder pain boxes
pains = [
    ("Members (Millions Covered)", "80% struggle to define 'deductible', 'coinsurance', or 'out-of-pocket max'. High anxiety during medical episodes when billing surprises occur."),
    ("Customer Service Agents", "Burdened by repetitive Tier-1 coverage questions (~30% of call volume). Results in burnout, 35%+ annual agent turnover, and high onboarding costs."),
    ("Health Plan & Operations", "~5M calls/yr across book of business. At ~$10 per inbound call, basic coverage navigation costs ~$15M-$20M alone (Illustrative assumption).")
]
for j, (p_title, p_desc) in enumerate(pains):
    add_card(s2, 6.9, 2.45 + (j * 1.35), 5.4, 1.2, bg_color=C_CARD_BG, border_color=C_BORDER)
    tb_p = s2.shapes.add_textbox(Inches(7.05), Inches(2.5 + (j * 1.35)), Inches(5.1), Inches(1.1))
    tfp = tb_p.text_frame
    tfp.word_wrap = True
    pp1 = tfp.paragraphs[0]
    pp1.text = f"• {p_title}"
    pp1.font.size = Pt(11)
    pp1.font.bold = True
    pp1.font.color.rgb = C_NAVY
    pp2 = tfp.add_paragraph()
    pp2.text = p_desc
    pp2.font.size = Pt(10)
    pp2.font.color.rgb = C_MUTED
    pp2.space_before = Pt(3)

add_footer(s2, 2)
add_speaker_notes(s2,
    "On this slide we examine the core friction: members are trapped in an escalation loop. Because health plan documents are written for regulatory compliance rather than readability, 80% of members experience health insurance literacy barriers. When self-service fails, they call. For a payer with 5 million inbound calls a year, basic coverage and claims status Q&A consumes 30% of call center capacity at roughly $10 per call. This wastes millions while degrading member CSAT.",
    "What empirical evidence suggests members will adopt an AI assistant instead of calling?",
    "Consumer behavior across banking and telecom shows that 68% of users prefer digital self-service if—and only if—it provides instant, personalized, and accurate answers. Our challenge isn't channel adoption; it's trust in accuracy.")

# ==============================================================================
# SLIDE 3: PERSONAS & JOBS-TO-BE-DONE
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
apply_background(s3)
add_header(s3, "Personas & Segmentation", "Five Personas Reveal Vastly Different Literacy, Channel Needs, and Emotional Stakes")

# Table of Personas
table_shape = s3.shapes.add_table(6, 5, Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.9))
t3 = table_shape.table
t3.columns[0].width = Inches(2.3)
t3.columns[1].width = Inches(1.6)
t3.columns[2].width = Inches(3.2)
t3.columns[3].width = Inches(2.4)
t3.columns[4].width = Inches(2.233)

headers3 = ["Persona & Profile", "Health Literacy", "Top Job-to-be-Done (JTBD)", "Channel Preference", "Emotional State"]
for col_idx, htext in enumerate(headers3):
    cell = t3.cell(0, col_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

p_data = [
    ("Alex, 24\nNew-to-Insurance", "Very Low\n(6th grade or below)", "'Explain what a deductible is in plain English and tell me if my urgent care visit is covered before I go.'", "Mobile App, Push / SMS\nAsync digital chats", "Anxious, Intimidated,\nCost-conscious"),
    ("Maria, 38\nFamily Caregiver", "Moderate\n(Manages 4 lives)", "'Track my family deductible progress, confirm pediatric ENT network status, and explain why a claim was denied.'", "Web Portal (laptop) &\nMobile App; saves PDFs", "Time-starved, Stressed,\nProtective"),
    ("David, 52\nChronic Condition", "High Claims Literacy\n(Low Clinical AI Trust)", "'Check if my branded insulin is Tier 2, verify prior-authorization renewal status, and monitor my OOP maximum.'", "Web Portal + Live Phone\nDemands audit trails", "Frustrated, Cautious,\nSkeptical of delays"),
    ("Eleanor, 68\nMedicare Advantage", "Low-to-Moderate\n(Fixed income)", "'Does my plan cover cataract surgery transportation and hearing aids? Tell me my exact copay.'", "Phone-first; prefers warm\nvoice agent transfer", "Vulnerable, Trust-sensitive,\nValues human touch"),
    ("Priya, 34\nHigh-Deductible (HDHP)", "High Financial Literacy\n(Analytical)", "'How much of my $4,000 deductible remains? If I schedule knee MRI in Nov vs Jan, what is my net out-of-pocket?'", "Web Portal, HSA integration,\nCost calculators", "Analytical, Proactive,\nPrice-sensitive")
]

for row_idx, row_vals in enumerate(p_data, start=1):
    for c_idx, val in enumerate(row_vals):
        cell = t3.cell(row_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if row_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(10)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
            p.font.color.rgb = C_PURPLE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s3, 3)
add_speaker_notes(s3,
    "Our user base is fundamentally heterogeneous. Designing for Alex, a 24-year-old on his first employer plan who doesn't know what coinsurance means, requires an entirely different UX than designing for Eleanor, a Medicare Advantage member on a fixed income, or David, who relies on complex prior-authorization for insulin. This is why our assistant includes adaptive reading levels—defaulting to 6th-grade Flesch-Kincaid—and provides seamless phone handoffs for senior members.",
    "How does the assistant dynamically detect and adapt to a user's literacy level?",
    "We use implicit signals—such as query syntactic complexity and frequent clicks on 'Explain Simpler'—combined with explicit profile toggles, ensuring we never patronize advanced users while never leaving low-literacy members behind.")

# ==============================================================================
# SLIDE 4: MARKET & COMPETITOR SCAN
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
apply_background(s4)
add_header(s4, "Market & Competitive Landscape", "The Payer Landscape Leaves an Unfilled White Space: High Grounding with Broad Self-Service")

# 2x2 Matrix Visual
add_card(s4, 0.8, 1.85, 6.8, 4.9)
tb_map = s4.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(6.4), Inches(0.35))
tb_map.text_frame.paragraphs[0].text = "Competitive Positioning Map: Trust & Grounding vs Capability"
tb_map.text_frame.paragraphs[0].font.size = Pt(12)
tb_map.text_frame.paragraphs[0].font.bold = True
tb_map.text_frame.paragraphs[0].font.color.rgb = C_NAVY

# 4 Quadrants
q1 = add_card(s4, 1.0, 2.45, 3.1, 1.9, bg_color=C_CARD_BG, border_color=C_BORDER) # Top Left: High Trust, Narrow
q2 = add_card(s4, 4.3, 2.45, 3.1, 1.9, bg_color=C_PURPLE_LIGHT, border_color=C_PURPLE) # Top Right: TARGET
q3 = add_card(s4, 1.0, 4.55, 3.1, 1.9, bg_color=C_BG_LIGHT, border_color=C_BORDER) # Bottom Left: Low Trust, Narrow
q4 = add_card(s4, 4.3, 4.55, 3.1, 1.9, bg_color=C_CARD_BG, border_color=C_BORDER) # Bottom Right: Low Trust, Broad

# Labels for quadrants
def fill_q(slide, left, top, title, players, color):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(2.9), Inches(1.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = color
    p2 = tf.add_paragraph()
    p2.text = players
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = C_NAVY
    p2.space_before = Pt(3)

fill_q(s4, 1.1, 2.5, "High Trust / Narrow Scope", "Traditional IVR & Rules Bots\n• Payer Phone Menus (Nuance)\n• Static Portal FAQ lookups\n• High accuracy, zero flexibility", C_TEAL)
fill_q(s4, 4.4, 2.5, "★ TARGET: Carelon Assistant", "High Trust / Broad Scope\n• Grounded RAG on SBC + Claims\n• Hallucination-free member navigation\n• Multilingual & Plain-language", C_PURPLE)
fill_q(s4, 1.1, 4.6, "Low Trust / Narrow Scope", "Basic Payer Chatbots (V1)\n• Generic decision trees\n• High drop-off / 'I don't understand'\n• Forces immediate call escalation", C_MUTED)
fill_q(s4, 4.4, 4.6, "Low Trust / Broad Scope", "Public LLMs & Symptom Checkers\n• ChatGPT, WebMD, Ada\n• No plan-specific benefit grounding\n• High hallucination liability on PHI", C_RED)

# Right Column: Insights & Market Verification
add_card(s4, 7.8, 1.85, 4.733, 4.9)
tb_comp = s4.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
tfc = tb_comp.text_frame
tfc.word_wrap = True

pc1 = tfc.paragraphs[0]
pc1.text = "Key Competitive Insights"
pc1.font.size = Pt(13)
pc1.font.bold = True
pc1.font.color.rgb = C_NAVY

c_points = [
    ("Payer Peers (Cigna, UHC, Humana):", "Most deploy rule-based chatbots handling basic ID cards and claims lookup; few offer contextual, document-grounded RAG Q&A."),
    ("Big Tech Infrastructure:", "Microsoft Azure Health Bot and AWS HealthLake offer strong compliance rails but require heavy payer-specific integration."),
    ("The Moat for Carelon / Elevance:", "Our competitive advantage lies in integrating our proprietary claims accumulator feeds, provider directories, and clinical management engine into one unified conversational layer."),
    ("⚠️ Market Fact Check Warning:", "Competitor feature sets evolve monthly; verify current commercial status before client/executive presentation.")
]

for cp_title, cp_desc in c_points:
    p_t = tfc.add_paragraph()
    p_t.text = cp_title
    p_t.font.size = Pt(10.5)
    p_t.font.bold = True
    p_t.font.color.rgb = C_TEAL if "Moat" not in cp_title else C_PURPLE
    p_t.space_before = Pt(8)
    p_d = tfc.add_paragraph()
    p_d.text = cp_desc
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = C_NAVY
    p_d.space_before = Pt(2)

add_footer(s4, 4)
add_speaker_notes(s4,
    "When we benchmark the market, we see four distinct quadrants. In the bottom-right, public LLMs like ChatGPT have broad conversational capabilities but cannot access individual claims and pose catastrophic hallucination risks. In the top-left, traditional IVRs and rules-based bots are safe but rigid. Our strategic target is the top-right white space: an enterprise-grade, retrieval-grounded assistant that has broad conversational flexibility while strictly bounded to verified health plan documents.",
    "What prevents a big tech vendor like Microsoft or Google from disintermediating this product?",
    "Big tech provides the commodity LLM infrastructure, but they do not possess the payer-specific data orchestration layer—our real-time claims accumulators, benefit hierarchies, and clinical escalation protocols. The moat is proprietary integration, not the raw LLM model.")

# ==============================================================================
# SLIDE 5: VISION, PRINCIPLES & METRICS
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
apply_background(s5)
add_header(s5, "Vision, Principles & Governance", "A North-Star of First-Contact Resolution Governed by Five Non-Negotiable Principles")

# Vision Callout
add_card(s5, 0.8, 1.85, 11.733, 0.9, bg_color=C_TEAL_LIGHT, border_color=C_TEAL)
tb_vis = s5.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(0.8))
tfv = tb_vis.text_frame
tfv.word_wrap = True
pv1 = tfv.paragraphs[0]
pv1.text = "PRODUCT VISION:"
pv1.font.size = Pt(10)
pv1.font.bold = True
pv1.font.color.rgb = C_TEAL
pv2 = tfv.add_paragraph()
pv2.text = "'Empower every Carelon member to understand and navigate their healthcare coverage instantly, with total confidence in accuracy, plain language, and zero channel-switching friction.'"
pv2.font.size = Pt(12.5)
pv2.font.bold = True
pv2.font.color.rgb = C_NAVY

# Left: 5 Product Principles
add_card(s5, 0.8, 2.9, 7.0, 3.9)
tb_pr = s5.shapes.add_textbox(Inches(1.0), Inches(3.05), Inches(6.6), Inches(3.6))
tfpr = tb_pr.text_frame
tfpr.word_wrap = True
ppr_h = tfpr.paragraphs[0]
ppr_h.text = "Five Core Product Principles"
ppr_h.font.size = Pt(13)
ppr_h.font.bold = True
ppr_h.font.color.rgb = C_NAVY

principles = [
    ("1. Trust Over Completeness", "Never fabricate. If confidence is <0.7, acknowledge uncertainty and warm-transfer."),
    ("2. 6th-Grade Plain Language", "Translate opaque jargon (e.g. 'allowed amount') into actionable explanations."),
    ("3. Zero Dead Ends", "Every interaction concludes with a verified answer, a logical next step, or a human handoff."),
    ("4. Equitable by Architecture", "Equal accuracy and responsiveness across Commercial, Medicare, and Medicaid."),
    ("5. Minimum Necessary PHI", "Access only data required to answer the specific intent; never expose full medical records.")
]
for p_num, p_body in principles:
    p_p1 = tfpr.add_paragraph()
    p_p1.text = f"{p_num}: {p_body}"
    p_p1.font.size = Pt(10)
    p_p1.font.color.rgb = C_NAVY
    p_p1.space_before = Pt(6)

# Right: North Star & Guardrail Metric Cards
add_card(s5, 8.0, 2.9, 4.533, 3.9)
tb_m = s5.shapes.add_textbox(Inches(8.2), Inches(3.05), Inches(4.1), Inches(3.6))
tfm = tb_m.text_frame
tfm.word_wrap = True

pm_h = tfm.paragraphs[0]
pm_h.text = "North-Star & Guardrail Metrics"
pm_h.font.size = Pt(13)
pm_h.font.bold = True
pm_h.font.color.rgb = C_NAVY

m_cards = [
    ("★ NORTH-STAR METRIC", "First-Contact Resolution Rate (FCR)", "Target: 65% within 12 months (Illustrative)", C_PURPLE),
    ("GUARDRAIL 1: SAFETY", "Clinical Escalation Recall: >99.0%", "Zero missed chest pain/distress signals", C_RED),
    ("GUARDRAIL 2: ACCURACY", "Hallucination Rate: <1.0%", "Audited against verified plan source documents", C_TEAL),
    ("GUARDRAIL 3: EQUITY", "Cross-Segment FCR Gap: <5.0 pp", "Parity between Commercial and Medicaid lines", C_AMBER)
]
for m_label, m_title, m_desc, m_color in m_cards:
    p_m1 = tfm.add_paragraph()
    p_m1.text = m_label
    p_m1.font.size = Pt(9)
    p_m1.font.bold = True
    p_m1.font.color.rgb = m_color
    p_m1.space_before = Pt(6)
    p_m2 = tfm.add_paragraph()
    p_m2.text = f"{m_title} ({m_desc})"
    p_m2.font.size = Pt(9.5)
    p_m2.font.color.rgb = C_NAVY

add_footer(s5, 5)
add_speaker_notes(s5,
    "Our product principles serve as engineering trade-off filters. For example, Principle 1 means that if the LLM cannot locate an explicit benefit clause, it is forbidden from guessing; it must admit uncertainty. Our North-Star Metric is First-Contact Resolution, targeted at 65%. But notice the guardrail metrics: we balance FCR with an hallucination rate under 1%, safety recall above 99%, and a cross-segment equity gap under 5 percentage points.",
    "If a member rates an answer 1 star, how does that factor into FCR?",
    "FCR is a composite metric: it requires both operational closure (no repeat contact within 7 days) and member verification (thumbs up or survey confirmation). A 1-star review automatically disqualifies the session from FCR and routes the transcript to the human-in-the-loop review queue.")

# ==============================================================================
# SLIDE 6: OPPORTUNITY SIZING
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
apply_background(s6)
add_header(s6, "Opportunity Sizing (Illustrative)", "Deflecting 20% of Automatable Inquiries Yields ~$4.0M in Gross Annual Call Savings")

# Left Column: Waterfall Math Table
add_card(s6, 0.8, 1.85, 7.5, 4.9)
tb_tab_h = s6.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(7.1), Inches(0.4))
tb_tab_h.text_frame.paragraphs[0].text = "Step-by-Step Opportunity Sizing Waterfall (All Numbers Illustrative)"
tb_tab_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_tab_h.text_frame.paragraphs[0].font.bold = True
tb_tab_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

t6_shape = s6.shapes.add_table(8, 3, Inches(1.0), Inches(2.45), Inches(7.1), Inches(4.0))
t6 = t6_shape.table
t6.columns[0].width = Inches(3.6)
t6.columns[1].width = Inches(1.5)
t6.columns[2].width = Inches(2.0)

t6_data = [
    ("Metric / Driver", "Value", "Baseline Logic & Assumptions"),
    ("Total Annual Inbound Member Calls", "5,000,000", "Mid-to-large payer book of business assumption"),
    ("Share of Automatable Contact Reasons", "40.0% (2.0M)", "Coverage Q&A, deductible, claims status, ID cards"),
    ("Current Digital Containment Rate", "25.0%", "Current web FAQ / basic chatbot performance"),
    ("Target Digital Deflection Rate (Y1)", "20.0%", "Realistic net deflection from conversational AI assistant"),
    ("Annual Inbound Calls Deflected", "400,000", "2.0M automatable inquiries × 20% deflection"),
    ("Average Cost per Live Agent Call", "$10.00", "Industry loaded benchmark: $8.00–$12.00 / call"),
    ("Gross Annual Operational Savings", "$4,000,000", "400,000 calls deflected × $10.00 / call")
]

for r_idx, r_vals in enumerate(t6_data):
    for c_idx, val in enumerate(r_vals):
        cell = t6.cell(r_idx, c_idx)
        cell.fill.solid()
        if r_idx == 0:
            cell.fill.fore_color.rgb = C_TEAL
        elif r_idx == 7:
            cell.fill.fore_color.rgb = C_PURPLE_LIGHT
        else:
            cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(9.5)
        p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else (C_PURPLE if r_idx == 7 else C_NAVY)
        if r_idx == 0 or r_idx == 7 or c_idx == 1:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right Column: Sensitivity Cards
add_card(s6, 8.5, 1.85, 4.033, 4.9)
tb_scen = s6.shapes.add_textbox(Inches(8.7), Inches(2.0), Inches(3.6), Inches(4.5))
tfsc = tb_scen.text_frame
tfsc.word_wrap = True

psh = tfsc.paragraphs[0]
psh.text = "Financial Sensitivity Analysis"
psh.font.size = Pt(13)
psh.font.bold = True
psh.font.color.rgb = C_NAVY

scenarios = [
    ("Conservative Scenario (10% Deflection)", "• 200,000 calls deflected\n• $2.0M Gross Annual Savings\n• 3-Year NPV: ~$2.1M (Break-even Mo 22)", C_MUTED),
    ("Base Case Scenario (20% Deflection)", "• 400,000 calls deflected\n• $4.0M Gross Annual Savings\n• 3-Year NPV: ~$5.2M (Break-even Mo 14)", C_TEAL),
    ("Aggressive Scenario (30% Deflection)", "• 600,000 calls deflected\n• $6.0M Gross Annual Savings\n• 3-Year NPV: ~$8.9M (Break-even Mo 9)", C_PURPLE)
]

for s_title, s_body, s_col in scenarios:
    p_st = tfsc.add_paragraph()
    p_st.text = s_title
    p_st.font.size = Pt(11)
    p_st.font.bold = True
    p_st.font.color.rgb = s_col
    p_st.space_before = Pt(8)
    p_sb = tfsc.add_paragraph()
    p_sb.text = s_body
    p_sb.font.size = Pt(9.5)
    p_sb.font.color.rgb = C_NAVY
    p_sb.space_before = Pt(2)

add_footer(s6, 6)
add_speaker_notes(s6,
    "Here is our illustrative opportunity sizing model. Assuming a 5 million annual call volume and an automatable share of 40%, our base-case target of 20% call deflection translates to 400,000 fewer inbound calls. At an industry-standard $10 cost per call, that generates $4.0 million in gross annual operational savings. Even in our conservative scenario of 10% deflection, the project generates $2.0 million annually and breaks even in month 22.",
    "Does deflecting calls lead to direct call-center staff layoffs?",
    "No, and that is a key change-management message for call-center leadership. Deflection eliminates low-value, transactional calls, freeing clinical and high-tier agents to handle complex appeals, high-risk care management, and members in clinical distress, reducing burnout.")

# ==============================================================================
# SLIDE 7: OPTIONS & PRIORITIZATION (RICE)
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
apply_background(s7)
add_header(s7, "Prioritization & Scope", "V1 Prioritizes High-Volume, Low-Clinical-Risk Inquiries; Complex Cost Estimators Deferred")

# RICE Table
t7_shape = s7.shapes.add_table(8, 7, Inches(0.8), Inches(1.85), Inches(8.5), Inches(4.9))
t7 = t7_shape.table
t7.columns[0].width = Inches(2.2)
t7.columns[1].width = Inches(0.9)
t7.columns[2].width = Inches(0.8)
t7.columns[3].width = Inches(1.1)
t7.columns[4].width = Inches(1.0)
t7.columns[5].width = Inches(1.1)
t7.columns[6].width = Inches(1.4)

r_headers = ["Use Case / Feature", "Reach", "Impact", "Confidence", "Effort", "RICE Score", "Release Scope"]
for c_idx, htext in enumerate(r_headers):
    cell = t7.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

rice_rows = [
    ("Coverage & Benefits Q&A", "5.0M", "3", "80%", "2 wks", "6,000", "✅ V1 MVP (Core)"),
    ("Deductible & OOP Accumulators", "5.0M", "3", "90%", "2 wks", "6,750", "✅ V1 MVP (Core)"),
    ("Claims Status & EOB Explainer", "5.0M", "3", "80%", "3 wks", "4,000", "✅ V1 MVP (Core)"),
    ("ID Card & Plan Doc Download", "5.0M", "2", "90%", "1 wk", "9,000", "✅ V1 MVP (Quick Win)"),
    ("Warm Agent Handoff w/ Context", "5.0M", "3", "90%", "2 wks", "6,750", "✅ V1 MVP (Safety)"),
    ("Treatment Cost Estimator", "4.0M", "4", "60%", "6 wks", "1,600", "⏸️ V2 (Complex Data)"),
    ("In-Network Provider Finder", "4.0M", "3", "70%", "5 wks", "1,680", "⏸️ V2 (Directory API)")
]

for r_idx, r_vals in enumerate(rice_rows, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t7.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(9)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        if c_idx == 6:
            p.font.bold = True
            p.font.color.rgb = C_GREEN if "V1" in val else C_MUTED
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right Column: Explicit Non-Goals & Rationale
add_card(s7, 9.5, 1.85, 3.033, 4.9)
tb_ng = s7.shapes.add_textbox(Inches(9.65), Inches(2.0), Inches(2.7), Inches(4.5))
tfng = tb_ng.text_frame
tfng.word_wrap = True

png_h = tfng.paragraphs[0]
png_h.text = "Explicit V1 Non-Goals"
png_h.font.size = Pt(12.5)
png_h.font.bold = True
png_h.font.color.rgb = C_NAVY

ng_items = [
    ("No Medical Advice / Diagnosis:", "Strict legal and clinical guardrail; assistant is an administrative navigator, not a clinician."),
    ("No Real-Time Cost Estimator:", "Deferred to V2. Contracted fee schedules and anesthesia/facility variables pose high liability for pricing inaccuracies."),
    ("No Rx Drug Formulary Tiering:", "Requires real-time PBM API integration with strict therapeutic interchange logic (V2 roadmap)."),
    ("No Premium Billing Disputes:", "Billing payment adjustments require licensed financial rep sign-off.")
]

for ngh, ngb in ng_items:
    p_1 = tfng.add_paragraph()
    p_1.text = ngh
    p_1.font.size = Pt(10)
    p_1.font.bold = True
    p_1.font.color.rgb = C_PURPLE
    p_1.space_before = Pt(6)
    p_2 = tfng.add_paragraph()
    p_2.text = ngb
    p_2.font.size = Pt(9)
    p_2.font.color.rgb = C_NAVY
    p_2.space_before = Pt(2)

add_footer(s7, 7)
add_speaker_notes(s7,
    "As Product Owner, my most critical responsibility is defining what NOT to build in V1. Through our RICE prioritization, we established that Coverage Q&A, Deductible Accumulators, and Claims Explainers offer the highest reach and confidence with manageable effort. Conversely, I explicitly rejected the Procedure Cost Estimator for V1. Why? Because delivering an inaccurate surgical quote destroys member trust and creates financial liability. That belongs in V2 after data pipelines mature.",
    "If a major employer client demands a Cost Estimator before signing, how do you handle it?",
    "I would offer a phased compromise: in V1, provide historical out-of-pocket range benchmarks with heavy disclaimers, while scheduling the full real-time contracted estimator as the anchor feature of V2, protecting data integrity without losing the deal.")

# ==============================================================================
# SLIDE 8: SOLUTION & EXPERIENCE DESIGN
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
apply_background(s8)
add_header(s8, "Solution & Experience Design", "A Transparent Conversational Flow Adapts Reading Level and Preserves Handoff Context")

# 4 Flow Steps Across Top
flow_steps = [
    ("1. Authenticate & Profile", "Member signs in via SSO/MFA. Assistant accesses member's specific plan ID, deductible accumulators, and language preference."),
    ("2. Natural Language Query", "Member types or speaks: 'Is my knee MRI covered and what will it cost?' Intent classified with entity extraction."),
    ("3. Grounded Retrieval & Synthesis", "Retrieval engine pulls Section 4 of 2026 SBC. Generates 6th-grade response with exact document citation."),
    ("4. Confidence Gate & Routing", "If confidence >=0.7: present answer + 'Explain Simpler' button. If <0.7: initiate warm handoff with full payload.")
]

for idx, (f_title, f_desc) in enumerate(flow_steps):
    left_pos = 0.8 + (idx * 2.95)
    add_card(s8, left_pos, 1.85, 2.85, 2.2, bg_color=C_CARD_BG, border_color=C_BORDER)
    tb_fs = s8.shapes.add_textbox(Inches(left_pos + 0.1), Inches(1.95), Inches(2.65), Inches(2.0))
    tfs = tb_fs.text_frame
    tfs.word_wrap = True
    p1 = tfs.paragraphs[0]
    p1.text = f_title
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = C_TEAL
    p2 = tfs.add_paragraph()
    p2.text = f_desc
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = C_NAVY
    p2.space_before = Pt(4)

# Bottom Container: 3 Core UX Differentiators
add_card(s8, 0.8, 4.25, 11.733, 2.5, bg_color=C_CARD_BG, border_color=C_BORDER)
tb_ux_h = s8.shapes.add_textbox(Inches(1.0), Inches(4.35), Inches(11.3), Inches(0.35))
tb_ux_h.text_frame.paragraphs[0].text = "Three Signature Healthcare UX Differentiators"
tb_ux_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_ux_h.text_frame.paragraphs[0].font.bold = True
tb_ux_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

ux_diffs = [
    ("✨ 'Explain Simpler' Button", "Every response features a 1-click simplification trigger that re-synthesizes the answer at a 4th-grade Flesch-Kincaid level, removing jargon for vulnerable or ESL members.", C_TEAL),
    ("📄 Verifiable Plan Source Citations", "Under every answer, the assistant cites the exact plan document section (e.g., 'Source: 2026 Summary of Benefits, Section 4: Diagnostic Imaging, p. 18').", C_PURPLE),
    ("🤝 Context-Preserving Warm Handoff", "When escalating, the live agent's CRM screen automatically pre-populates the member's query, SBC section, and AI attempt—the member NEVER repeats themselves.", C_GREEN)
]

for u_idx, (u_title, u_desc, u_color) in enumerate(ux_diffs):
    u_left = 1.0 + (u_idx * 3.8)
    card_u = add_card(s8, u_left, 4.8, 3.6, 1.75, bg_color=C_BG_LIGHT, border_color=C_BORDER)
    tb_u = s8.shapes.add_textbox(Inches(u_left + 0.15), Inches(4.9), Inches(3.3), Inches(1.5))
    tfu = tb_u.text_frame
    tfu.word_wrap = True
    pu1 = tfu.paragraphs[0]
    pu1.text = u_title
    pu1.font.size = Pt(11)
    pu1.font.bold = True
    pu1.font.color.rgb = u_color
    pu2 = tfu.add_paragraph()
    pu2.text = u_desc
    pu2.font.size = Pt(9.5)
    pu2.font.color.rgb = C_NAVY
    pu2.space_before = Pt(4)

add_footer(s8, 8)
add_speaker_notes(s8,
    "Here is the experience design in motion. A member asks if their MRI is covered. The assistant queries their specific plan document, recognizes that diagnostic imaging requires prior auth, and calculates their remaining deductible in plain language. But notice our three signature UX features: the 'Explain Simpler' button which instantly drops readability to a 4th-grade level, exact document citation links, and most importantly, the zero-repetition handoff payload passed to the live representative.",
    "What if the member clicks 'Explain Simpler' and still doesn't understand?",
    "If 'Explain Simpler' is clicked twice in a single session without resolution, the system interprets this as an implicit breakdown and proactively offers: 'This is complex—would you like me to connect you with a plan specialist right now?'")

# ==============================================================================
# SLIDE 9: ARCHITECTURE & DATA FLOW
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
apply_background(s9)
add_header(s9, "Business-Readable Architecture", "A Grounded RAG Pipeline Ensures Zero Ungrounded Generation and Full Auditability")

# Layer Diagram (Horizontal Stack)
layers = [
    ("CHANNEL LAYER", "Mobile App (iOS/Android)  •  Web Member Portal  •  Two-Way SMS  •  Interactive Voice Response (IVR, V2)", C_TEAL),
    ("SECURITY & AUTH", "OAuth 2.0 / SAML SSO  •  Multi-Factor Authentication (MFA)  •  HIPAA Consent Check  •  Session Encryption (TLS 1.3)", C_NAVY),
    ("ORCHESTRATION & SAFETY", "Semantic Intent Classifier  •  Clinical Emergency Regex & Classifier  •  PII/PHI Sanitizer  •  Routing Controller", C_PURPLE),
    ("GROUNDED RETRIEVAL (RAG)", "Vector DB of Plan Documents (SBC PDFs, Benefit Grids)  •  Real-Time Accumulator APIs (Deductible / OOP Max)", C_TEAL),
    ("LLM SYNTHESIS & AUDIT", "Restricted Enterprise LLM (Zero Training Retention)  •  Confidence Scorer  •  Immutable Audit Log Store", C_NAVY)
]

for l_idx, (l_title, l_desc, l_col) in enumerate(layers):
    l_top = 1.85 + (l_idx * 0.95)
    card_l = add_card(s9, 0.8, l_top, 11.733, 0.85, bg_color=C_CARD_BG, border_color=l_col)
    
    # Left tag
    tb_lt = s9.shapes.add_textbox(Inches(1.0), Inches(l_top + 0.15), Inches(3.0), Inches(0.55))
    tflt = tb_lt.text_frame
    p1 = tflt.paragraphs[0]
    p1.text = l_title
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = l_col
    
    # Right desc
    tb_ld = s9.shapes.add_textbox(Inches(4.2), Inches(l_top + 0.15), Inches(8.1), Inches(0.55))
    tfld = tb_ld.text_frame
    p2 = tfld.paragraphs[0]
    p2.text = l_desc
    p2.font.size = Pt(10)
    p2.font.color.rgb = C_NAVY

# Bottom Technical Summary Callout
add_card(s9, 0.8, 6.4, 11.733, 0.55, bg_color=C_PURPLE_LIGHT, border_color=C_PURPLE)
tb_h_ctrl = s9.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.45))
p_hctrl = tb_h_ctrl.text_frame.paragraphs[0]
p_hctrl.text = "HALLUCINATION CONTROL: Model context is strictly bounded to retrieved chunks. System prompt: 'If the retrieved text does not contain the explicit benefit answer, reply: I cannot find this in your plan documents and trigger warm handoff.'"
p_hctrl.font.size = Pt(9.5)
p_hctrl.font.bold = True
p_hctrl.font.color.rgb = C_PURPLE

add_footer(s9, 9)
add_speaker_notes(s9,
    "This architecture is built for auditability and compliance. The key distinction from consumer AI is that our LLM has zero direct internet access and is never trained on member data. Instead, it operates strictly as a synthesis engine over retrieved chunks from the member's verified Summary of Benefits and real-time claims accumulators. Every query passes through a clinical emergency filter before intent classification. If confidence falls below 0.7, the system is hardcoded to route to human agents.",
    "How do you handle document versioning when health plan benefits change mid-year?",
    "Every document chunk in our vector database is tagged with Plan_ID, Group_Number, and Effective_Date_Range. Ingesting a mid-year plan amendment automatically deprecates the older chunk, ensuring the RAG engine retrieves only active benefit rules.")

# ==============================================================================
# SLIDE 10: TRUST, SAFETY, COMPLIANCE & FAIRNESS
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
apply_background(s10)
add_header(s10, "Trust, Safety & Equity", "Clinical Safety, PHI Privacy, and Cross-Population Parity Are Operational Gates")

# 3 Pillars
p_cols = [
    ("🛡️ HIPAA & PHI Privacy", [
        ("Minimum Necessary Access:", "Assistant pulls only the accumulator or benefit line needed for the specific intent."),
        ("Zero Training Retention:", "Enterprise BAA ensures prompts and member data are never used for model training."),
        ("Data Encryption:", "End-to-end TLS 1.3 in transit and AES-256 at rest; transcripts purged after 90 days.")
    ], C_TEAL),
    ("🚨 Clinical Safety Guardrails", [
        ("No Diagnostic Advice:", "Hard refusal on symptom self-treatment; assistant states: 'I cannot provide clinical advice.'"),
        ("Emergency Intercept:", "Immediate regex + NLP detection of chest pain, stroke, self-harm → renders 911/988 directives."),
        ("CMS Appeals & Grievances:", "Mandatory warm handoff for Medicare Advantage complaints, meeting CMS compliance.")
    ], C_RED),
    ("⚖️ Cross-Population Fairness", [
        ("Cross-Segment FCR Monitoring:", "Hard operational ceiling: maximum 5pp FCR gap between Commercial, MA, and Medicaid."),
        ("Multilingual Validation:", "Spanish MVP tested against native plan docs, not machine-translated English."),
        ("Literacy Invariance:", "Readability target: 6th-grade default, adaptable to 4th-grade via 'Explain Simpler'.")
    ], C_PURPLE)
]

for col_idx, (col_hdr, items, col_theme) in enumerate(p_cols):
    c_left = 0.8 + (col_idx * 3.95)
    add_card(s10, c_left, 1.85, 3.833, 4.9, bg_color=C_CARD_BG, border_color=C_BORDER)
    
    tb_ch = s10.shapes.add_textbox(Inches(c_left + 0.15), Inches(2.0), Inches(3.5), Inches(0.45))
    tb_ch.text_frame.paragraphs[0].text = col_hdr
    tb_ch.text_frame.paragraphs[0].font.size = Pt(12)
    tb_ch.text_frame.paragraphs[0].font.bold = True
    tb_ch.text_frame.paragraphs[0].font.color.rgb = col_theme
    
    tb_ci = s10.shapes.add_textbox(Inches(c_left + 0.15), Inches(2.5), Inches(3.5), Inches(4.1))
    tfci = tb_ci.text_frame
    tfci.word_wrap = True
    
    for it_hdr, it_body in items:
        p_h = tfci.add_paragraph() if tfci.paragraphs[0].text else tfci.paragraphs[0]
        p_h.text = f"• {it_hdr}"
        p_h.font.size = Pt(10)
        p_h.font.bold = True
        p_h.font.color.rgb = C_NAVY
        p_h.space_before = Pt(6)
        
        p_b = tfci.add_paragraph()
        p_b.text = it_body
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(2)

add_footer(s10, 10)
add_speaker_notes(s10,
    "Safety is not an add-on; it is our license to operate. On this slide we detail our three defense pillars. Under clinical safety, any indication of chest pain, shortness of breath, or behavioral distress bypasses the generative LLM entirely and renders an immediate 911/988 emergency directive with a live nurse transfer. On fairness, we enforce a strict 5-percentage-point parity threshold: if the FCR for Medicaid or Spanish-speaking members lags Commercial by more than 5 points, it triggers an immediate quality hold.",
    "How do you test that the emergency classifier doesn't produce false negatives on atypical distress phrases?",
    "We conduct rigorous monthly red-teaming using a synthetic adversarial dataset of 5,000 clinically validated distress phrasing variants across English and Spanish. Safety recall must remain above 99.0% to pass CI/CD pipeline deployment.")

# ==============================================================================
# SLIDE 11: BUSINESS CASE & FINANCIAL SCENARIOS
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
apply_background(s11)
add_header(s11, "Business Case & Financial Scenarios", "A Disciplined Financial Model Demonstrates Month-14 Break-Even and $5.2M 3-Year NPV")

# Left Column: Financial Model Table
add_card(s11, 0.8, 1.85, 7.5, 4.9)
tb_tab_hdr = s11.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(7.1), Inches(0.4))
tb_tab_hdr.text_frame.paragraphs[0].text = "3-Year Pro-Forma Scenarios (All Figures Illustrative Assumptions)"
tb_tab_hdr.text_frame.paragraphs[0].font.size = Pt(12)
tb_tab_hdr.text_frame.paragraphs[0].font.bold = True
tb_tab_hdr.text_frame.paragraphs[0].font.color.rgb = C_NAVY

t11_shape = s11.shapes.add_table(6, 4, Inches(1.0), Inches(2.45), Inches(7.1), Inches(4.0))
t11 = t11_shape.table
t11.columns[0].width = Inches(2.6)
t11.columns[1].width = Inches(1.5)
t11.columns[2].width = Inches(1.5)
t11.columns[3].width = Inches(1.5)

t11_headers = ["Financial Metric", "Conservative", "Base Case (Target)", "Aggressive"]
for c_idx, htext in enumerate(t11_headers):
    cell = t11.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

t11_data = [
    ("Call Deflection Rate", "10.0%", "20.0%", "30.0%"),
    ("Annual Calls Deflected", "200,000", "400,000", "600,000"),
    ("Gross Annual Cost Savings", "$2,000,000", "$4,000,000", "$6,000,000"),
    ("Year 1 Total Investment", "$3,200,000", "$3,200,000", "$3,200,000"),
    ("3-Year Cumulative NPV (10% WACC)", "$2,100,000", "$5,200,000", "$8,900,000")
]

for r_idx, r_vals in enumerate(t11_data, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t11.cell(r_idx, c_idx)
        cell.fill.solid()
        if r_idx == 5:
            cell.fill.fore_color.rgb = C_PURPLE_LIGHT
        else:
            cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0 or r_idx == 5:
            p.font.bold = True
        if r_idx == 5 and c_idx > 0:
            p.font.color.rgb = C_PURPLE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right Column: Investment Breakdown & Strategic Value
add_card(s11, 8.5, 1.85, 4.033, 4.9)
tb_inv = s11.shapes.add_textbox(Inches(8.7), Inches(2.0), Inches(3.6), Inches(4.5))
tfi = tb_inv.text_frame
tfi.word_wrap = True

pi_h = tfi.paragraphs[0]
pi_h.text = "Cost Structure & Intangibles"
pi_h.font.size = Pt(13)
pi_h.font.bold = True
pi_h.font.color.rgb = C_NAVY

inv_details = [
    ("Year 1 CapEx / Implementation:", "$3.2M ($1.8M engineering, $800k data/RAG infrastructure, $400k security/compliance audit, $200k training)."),
    ("Ongoing OpEx (Years 2-3):", "$1.4M / year ($600k LLM token compute & cloud host, $500k human review team 4 FTE, $300k model maintenance)."),
    ("Strategic Non-Call Benefits:", "• +0.5 point member CSAT increase\n• Reduced agent attrition (saves $15k per retained representative)\n• Improved HEDIS CAHPS scores impacting quality bonus payments.")
]

for ih, ib in inv_details:
    p_1 = tfi.add_paragraph()
    p_1.text = ih
    p_1.font.size = Pt(10.5)
    p_1.font.bold = True
    p_1.font.color.rgb = C_PURPLE if "CapEx" in ih else C_TEAL
    p_1.space_before = Pt(8)
    p_2 = tfi.add_paragraph()
    p_2.text = ib
    p_2.font.size = Pt(9.5)
    p_2.font.color.rgb = C_NAVY
    p_2.space_before = Pt(2)

add_footer(s11, 11)
add_speaker_notes(s11,
    "Turning to the business case: our Year 1 total investment is modeled at $3.2 million, encompassing engineering, vector database construction, compliance audits, and initial rollout. In our base case, net annual operating savings of $2.6 million yield a break-even at month 14 and a 3-year NPV of $5.2 million. Beyond call deflection, this delivers strategic value through improved member retention and higher HEDIS CAHPS survey scores.",
    "What is your response if token compute costs surge due to high user adoption?",
    "Conversational RAG uses prompt-caching and vector search to minimize token loads. Furthermore, our business case shows that each deflected call saves $10, whereas the LLM inference cost for a typical 4-turn chat is under $0.05. Higher adoption actually expands our operating margin.")

# ==============================================================================
# SLIDE 12: METRICS, PILOT & ROLLOUT
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
apply_background(s12)
add_header(s12, "Measurement, Pilot & Scale Gates", "A 90-Day Randomized A/B Pilot Validates Causal Value Before Scaled Enterprise Rollout")

# Top Half: Pilot Experimentation Design
add_card(s12, 0.8, 1.85, 11.733, 2.3, bg_color=C_CARD_BG, border_color=C_BORDER)
tb_pilot = s12.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.3), Inches(2.0))
tfp = tb_pilot.text_frame
tfp.word_wrap = True

pp_h = tfp.paragraphs[0]
pp_h.text = "90-Day Randomized Controlled Pilot Architecture"
pp_h.font.size = Pt(12)
pp_h.font.bold = True
pp_h.font.color.rgb = C_NAVY

pilot_cards = [
    ("Population Cohort:", "Two representative plan books: 1 Commercial Employer Group (50k members) + 1 Medicaid Managed Care plan (50k members)."),
    ("Randomized Assignment:", "50% Test group receives the AI Assistant widget; 50% Control group receives current standard portal search & FAQ navigation."),
    ("Causal Measurement:", "Proves whether deflection and CSAT changes are causally attributable to the AI assistant versus seasonal claims cycles.")
]
for p_lbl, p_val in pilot_cards:
    p_l = tfp.add_paragraph()
    p_l.text = f"{p_lbl} {p_val}"
    p_l.font.size = Pt(10)
    p_l.font.color.rgb = C_NAVY
    p_l.space_before = Pt(4)

# Bottom Half: Two Cards (Scale Gates vs Kill Criteria)
add_card(s12, 0.8, 4.35, 5.7, 2.5, bg_color=C_TEAL_LIGHT, border_color=C_TEAL)
tb_scale = s12.shapes.add_textbox(Inches(1.0), Inches(4.45), Inches(5.3), Inches(2.2))
tfs = tb_scale.text_frame
tfs.word_wrap = True
ps_h = tfs.paragraphs[0]
ps_h.text = "🟢 Scale Decision Gates (Proceed to V2)"
ps_h.font.size = Pt(12)
ps_h.font.bold = True
ps_h.font.color.rgb = C_TEAL

scale_gates = [
    "• First-Contact Resolution (FCR) >= 60% in pilot cohort",
    "• Net call deflection >= 15% across target automatable intents",
    "• Member CSAT >= 4.2 / 5.0 on post-session pulse",
    "• Safety escalation recall >= 99.0% (Zero missed emergencies)",
    "• Cross-segment equity gap <= 5.0 pp (Commercial vs Medicaid)"
]
for sg in scale_gates:
    p_g = tfs.add_paragraph()
    p_g.text = sg
    p_g.font.size = Pt(9.5)
    p_g.font.color.rgb = C_NAVY
    p_g.space_before = Pt(3)

add_card(s12, 6.7, 4.35, 5.833, 2.5, bg_color=RGBColor(254, 242, 242), border_color=C_RED)
tb_kill = s12.shapes.add_textbox(Inches(6.9), Inches(4.45), Inches(5.4), Inches(2.2))
tfk = tb_kill.text_frame
tfk.word_wrap = True
pk_h = tfk.paragraphs[0]
pk_h.text = "🔴 Kill / Pivot Criteria (Pause & Remediate)"
pk_h.font.size = Pt(12)
pk_h.font.bold = True
pk_h.font.color.rgb = C_RED

kill_gates = [
    "• Safety trigger recall drops <95.0% (Immediate hard pause)",
    "• Verified hallucination rate on covered benefits exceeds 1.5%",
    "• Medicaid member CSAT drops >0.5 points below control group",
    "• Live agent handle time (AHT) increases due to messy handoffs",
    "• Unresolvable BAA or CMS compliance roadblock emerges"
]
for kg in kill_gates:
    p_k = tfk.add_paragraph()
    p_k.text = kg
    p_k.font.size = Pt(9.5)
    p_k.font.color.rgb = C_NAVY
    p_k.space_before = Pt(3)

add_footer(s12, 12)
add_speaker_notes(s12,
    "To validate value, we run a 90-day randomized controlled trial with 100,000 members split evenly between test and control. We don't just roll this out and hope for the best. On the bottom of this slide, notice our explicit go/no-go gates. To scale, we must hit 60% FCR and under 5pp equity disparity. If safety recall falls below 95% or hallucination exceeds 1.5%, we execute an immediate operational halt to remediate.",
    "Why not pilot in a single commercial group first before involving Medicaid?",
    "Testing in commercial-only creates a survivorship bias—commercial members generally have higher digital and health literacy. Testing Medicaid concurrently ensures our plain-language models and accessibility features are validated under real-world conditions from day one.")

# ==============================================================================
# SLIDE 13: AGILE EXECUTION (PRD, STORIES, PI PLAN)
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
apply_background(s13)
add_header(s13, "Agile Execution & Delivery", "Delivering MVP in 16 Weeks Across Three Epics with Rigid Given/When/Then Criteria")

# Left Column: 5 Sample User Stories
add_card(s13, 0.8, 1.85, 7.3, 4.9)
tb_us_h = s13.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(6.9), Inches(0.4))
tb_us_h.text_frame.paragraphs[0].text = "Core User Stories with Acceptance Criteria"
tb_us_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_us_h.text_frame.paragraphs[0].font.bold = True
tb_us_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

stories = [
    ("US-01: Coverage Q&A (Core Engine)", "Given an authenticated member asks 'Is my physical therapy covered?', When the RAG engine retrieves Section 4 of the plan SBC with confidence >0.7, Then display copay amount, visit limit, and citation link within 3 seconds."),
    ("US-02: Safety & Emergency Intercept (Safety)", "Given a member types 'I am experiencing severe chest pain', When the safety classifier intercepts the prompt, Then bypass LLM, render red 911 emergency directive, and auto-initiate warm agent transfer within 1.5 seconds."),
    ("US-03: Real-Time Deductible Tracker (Account)", "Given an HDHP member asks 'How much deductible do I have left?', When claims accumulator API returns data <60s old, Then render visual progress bar showing paid vs remaining balances."),
    ("US-04: Multilingual Parity (Fairness)", "Given a Spanish-speaking Medicaid member submits a query in Spanish, When the assistant synthesizes the answer, Then response is rendered in native Spanish with equal factual accuracy (<5pp FCR gap)."),
    ("US-05: Context-Preserving Handoff (Ops)", "Given the assistant initiates human escalation, When the live agent accepts the session, Then the CRM screen displays member ID, intent, partial transcript, and reason for escalation.")
]

tb_st_body = s13.shapes.add_textbox(Inches(1.0), Inches(2.4), Inches(6.9), Inches(4.2))
tfst = tb_st_body.text_frame
tfst.word_wrap = True

for s_title, s_crit in stories:
    p_t = tfst.add_paragraph() if tfst.paragraphs[0].text else tfst.paragraphs[0]
    p_t.text = s_title
    p_t.font.size = Pt(9.5)
    p_t.font.bold = True
    p_t.font.color.rgb = C_PURPLE if "Safety" in s_title or "Fairness" in s_title else C_TEAL
    p_t.space_before = Pt(4)
    p_c = tfst.add_paragraph()
    p_c.text = s_crit
    p_c.font.size = Pt(8.5)
    p_c.font.color.rgb = C_NAVY
    p_c.space_before = Pt(1)

# Right Column: PI Plan & Releases
add_card(s13, 8.3, 1.85, 4.233, 4.9)
tb_pi = s13.shapes.add_textbox(Inches(8.5), Inches(2.0), Inches(3.8), Inches(4.5))
tfpi = tb_pi.text_frame
tfpi.word_wrap = True

ppi_h = tfpi.paragraphs[0]
ppi_h.text = "One-Program Increment (PI) Roadmap"
ppi_h.font.size = Pt(12)
ppi_h.font.bold = True
ppi_h.font.color.rgb = C_NAVY

pi_sprints = [
    ("Sprint 1–2 (Wks 1–4): Foundation", "• OAuth/SSO integration, HIPAA compliance review\n• Ingest 50 SBC plan documents into vector DB"),
    ("Sprint 3–4 (Wks 5–8): RAG Pipeline", "• Grounded synthesis engine & confidence scoring\n• Build 'Explain Simpler' 4th-grade re-writer"),
    ("Sprint 5–6 (Wks 9–12): Integration", "• Accumulator claims API & Claims explainer\n• Clinical emergency classifier & safety intercept"),
    ("Sprint 7–8 (Wks 13–16): Hardening & Pilot", "• CRM handoff payload & Spanish language pipeline\n• Deploy 90-day A/B pilot in 2 target markets"),
    ("V2 Release (Q3–Q4): Scale", "• Treatment Cost Estimator & Provider Finder\n• IVR voice integration for Medicare seniors")
]

for pih, pib in pi_sprints:
    p_h = tfpi.add_paragraph()
    p_h.text = pih
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_NAVY
    p_h.space_before = Pt(5)
    p_b = tfpi.add_paragraph()
    p_b.text = pib
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_MUTED
    p_b.space_before = Pt(1)

add_footer(s13, 13)
add_speaker_notes(s13,
    "As an Agile Product Owner, I maintain rigorous Given/When/Then acceptance criteria. Notice Story 2 and Story 4: safety and fairness are treated as first-class backlog tickets, not afterthoughts. Story 2 requires emergency prompts to bypass the LLM and render a 911 directive in under 1.5 seconds. Our PI roadmap is structured into 8 two-week sprints: foundational infrastructure in sprints 1-4, accumulator and safety integration in sprints 5-6, and pilot launch at sprint 8.",
    "How do you handle a scenario where engineering falls behind on the Spanish language model during Sprint 6?",
    "Under SAFe WSJF (Weighted Shortest Job First), if capacity is constrained, we never compromise safety or security. We would launch the pilot with English fully hardened, while isolating the Spanish language pipeline into a fast-follower release in Sprint 9, maintaining transparency with clinical equity stakeholders.")

# ==============================================================================
# SLIDE 14: STAKEHOLDERS, RISKS & 90-DAY PLAN
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
apply_background(s14)
add_header(s14, "Governance, Risks & 90-Day Plan", "Cross-Functional Alignment and Five Leadership Decisions Unlock Rapid Time-to-Value")

# Top Half: 90-Day Plan (3 Columns)
add_card(s14, 0.8, 1.85, 11.733, 2.3, bg_color=C_CARD_BG, border_color=C_BORDER)
tb_90 = s14.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.3), Inches(2.0))
tf90 = tb_90.text_frame
tf90.word_wrap = True

p90_h = tf90.paragraphs[0]
p90_h.text = "90-Day Transition & Execution Plan (Day 1 to Decision Gate)"
p90_h.font.size = Pt(12)
p90_h.font.bold = True
p90_h.font.color.rgb = C_NAVY

phases = [
    ("Days 1–30: Discovery & Data Audit", "• 15 stakeholder interviews across operations, clinical, and compliance.\n• Audit call center IVR categorizations and sample 200 plan SBC documents.\n• Finalize enterprise BAA with LLM vendor and lock PRD scope."),
    ("Days 31–60: Build & Safety Red-Team", "• Stand up vector database and claims accumulator API connectors.\n• Conduct adversarial clinical red-teaming (5,000 emergency test prompts).\n• Implement agent CRM handoff payload and 'Explain Simpler' engine."),
    ("Days 61–90: Pilot Launch & Evaluation", "• Launch 90-day A/B pilot in 1 Commercial and 1 Medicaid market (100k members).\n• Monitor real-time telemetry (FCR, hallucination rate, equity gap).\n• Executive review at Day 90 with recommendation to scale or pivot.")
]

for ph_t, ph_b in phases:
    p_ph = tf90.add_paragraph()
    p_ph.text = f"{ph_t}: {ph_b}"
    p_ph.font.size = Pt(9.5)
    p_ph.font.color.rgb = C_NAVY
    p_ph.space_before = Pt(4)

# Bottom Left: RACI & Risks
add_card(s14, 0.8, 4.35, 6.5, 2.5, bg_color=C_CARD_BG, border_color=C_BORDER)
tb_raci = s14.shapes.add_textbox(Inches(1.0), Inches(4.45), Inches(6.1), Inches(2.2))
tfr = tb_raci.text_frame
tfr.word_wrap = True
pr_h = tfr.paragraphs[0]
pr_h.text = "Key Risks & Mitigations"
pr_h.font.size = Pt(11.5)
pr_h.font.bold = True
pr_h.font.color.rgb = C_NAVY

risks_list = [
    ("Hallucination Risk:", "RAG constraint + confidence threshold (<0.7 routes to human). Target <1% error."),
    ("Agent Resistance / FUD:", "Position AI as 'Agent Assistant'; saves 2 mins/call on repetitive SBC lookup."),
    ("Regulatory / CMS Shift:", "Quarterly compliance review of state AI rules and Medicare marketing guidelines.")
]
for rk, rm in risks_list:
    p_rk = tfr.add_paragraph()
    p_rk.text = f"• {rk} {rm}"
    p_rk.font.size = Pt(9)
    p_rk.font.color.rgb = C_NAVY
    p_rk.space_before = Pt(4)

# Bottom Right: 5 Decisions Needed
add_card(s14, 7.5, 4.35, 5.033, 2.5, bg_color=C_PURPLE_LIGHT, border_color=C_PURPLE)
tb_dec = s14.shapes.add_textbox(Inches(7.7), Inches(4.45), Inches(4.6), Inches(2.2))
tfd = tb_dec.text_frame
tfd.word_wrap = True
pd_h = tfd.paragraphs[0]
pd_h.text = "5 Decisions Needed from Leadership"
pd_h.font.size = Pt(11.5)
pd_h.font.bold = True
pd_h.font.color.rgb = C_PURPLE

decs = [
    "1. Authorize Y1 CapEx budget ($3.2M illustrative baseline)",
    "2. Sign off on BAA vendor terms and 90-day PHI retention policy",
    "3. Approve 90-day pilot markets (1 Commercial + 1 Medicaid)",
    "4. Align Call Center Ops to support context-rich warm handoffs",
    "5. Form Cross-Functional Clinical & AI Governance Committee"
]
for d in decs:
    p_d = tfd.add_paragraph()
    p_d.text = d
    p_d.font.size = Pt(9)
    p_d.font.color.rgb = C_NAVY
    p_d.space_before = Pt(3)

add_footer(s14, 14)
add_speaker_notes(s14,
    "To conclude, our 90-day roadmap moves rapidly from discovery to an evaluated pilot. In the first 30 days, we audit claims and SBC feeds and lock the PRD. In days 31 to 60, we build the RAG pipeline and conduct rigorous safety red-teaming. In days 61 to 90, we launch the pilot in two markets. To unblock this, I need five specific decisions from leadership, highlighted on the right, starting with BAA approval and pilot population selection.",
    "What is your immediate priority on Day 1 if you are offered this role?",
    "Day 1 is listening and auditing. I will sit with call-center agents listening to live calls to observe exact member terminology, review the top 50 denial and accumulator reason codes, and establish relationships with our clinical safety and compliance leads. Great product management starts on the front lines.")

# ==============================================================================
# APPENDIX SLIDES (15 - 18)
# ==============================================================================
# Slide 15: Appendix A - Assumptions Register
s15 = prs.slides.add_slide(blank_layout)
apply_background(s15)
add_header(s15, "Appendix A: Assumptions Register", "Full Transparency on Baseline Data, Sizing Formulas, and Operational Benchmarks")

t15_shape = s15.shapes.add_table(8, 4, Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.9))
t15 = t15_shape.table
t15.columns[0].width = Inches(2.8)
t15.columns[1].width = Inches(1.8)
t15.columns[2].width = Inches(4.333)
t15.columns[3].width = Inches(2.8)

a_headers = ["Parameter / Metric", "Assumed Value", "Source Logic & Healthcare Rationale", "Sensitivity Impact"]
for c_idx, htext in enumerate(a_headers):
    cell = t15.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

assumptions_data = [
    ("Total Inbound Call Volume", "5,000,000 calls / yr", "Illustrative mid-to-large payer book of business.", "+/- 1M calls = +/- $800k net benefit."),
    ("Cost per Live Inbound Call", "$10.00 / call", "Industry contact-center loaded benchmark ($8-$12).", "+/- $2/call = +/- $800k annual gross."),
    ("Automatable Intent Share", "40.0% (2.0M calls)", "Coverage Q&A, deductible, claims status, ID cards.", "Core boundary for V1 addressable market."),
    ("Base Call Deflection Rate", "20.0% of automatable", "Derived from best-in-class healthcare AI benchmarks.", "+/- 5pp deflection = +/- $1.0M annual savings."),
    ("Year 1 Implementation CapEx", "$3,200,000", "Engineering, RAG pipeline, audits, compliance.", "Fixed upfront investment amortized over 3 yrs."),
    ("Ongoing Annual OpEx", "$1,400,000 / yr", "LLM tokens, cloud host, 4 FTE review team.", "Scales slightly with token inference volume."),
    ("Equity Disparity Threshold", "Max 5.0 pp gap", "Derived from NCQA HEDIS health equity standards.", "Strict release gate for all algorithmic features.")
]

for r_idx, r_vals in enumerate(assumptions_data, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t15.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s15, 15)
add_speaker_notes(s15, "Appendix A documents every numerical assumption, its operational basis, and its sensitivity impact. No number is presented as real internal Carelon data.", "How sensitive is the business case to changes in live call costs?", "Every $1 change in cost-per-call impacts gross savings by $400,000 under our base case.")

# Slide 16: Appendix B - Glossary
s16 = prs.slides.add_slide(blank_layout)
apply_background(s16)
add_header(s16, "Appendix B: US Healthcare Glossary", "Standard Definitions for Payer Operations, Regulatory Rules, and AI Terminology")

glossary_items = [
    ("Deductible:", "The annual amount a member pays out-of-pocket for covered services before the plan begins paying."),
    ("Out-of-Pocket Maximum (OOP Max):", "The most a member must pay for covered services in a plan year; insurance pays 100% thereafter."),
    ("Coinsurance:", "The member's percentage share of the costs of a covered service (e.g., 20%) paid after deductible is met."),
    ("Copayment (Copay):", "A fixed dollar amount (e.g., $30) paid by the member for a covered healthcare service at time of care."),
    ("Explanation of Benefits (EOB):", "A statement sent by the health plan explaining what medical treatments were billed, paid, and member owes."),
    ("Prior Authorization (PA):", "Pre-service approval required by the insurer before certain tests or drugs are covered."),
    ("Summary of Benefits & Coverage (SBC):", "Standardized, federally mandated document outlining what the plan covers and member costs."),
    ("HEDIS / CAHPS:", "Healthcare Effectiveness Data & Information Set and Consumer Assessment surveys measuring quality & satisfaction."),
    ("PMPM (Per Member Per Month):", "Standard unit of healthcare financial accounting measuring cost per covered individual per month."),
    ("Retrieval-Augmented Generation (RAG):", "AI architecture grounding generative models strictly in retrieved proprietary documents.")
]

add_card(s16, 0.8, 1.85, 11.733, 4.9)
tb_gl = s16.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.5))
tfgl = tb_gl.text_frame
tfgl.word_wrap = True

for term, definition in glossary_items:
    p = tfgl.add_paragraph() if tfgl.paragraphs[0].text else tfgl.paragraphs[0]
    p.text = f"{term} {definition}"
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_NAVY
    p.space_before = Pt(5)

add_footer(s16, 16)
add_speaker_notes(s16, "Appendix B provides standard US healthcare terminology reference definitions for all business and technical terms used in the presentation.", "Why is the distinction between copay and coinsurance so critical in assistant design?", "Over 60% of member billing complaints stem from members confusing flat copays with percentage-based coinsurance after deductible.")

# Slide 17: Appendix C - Risk Register
s17 = prs.slides.add_slide(blank_layout)
apply_background(s17)
add_header(s17, "Appendix C: Detailed Risk Register", "Comprehensive Risk Assessment Across Clinical, Legal, Operations, and Technical Domains")

t17_shape = s17.shapes.add_table(7, 5, Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.9))
t17 = t17_shape.table
t17.columns[0].width = Inches(2.2)
t17.columns[1].width = Inches(1.2)
t17.columns[2].width = Inches(1.2)
t17.columns[3].width = Inches(4.933)
t17.columns[4].width = Inches(2.2)

r_hdrs = ["Risk Description", "Likelihood", "Impact", "Planned Mitigation Strategy", "Owner"]
for c_idx, htext in enumerate(r_hdrs):
    cell = t17.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

risks_data = [
    ("Hallucination on covered benefits", "Medium", "High", "RAG restriction to verified SBC text; confidence score threshold (<0.7 escalates).", "Data Science / PM"),
    ("Safety classifier misses acute distress", "Low", "Critical", "Dual-stage regex + NLP safety classifier; monthly 5k-prompt red-team audits.", "Clinical Quality / Lead"),
    ("State insurance AI regulatory shifts", "Medium", "High", "Active monitoring of NAIC model AI bulletin; legal review before new state rollouts.", "Legal & Compliance"),
    ("Call-center agent resistance / fear", "Medium", "Medium", "Change mgmt: AI pitched as agent helper; agents review drafts; time-savings shown.", "Call Center Ops"),
    ("Disparate impact across languages", "Medium", "High", "Dedicated Spanish evaluation suite; strict 5pp FCR parity threshold gate.", "Product Owner (Me)"),
    ("Claims accumulator API downtime", "Low", "High", "Circuit-breaker pattern: display cached accumulator with timestamp & fallback note.", "IT / Engineering")
]

for r_idx, r_vals in enumerate(risks_data, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t17.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s17, 17)
add_speaker_notes(s17, "Appendix C details our risk register. The highest catastrophic risk is a clinical safety miss, which is mitigated via dual-stage classifiers and monthly red-teaming.", "Who holds the final sign-off authority on safety thresholds?", "The Clinical Quality and Chief Medical Officer representatives hold hard veto power over safety gates.")

# Slide 18: Appendix D - Additional User Stories & NFRs
s18 = prs.slides.add_slide(blank_layout)
apply_background(s18)
add_header(s18, "Appendix D: Extended User Stories & NFRs", "Non-Functional Requirements and Secondary Backlog Stories for V2 Enablement")

add_card(s18, 0.8, 1.85, 11.733, 4.9)
tb_s18 = s18.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.5))
tfs18 = tb_s18.text_frame
tfs18.word_wrap = True

ext_items = [
    ("US-06: EOB Denial Plain-Language Explainer (V1)", "Given a member receives a denied claim notice, When they ask 'Why did my insurance not pay for Dr. Smith?', Then the assistant explains the specific denial reason code (e.g. out-of-network provider) in 6th-grade language and outlines the 90-day appeal timeline."),
    ("US-07: Low-Confidence Honest Handshake (V1 Safety)", "Given a member asks an ambiguous question where RAG retrieval confidence is <0.7, When the assistant responds, Then it states 'I cannot verify this specific exception in your plan documents' and offers a 1-click warm transfer."),
    ("NFR-01: Response Latency", "95% of standard conversational Q&A queries must return the complete synthesized response within 2.5 seconds under peak concurrency of 5,000 active sessions."),
    ("NFR-02: Accessibility Standards (WCAG 2.1 AA)", "All digital assistant interfaces must maintain a minimum 4.5:1 color contrast ratio, full keyboard navigability, and full compatibility with screen-readers (JAWS / NVDA)."),
    ("NFR-03: Security & Session Isolation", "Each member chat session must be strictly isolated in memory; session memory is flushed immediately upon logout or after 15 minutes of inactivity.")
]

for title_e, body_e in ext_items:
    p1 = tfs18.add_paragraph() if tfs18.paragraphs[0].text else tfs18.paragraphs[0]
    p1.text = title_e
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_PURPLE if "NFR" in title_e else C_TEAL
    p1.space_before = Pt(5)
    p2 = tfs18.add_paragraph()
    p2.text = body_e
    p2.font.size = Pt(9)
    p2.font.color.rgb = C_NAVY
    p2.space_before = Pt(1)

add_footer(s18, 18)
add_speaker_notes(s18, "Appendix D outlines secondary backlog stories and critical non-functional requirements including 2.5s response latency, WCAG 2.1 AA compliance, and strict memory isolation.", "Why is 2.5-second latency critical in conversational RAG?", "In digital chat, latency exceeding 3 seconds causes perceived friction and a 20% increase in premature member bounce or channel switching.")

# Save presentation
out_path = r"D:\Preksha_carelon\Carelon_AI_Assistant_Case_Study.pptx"
prs.save(out_path)
print(f"SUCCESS: Saved 18-slide presentation to {out_path}")
