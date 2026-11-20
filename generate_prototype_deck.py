import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

C_NAVY = RGBColor(26, 26, 46)        # #1A1A2E - Deep Navy Primary
C_TEAL = RGBColor(13, 124, 143)      # #0D7C8F - Calm Healthcare Teal
C_PURPLE = RGBColor(91, 45, 142)     # #5B2D8E - Carelon Brand Purple
C_BG_LIGHT = RGBColor(248, 249, 250) # #F8F9FA - Neutral Canvas
C_CARD_BG = RGBColor(255, 255, 255)  # White Card
C_BORDER = RGBColor(226, 230, 239)   # Subtle Border
C_MUTED = RGBColor(106, 115, 131)    # Muted Body Text
C_TEAL_LIGHT = RGBColor(235, 246, 248) # Soft Teal Tint
C_PURPLE_LIGHT = RGBColor(246, 240, 255) # Soft Purple Tint
C_GREEN = RGBColor(46, 204, 113)     # #2ECC71 - Success Green
C_RED = RGBColor(231, 76, 60)        # #E74C3C - Red Alert
C_AMBER = RGBColor(243, 156, 18)     # #F39C12 - Warning

def apply_background(slide):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_BG_LIGHT
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, tag_text, headline_text, max_w=12.0):
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.36), Inches(11.5), Inches(0.32))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.name = "Calibri"
    p_tag.font.color.rgb = C_TEAL

    hl_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(max_w), Inches(0.85))
    tf_hl = hl_box.text_frame
    tf_hl.word_wrap = True
    tf_hl.margin_left = tf_hl.margin_top = tf_hl.margin_right = tf_hl.margin_bottom = 0
    p_hl = tf_hl.paragraphs[0]
    p_hl.text = headline_text
    p_hl.font.size = Pt(20)
    p_hl.font.bold = True
    p_hl.font.name = "Calibri"
    p_hl.font.color.rgb = C_NAVY

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.58), Inches(1.5), Inches(0.04))
    line.fill.solid()
    line.fill.fore_color.rgb = C_PURPLE
    line.line.fill.background()

def add_proto_callout(slide, view_name, click_action, left=0.8, top=6.4, width=11.733, height=0.55):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = C_PURPLE_LIGHT
    card.line.color.rgb = C_PURPLE
    card.line.width = Pt(1)

    tb = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.08), Inches(width - 0.3), Inches(height - 0.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"💻 SEE IT IN THE PROTOTYPE: [{view_name}]  →  {click_action}"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = C_PURPLE
    p.font.name = "Calibri"

def add_footer(slide, current_idx, total_slides=18):
    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(11.733), Inches(0.32))
    tf = footer_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"Carelon Global Solutions | Sr. BA II Submission | Member AI Assistant (Prototype-Aligned)  •  Illustrative / Synthetic Data              Slide {current_idx:02d} of {total_slides:02d}"
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
add_card(s1, 0.8, 0.8, 11.733, 5.9, bg_color=C_CARD_BG, border_color=C_BORDER)

accent_bar = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(0.2), Inches(5.9))
accent_bar.fill.solid()
accent_bar.fill.fore_color.rgb = C_PURPLE
accent_bar.line.fill.background()

tb_title = s1.shapes.add_textbox(Inches(1.4), Inches(1.2), Inches(10.5), Inches(3.2))
tf1 = tb_title.text_frame
tf1.word_wrap = True

p0 = tf1.paragraphs[0]
p0.text = "CARELON GLOBAL SOLUTIONS | PRODUCT STRATEGY CASE SUBMISSION"
p0.font.size = Pt(11)
p0.font.bold = True
p0.font.color.rgb = C_TEAL

p1 = tf1.add_paragraph()
p1.text = "Designing an AI Assistant for Health Plan Members"
p1.font.size = Pt(28)
p1.font.bold = True
p1.font.color.rgb = C_NAVY
p1.space_before = Pt(8)

p2 = tf1.add_paragraph()
p2.text = "Demonstrated Through the Working Interactive Prototype: 'Your Health Plan Assistant'"
p2.font.size = Pt(14)
p2.font.color.rgb = C_MUTED
p2.space_before = Pt(6)

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
pt_body.text = "By grounding conversational AI strictly in verified member plan documents and accumulator APIs, we achieve 61.3% first-contact resolution and reduce cost per contact from $10.00 to $1.84—proven live in our working prototype across five distinct member personas."
pt_body.font.size = Pt(13)
pt_body.font.color.rgb = C_NAVY
pt_body.font.bold = True
pt_body.space_before = Pt(4)

tb_meta = s1.shapes.add_textbox(Inches(1.4), Inches(5.3), Inches(10.5), Inches(0.8))
tf_meta = tb_meta.text_frame
p_meta = tf_meta.paragraphs[0]
p_meta.text = "Role: Sr. Business Analyst II (Product Owner, US Healthcare Agile)  •  Elevance Health Subsidiary  •  Prototype Version: v1.0 Production Candidate"
p_meta.font.size = Pt(11)
p_meta.font.color.rgb = C_MUTED

add_footer(s1, 1)
add_speaker_notes(s1,
    "Welcome. Today I am presenting an operational product strategy for an AI Assistant designed for health plan members. Every feature, metric, and persona in this deck is backed by a fully interactive prototype I built for this submission. The thesis is simple: we replace multi-channel confusion with document-grounded answers, driving first-contact resolution to 61.3% and cutting cost-per-contact from $10 to $1.84 while protecting clinical safety.",
    "Why build a prototype instead of just delivering static slides?",
    "Product Owners must translate business ambiguity into tangible, working software. Building the prototype allowed me to stress-test prompt groundings, model the RICE backlog mathematically, and verify that our clinical escalation rules actually fire in real time.")

# ==============================================================================
# SLIDE 2: PROBLEM & OPERATIONAL IMPACT
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
apply_background(s2)
add_header(s2, "Problem & Operational Reality", "Members Face a Friction Loop That Drives 78% of Inbound Contacts to High-Cost Phone Queues")

# Left: Escalation Loop Diagram
add_card(s2, 0.8, 1.8, 5.7, 4.4)
tb_loop_h = s2.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.3), Inches(0.35))
tb_loop_h.text_frame.paragraphs[0].text = "The Multi-Channel Escalation Loop Today"
tb_loop_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_loop_h.text_frame.paragraphs[0].font.bold = True
tb_loop_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

loop_items = [
    ("Step 1: Digital Search Friction", "Member logs into portal seeking coverage info (e.g. MRI copay, deductible progress)."),
    ("Step 2: Jargon & FAQ Dead End", "Dense 90-page Summary of Benefits or generic bot fails to answer plan-specific nuances."),
    ("Step 3: Call Escalation (78% Phone)", "Member abandons digital channel and calls live rep. Average wait time 6-8 mins."),
    ("Step 4: Repetitive Cost ($10/contact)", "Agent manually looks up exact same SBC document. Resolution takes 7m 00s.")
]
for i, (l_title, l_desc) in enumerate(loop_items):
    add_card(s2, 1.0, 2.4 + (i * 0.92), 5.3, 0.82, bg_color=C_TEAL_LIGHT if i%2==0 else C_PURPLE_LIGHT, border_color=C_BORDER)
    tb_li = s2.shapes.add_textbox(Inches(1.1), Inches(2.45 + (i * 0.92)), Inches(5.1), Inches(0.72))
    tfli = tb_li.text_frame
    p1 = tfli.paragraphs[0]
    p1.text = l_title
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_TEAL if i%2==0 else C_PURPLE
    p2 = tfli.add_paragraph()
    p2.text = l_desc
    p2.font.size = Pt(9)
    p2.font.color.rgb = C_NAVY

# Right: The Deflection Shift (Matching Prototype Chart)
add_card(s2, 6.7, 1.8, 5.833, 4.4)
tb_def_h = s2.shapes.add_textbox(Inches(6.9), Inches(1.95), Inches(5.4), Inches(0.35))
tb_def_h.text_frame.paragraphs[0].text = "Prototype Telemetry: Call Deflection Before vs After AI"
tb_def_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_def_h.text_frame.paragraphs[0].font.bold = True
tb_def_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

# Table comparing Before vs After (exact numbers from prototype chart)
t2_shape = s2.shapes.add_table(5, 3, Inches(6.9), Inches(2.4), Inches(5.4), Inches(2.4))
t2 = t2_shape.table
t2.columns[0].width = Inches(2.2)
t2.columns[1].width = Inches(1.6)
t2.columns[2].width = Inches(1.6)

t2_data = [
    ("Channel Mix", "Before AI (Legacy)", "After AI (Prototype)"),
    ("📞 Phone Queue", "78.0%", "52.0% (-26pp Deflected)"),
    ("🤖 AI Assistant Resolved", "0.0%", "22.0% (Self-Service)"),
    ("💬 Live Chat", "15.0%", "21.0% (Warm Handoffs)"),
    ("🌐 Web Self-Serve", "7.0%", "5.0% (Legacy Search)")
]
for r_idx, r_vals in enumerate(t2_data):
    for c_idx, val in enumerate(r_vals):
        cell = t2.cell(r_idx, c_idx)
        cell.fill.solid()
        if r_idx == 0:
            cell.fill.fore_color.rgb = C_TEAL
        elif r_idx == 2:
            cell.fill.fore_color.rgb = C_PURPLE_LIGHT
        else:
            cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(9.5)
        p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else (C_PURPLE if r_idx == 2 else C_NAVY)
        if r_idx == 0 or r_idx == 2 or c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

tb_call_cost = s2.shapes.add_textbox(Inches(6.9), Inches(4.95), Inches(5.4), Inches(1.1))
tfc = tb_call_cost.text_frame
p_cc1 = tfc.paragraphs[0]
p_cc1.text = "• Cost-to-Serve Impact: Phone contact cost ($10.00) vs. AI Assistant contact ($1.84)"
p_cc1.font.size = Pt(10)
p_cc1.font.bold = True
p_cc1.font.color.rgb = C_PURPLE
p_cc2 = tfc.add_paragraph()
p_cc2.text = "• Resolution Time: Legacy phone handle time averages 7m 00s vs. Prototype AI average 2m 47s (↓ 4m 13s savings)."
p_cc2.font.size = Pt(9.5)
p_cc2.font.color.rgb = C_MUTED
p_cc2.space_before = Pt(3)

add_proto_callout(s2, "View 2: Product Dashboard", "Navigate to 'Call Deflection: Before vs. After AI' chart & 'Cost per Contact' KPI ($1.84).")
add_footer(s2, 2)
add_speaker_notes(s2,
    "Here we quantify the problem using the exact telemetry in our prototype dashboard. Today, 78% of member inquiries end up in high-cost phone queues because members cannot self-serve through dense plan documents. With the AI Assistant deployed, our telemetry demonstrates phone volume falling to 52%, while AI self-service absorbs 22% of interactions, slashing the loaded cost per contact from $10 down to $1.84.",
    "What specific inquiries represent that 22% resolved by AI?",
    "Our prototype intent distribution shows Coverage Q&A at 34%, Claims/EOB lookup at 28%, and Deductible accumulators at 18%. These are repetitive, deterministic lookups that require zero clinical triage.")

# ==============================================================================
# SLIDE 3: PERSONAS & SEGMENTS (MATCHING DASHBOARD)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
apply_background(s3)
add_header(s3, "Personas & Segmentation", "Five Distinct Personas with Explicit Plan Configurations Directly Built into the Prototype")

# Table of 5 Personas exactly matching View 1
t3_shape = s3.shapes.add_table(6, 5, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.4))
t3 = t3_shape.table
t3.columns[0].width = Inches(2.2)
t3.columns[1].width = Inches(2.0)
t3.columns[2].width = Inches(3.0)
t3.columns[3].width = Inches(2.333)
t3.columns[4].width = Inches(2.2)

h3 = ["Persona & Plan", "Deductible & OOP Max", "Top Job-to-be-Done (JTBD)", "Channel & Tone Need", "Emotional State"]
for c_idx, htext in enumerate(h3):
    cell = t3.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

p_data_dash = [
    ("Alex\nNew to Insurance\nPlan: BlueStar Silver", "Deductible: $3,000 (Paid $450)\nOOP Max: $7,000\nPCP Copay: $30", "Understand basic coverage terms without jargon; verify if urgent care or MRI is covered.", "Mobile app, async chat;\nRequires 6th-grade language", "Anxious, overwhelmed,\nCost-conscious"),
    ("Maria\nFamily Caregiver\nPlan: FamilyCare Gold", "Deductible: $6,000 (Paid $2,100)\nOOP Max: $12,000\nCovered Lives: 2 kids", "Track family aggregate deductible, verify in-network pediatric ENT, appeal denied claim.", "Web portal + mobile app;\nNeeds printable audit trails", "Time-poor, stressed,\nProtective"),
    ("David\nChronic Condition\nPlan: ChronicCare Plus", "Deductible: $1,500 (Paid $1,200)\nOOP Max: $4,000\nRx: Insulin Tier 2", "Verify insulin copays, check prior-authorization renewal, monitor out-of-pocket maximum.", "Web portal + live phone;\nValues consistency & speed", "Frustrated by delays,\nHealth-compromised"),
    ("Eleanor\nSenior / Medicare\nPlan: MediGuard MA", "Deductible: $0\nPCP Copay: $0 | Spec: $45\nPart B Premium: Included", "Understand Medicare Advantage supplemental benefits (vision/dental); request live agent handoff.", "Phone-first preference;\nPatient, warm human voice", "Cautious, vulnerable,\nValues trust"),
    ("Priya\nHDHP Member\nPlan: HSA Select HDHP", "Deductible: $4,000 (Paid $1,850)\nOOP Max: $8,000\nHSA Balance: $2,340", "Calculate remaining deductible before procedure; estimate surgery cost vs HSA balance.", "Self-service web portal;\nAnalytical, numbers-focused", "Strategic, calculating,\nCost-optimizing")
]

for r_idx, r_vals in enumerate(p_data_dash, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t3.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
            p.font.color.rgb = C_PURPLE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_proto_callout(s3, "View 1: Member Experience", "Click the persona pills (Alex, Maria, David, Eleanor, Priya) in the left sidebar to watch plan details dynamically update.")
add_footer(s3, 3)
add_speaker_notes(s3,
    "In the prototype, you can click across all five personas. Each triggers a tailored greeting and injects distinct plan rules: Alex has a $3,000 deductible with $450 paid; Eleanor is on Medicare Advantage with zero deductible and $45 specialist copays; Priya tracks her HSA balance. Designing across these personas forces us to support adaptive reading levels and plan-aware calculations.",
    "Why does Eleanor's persona emphasize phone handoff over AI chat?",
    "CMS regulations for Medicare Advantage mandate clear, accessible pathways to human representatives. For Eleanor, the AI assistant's primary value is providing rapid preliminary answers while always keeping a visible 'Talk to a Person' button front and center.")

# ==============================================================================
# SLIDE 4: MARKET & COMPETITOR SCAN
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
apply_background(s4)
add_header(s4, "Market & Competitive Scan", "The Payer White Space Lies in Contextual Grounding Over Generic Conversational Bots")

# 2x2 Positioning Map Visual
add_card(s4, 0.8, 1.8, 6.8, 4.4)
tb_map_h = s4.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.4), Inches(0.35))
tb_map_h.text_frame.paragraphs[0].text = "Payer AI Positioning Map: Domain Grounding vs Interactive Scope"
tb_map_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_map_h.text_frame.paragraphs[0].font.bold = True
tb_map_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

# 4 Quadrants
def fill_pos(slide, left, top, title, players, col, bg):
    add_card(slide, left, top, 3.1, 1.65, bg_color=bg, border_color=col)
    tb = slide.shapes.add_textbox(Inches(left + 0.1), Inches(top + 0.1), Inches(2.9), Inches(1.45))
    tf = tb.text_frame
    p1 = tf.paragraphs[0]
    p1.text = title
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = col
    p2 = tf.add_paragraph()
    p2.text = players
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = C_NAVY
    p2.space_before = Pt(3)

fill_pos(s4, 1.0, 2.4, "High Grounding / Narrow Scope", "Traditional IVR & Rules Bots\n• Payer IVR (Nuance)\n• Static portal search / FAQs\n• Safe, but rigid & high drop-off", C_TEAL, C_CARD_BG)
fill_pos(s4, 4.3, 2.4, "★ TARGET: Carelon Assistant", "High Grounding / Broad Scope\n• Grounded RAG on SBC + Claims\n• Verified Plan Citations\n• Seamless Contextual Escalation", C_PURPLE, C_PURPLE_LIGHT)
fill_pos(s4, 1.0, 4.25, "Low Grounding / Narrow Scope", "Legacy Decision Trees (V1 Chat)\n• Rigid intent matching\n• Frequent 'I don't understand'\n• Forces immediate call escalation", C_MUTED, C_BG_LIGHT)
fill_pos(s4, 4.3, 4.25, "Low Grounding / Broad Scope", "Public LLMs & Symptom Checkers\n• ChatGPT, WebMD, Ada\n• No plan benefit grounding\n• Severe hallucination & PHI risk", C_RED, C_CARD_BG)

# Right: Competitive Observations
add_card(s4, 7.8, 1.8, 4.733, 4.4)
tb_comp_h = s4.shapes.add_textbox(Inches(8.0), Inches(1.95), Inches(4.3), Inches(4.1))
tfc = tb_comp_h.text_frame
tfc.word_wrap = True

pc1 = tfc.paragraphs[0]
pc1.text = "Competitive Assessment & Differentiation"
pc1.font.size = Pt(12)
pc1.font.bold = True
pc1.font.color.rgb = C_NAVY

c_obs = [
    ("Payer Peers (Cigna, UHC, Humana):", "Most deploy intent-tree bots that handle basic ID card retrieval but cannot parse complex EOBs or provide grounded benefit citations."),
    ("Big Tech Providers (Azure, AWS):", "Offer HIPAA-eligible cloud frameworks (Azure Health Bot), but lack out-of-the-box payer claims accumulator orchestration."),
    ("The Carelon Advantage:", "Direct integration with Elevance Health claims data and clinical governance, enabling the 'Show How It Works' audit trail."),
    ("⚠️ Market Verification:", "Verify competitor commercial feature sets before external executive presentations.")
]
for ch, cb in c_obs:
    p_h = tfc.add_paragraph()
    p_h.text = ch
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_PURPLE if "Advantage" in ch else C_TEAL
    p_h.space_before = Pt(5)
    p_b = tfc.add_paragraph()
    p_b.text = cb
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_NAVY
    p_b.space_before = Pt(1)

add_proto_callout(s4, "View 1: Member Experience", "Toggle 'Show how it works' under any response to see the retrieval source, guardrail checks, and confidence score.")
add_footer(s4, 4)
add_speaker_notes(s4,
    "The competitive positioning map shows why generic AI fails in healthcare. Public LLMs have conversational breadth but zero plan grounding and unacceptable hallucination rates. Legacy IVRs are safe but rigid. Our prototype demonstrates the target top-right quadrant: an assistant grounded in individual member plan documents that displays its exact confidence and retrieval source on every interaction.",
    "How does the assistant protect against proprietary plan documents leaking across employers?",
    "Every plan document is isolated by Tenant_ID and Group_Number in the vector index. The session token restricts retrieval to the member's authenticated Group ID, making cross-employer data leakage impossible.")

# ==============================================================================
# SLIDE 5: VISION, PRINCIPLES & METRICS (MATCHING PROTOTYPE KPIS)
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
apply_background(s5)
add_header(s5, "Vision, Principles & KPIs", "Product Principles Operationalized Across Six Live Telemetry KPI Cards in the Dashboard")

# Top Banner: 6 KPI Cards Matching View 2 Exactly
kpis_dash = [
    ("FCR Rate", "61.3%", "↑ +8.2pp vs last mo", C_GREEN),
    ("Escalation Rate", "14.7%", "↓ −3.1pp", C_GREEN),
    ("Member CSAT", "4.2/5.0", "↑ +0.4", C_GREEN),
    ("Avg Resolution Time", "2m 47s", "↓ −4m 13s vs phone", C_GREEN),
    ("Cost per Contact", "$1.84", "↓ −$8.16 vs phone", C_GREEN),
    ("Safety Escalation Recall", "99.4%", "→ Stable", C_AMBER)
]

for k_idx, (k_lbl, k_val, k_tr, k_col) in enumerate(kpis_dash):
    k_left = 0.8 + (k_idx * 1.98)
    card_k = add_card(s5, k_left, 1.8, 1.88, 1.5, bg_color=C_CARD_BG, border_color=C_BORDER)
    tb_k = s5.shapes.add_textbox(Inches(k_left + 0.1), Inches(1.88), Inches(1.68), Inches(1.3))
    tfk = tb_k.text_frame
    tfk.word_wrap = True
    pk1 = tfk.paragraphs[0]
    pk1.text = k_lbl
    pk1.font.size = Pt(9)
    pk1.font.bold = True
    pk1.font.color.rgb = C_MUTED
    pk2 = tfk.add_paragraph()
    pk2.text = k_val
    pk2.font.size = Pt(15)
    pk2.font.bold = True
    pk2.font.color.rgb = C_NAVY
    pk2.space_before = Pt(2)
    pk3 = tfk.add_paragraph()
    pk3.text = k_tr
    pk3.font.size = Pt(8)
    pk3.font.bold = True
    pk3.font.color.rgb = k_col
    pk3.space_before = Pt(2)

# Bottom Container: Principles & Guardrails
add_card(s5, 0.8, 3.45, 11.733, 2.75)
tb_pr_h = s5.shapes.add_textbox(Inches(1.0), Inches(3.55), Inches(11.3), Inches(0.35))
tb_pr_h.text_frame.paragraphs[0].text = "Five Core Product Principles Governing the Prototype"
tb_pr_h.text_frame.paragraphs[0].font.size = Pt(12)
tb_pr_h.text_frame.paragraphs[0].font.bold = True
tb_pr_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

princ_list = [
    ("1. Trust Over Completeness:", "If RAG confidence is <0.7, acknowledge uncertainty and warm-transfer. Never guess benefits."),
    ("2. 6th-Grade Plain Language:", "Default to accessible readability; offer 1-click 'Explain Simpler' button (drops to 4th grade)."),
    ("3. Zero Dead Ends:", "Every interaction provides an answer, a clear next step, or a warm handoff with full context preserved."),
    ("4. Algorithmic Equity:", "Maintain less than a 5pp FCR gap across Commercial, MA, and Medicaid populations."),
    ("5. Minimum Necessary PHI:", "Pull only the accumulator or benefit line needed for the specific intent; never expose full clinical records.")
]

tb_pr_body = s5.shapes.add_textbox(Inches(1.0), Inches(3.95), Inches(11.3), Inches(2.1))
tfpb = tb_pr_body.text_frame
tfpb.word_wrap = True
for pr_t, pr_d in princ_list:
    p_pr = tfpb.add_paragraph() if tfpb.paragraphs[0].text else tfpb.paragraphs[0]
    p_pr.text = f"• {pr_t} {pr_d}"
    p_pr.font.size = Pt(9.5)
    p_pr.font.color.rgb = C_NAVY
    p_pr.space_before = Pt(4)

add_proto_callout(s5, "View 2: Product Dashboard", "Inspect the top KPI row: FCR (61.3%), Cost per Contact ($1.84), and Safety Recall (99.4%) with trend arrows.")
add_footer(s5, 5)
add_speaker_notes(s5,
    "The six KPI cards on this slide directly mirror the top row of our Product Dashboard in View 2. Our North-Star Metric is First-Contact Resolution, currently performing at 61.3%—up 8.2 points. Notice our guardrails: Escalation Rate is 14.7%, Cost per Contact is down to $1.84, and Safety Escalation Recall is 99.4%. These metrics prove we aren't deflecting calls by frustrating members; we are resolving their issues on first contact.",
    "What causes the 14.7% escalation rate?",
    "Escalations represent healthy system behavior: low-confidence edge cases, acute clinical symptoms, and formal appeals. Our objective is not 0% escalation, which would indicate dangerous suppression; 12% to 15% is the optimal operational window.")

# ==============================================================================
# SLIDE 6: OPTIONS & PRIORITIZATION (MATCHING PROTOTYPE BACKLOG)
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
apply_background(s6)
add_header(s6, "Backlog & Prioritization (RICE)", "Prioritizing High-Volume, Low-Clinical-Risk Inquiries; Complex Cost Estimators Deferred")

# Backlog Table from View 3
t6_shape = s6.shapes.add_table(8, 7, Inches(0.8), Inches(1.8), Inches(8.5), Inches(4.4))
t6 = t6_shape.table
t6.columns[0].width = Inches(2.1)
t6.columns[1].width = Inches(0.9)
t6.columns[2].width = Inches(0.9)
t6.columns[3].width = Inches(1.1)
t6.columns[4].width = Inches(1.1)
t6.columns[5].width = Inches(1.1)
t6.columns[6].width = Inches(1.3)

b_hdrs = ["Epic / Story", "Reach", "Impact", "Confidence", "Effort", "RICE Score", "Phase"]
for c_idx, htext in enumerate(b_hdrs):
    cell = t6.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Exact data from View 3 of Prototype
dash_backlog = [
    ("Warm Handoff w/ Context", "5,000", "4", "90%", "2 wks", "9,000", "✅ MVP (P1)"),
    ("Safety Escalation Trigger", "5,000", "5", "95%", "3 wks", "7,917", "✅ MVP (P0)"),
    ("Deductible/OOP Tracker", "5,000", "4", "90%", "3 wks", "6,000", "✅ MVP (P1)"),
    ("Coverage & Benefits Q&A", "5,000", "4", "85%", "4 wks", "4,250", "✅ MVP (P1)"),
    ("Claims Status + EOB", "5,000", "4", "80%", "4 wks", "4,000", "✅ MVP (P1)"),
    ("Cost Estimator", "4,000", "4", "60%", "8 wks", "1,200", "⏸️ V2 (P3)"),
    ("Provider Search", "4,000", "3", "65%", "6 wks", "1,300", "⏸️ V2 (P3)")
]

for r_idx, r_vals in enumerate(dash_backlog, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t6.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        if c_idx == 5:
            p.font.bold = True
            p.font.color.rgb = C_PURPLE
        if c_idx == 6:
            p.font.bold = True
            p.font.color.rgb = C_GREEN if "MVP" in val else C_MUTED
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right: Non-Goals Rationale
add_card(s6, 9.5, 1.8, 3.033, 4.4)
tb_ng_h = s6.shapes.add_textbox(Inches(9.65), Inches(1.95), Inches(2.7), Inches(4.1))
tfng = tb_ng_h.text_frame
tfng.word_wrap = True

p_ng1 = tfng.paragraphs[0]
p_ng1.text = "Deliberate V1 Non-Goals"
p_ng1.font.size = Pt(12)
p_ng1.font.bold = True
p_ng1.font.color.rgb = C_NAVY

ng_items = [
    ("Cost Estimator (V2):", "RICE score 1,200. High effort (8 wks) due to contracted fee schedules and facility fee variations. Defer to protect pricing trust."),
    ("Provider Search (V2):", "RICE score 1,300. Requires real-time provider network directory synchronization."),
    ("Rx Formulary Lookup (V3):", "Requires real-time PBM API integration with therapeutic tiering logic."),
    ("No Clinical Advice:", "Hard non-goal across all versions; strictly administrative.")
]
for ngh, ngb in ng_items:
    p_h = tfng.add_paragraph()
    p_h.text = ngh
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_PURPLE
    p_h.space_before = Pt(5)
    p_b = tfng.add_paragraph()
    p_b.text = ngb
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_NAVY
    p_b.space_before = Pt(1)

add_proto_callout(s6, "View 3: Roadmap & Backlog", "Move the Effort slider on 'Coverage & Benefits Q&A' and click 'Recompute All' to see RICE scores recalculate live.")
add_footer(s6, 6)
add_speaker_notes(s6,
    "In View 3 of the prototype, you can interact directly with our RICE backlog. Notice how Warm Handoff with Context and Safety Escalation Trigger rise to the very top with RICE scores of 9,000 and 7,917. Conversely, Cost Estimator scores only 1,200 due to its 8-week effort and 60% confidence rating. We deliberately defer Cost Estimator to V2 to protect member trust from inaccurate fee quotes.",
    "Why does Warm Handoff score higher than core Coverage Q&A?",
    "Because warm handoff protects the entire customer service operation. If the AI fails and forces a cold transfer where the member must re-explain everything, CSAT plummets and handle time increases. Handoff is an operational safety net that enables all other AI features.")

# ==============================================================================
# SLIDE 7: SOLUTION WALKTHROUGH (THE PRODUCT IN THE DASHBOARD)
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
apply_background(s7)
add_header(s7, "Solution Walkthrough: Member Chat", "Five Scripted Scenarios with Reading Adaptation and RAG Transparency in View 1")

# 5 Scripted Scenarios Matching View 1
scenarios_flow = [
    ("Flow 1: MRI Coverage Q&A", "Chip: 'Check Coverage'\n• Verifies BlueStar Silver SBC Sec 4\n• Explains $3,000 deductible & 20% coinsurance\n• Estimates $240–$480 in-network cost\n• Highlights prior-authorization flag"),
    ("Flow 2: Denied Claim Explainer", "Chip: 'Claim Status'\n• Retrieves denied Claim #CLM-2026-8847\n• Explains denial: Out-of-network provider\n• Plain-English guidance on out-of-network rules\n• Outlines 90-day appeal timeline (Jan 14, 2027)"),
    ("Flow 3: Knee Surgery Cost Estimate", "Chip: 'Estimate Cost'\n• Interactive deductible progress bar (15% paid)\n• Displays $2,550 remaining deductible\n• Outpatient ($3,800–$5,200) vs inpatient ranges\n• Prompts for in-network orthopedic surgeon"),
    ("Flow 4: Clinical Safety Intercept", "Prompt: 'I am having chest pain'\n• 🚨 Immediate emergency override (0ms delay)\n• Urges call to 911 or nearest Emergency Room\n• Displays 988 Suicide & Crisis Lifeline info\n• Live Agent Maria T. auto-joins chat"),
    ("Flow 5: Low-Confidence Handshake", "Prompt: 'Prior auth exception'\n• Confidence flagged as Medium (<0.7)\n• States: 'I recommend confirming with a specialist'\n• Pre-compiles structured handoff payload\n• Transfers member with zero question repeat")
]

for idx, (s_title, s_desc) in enumerate(scenarios_flow):
    s_left = 0.8 + (idx * 2.38)
    card_s = add_card(s7, s_left, 1.8, 2.28, 4.4, bg_color=C_CARD_BG, border_color=C_BORDER)
    tb_s = s7.shapes.add_textbox(Inches(s_left + 0.1), Inches(1.9), Inches(2.08), Inches(4.2))
    tfs = tb_s.text_frame
    tfs.word_wrap = True
    ps1 = tfs.paragraphs[0]
    ps1.text = s_title
    ps1.font.size = Pt(10)
    ps1.font.bold = True
    ps1.font.color.rgb = C_RED if "Safety" in s_title else (C_PURPLE if "Low-Confidence" in s_title else C_TEAL)
    ps2 = tfs.add_paragraph()
    ps2.text = s_desc
    ps2.font.size = Pt(8.5)
    ps2.font.color.rgb = C_NAVY
    ps2.space_before = Pt(4)

add_proto_callout(s7, "View 1: Member Experience", "Click 'Check Coverage', 'Claim Status', 'Estimate Cost', or type 'chest pain' to trigger these exact flows live.")
add_footer(s7, 7)
add_speaker_notes(s7,
    "View 1 of our prototype demonstrates five distinct conversational flows. Flow 1 and Flow 2 show our core coverage and claims explainer. Flow 3 demonstrates how deductible accumulators integrate directly into cost estimates. Most critically, Flow 4 demonstrates our safety intercept: typing 'chest pain' immediately triggers an emergency directive and auto-connects agent Maria T. Flow 5 demonstrates our honest 'I don't know' protocol for low-confidence queries.",
    "What happens if a member clicks the 'Explain simpler' button on the MRI coverage answer?",
    "In the prototype, clicking 'Explain simpler' replaces the original response with an alternative synthesized at a 4th-grade reading level, stripping out terms like 'diagnostic imaging' and 'coinsurance' in favor of 'tests' and 'your share of the bill.'")

# ==============================================================================
# SLIDE 8: ARCHITECTURE & DATA (SHOW HOW IT WORKS)
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
apply_background(s8)
add_header(s8, "Business-Readable Architecture", "A 5-Tier Retrieval Pipeline Directly Inspected via the Prototype's 'Show How It Works' Panel")

# 5 Architecture Tiers
arch_tiers = [
    ("TIER 1: CHANNELS", "Web Portal Widget (Prototype View 1)  •  Mobile App Webview  •  Two-Way SMS  •  Future IVR (V2)", C_TEAL),
    ("TIER 2: SECURITY & AUTH", "SSO / MFA Token Validation  •  HIPAA Minimum Necessary Consent  •  Scoped User Session (15m timeout)", C_NAVY),
    ("TIER 3: ORCHESTRATION & SAFETY", "Semantic Intent Classifier  •  Deterministic Clinical Emergency Regex  •  PII Redaction Engine", C_RED),
    ("TIER 4: GROUNDED RAG RETRIEVAL", "Vector Database of SBCs & Plan Documents  •  Real-Time Claims Accumulator API (Deductible / OOP Max)", C_PURPLE),
    ("TIER 5: AUDITABLE RESPONSE", "Response Generator w/ Source Citation  •  Confidence Scorer (<0.7 threshold)  •  Immutable Audit Logger", C_GREEN)
]

for a_idx, (a_title, a_desc, a_col) in enumerate(arch_tiers):
    a_top = 1.8 + (a_idx * 0.88)
    card_a = add_card(s8, 0.8, a_top, 11.733, 0.78, bg_color=C_CARD_BG, border_color=a_col)
    tb_at = s8.shapes.add_textbox(Inches(1.0), Inches(a_top + 0.12), Inches(3.2), Inches(0.55))
    tb_at.text_frame.paragraphs[0].text = a_title
    tb_at.text_frame.paragraphs[0].font.size = Pt(10)
    tb_at.text_frame.paragraphs[0].font.bold = True
    tb_at.text_frame.paragraphs[0].font.color.rgb = a_col

    tb_ad = s8.shapes.add_textbox(Inches(4.4), Inches(a_top + 0.12), Inches(7.9), Inches(0.55))
    tb_ad.text_frame.paragraphs[0].text = a_desc
    tb_ad.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_ad.text_frame.paragraphs[0].font.color.rgb = C_NAVY

add_proto_callout(s8, "View 1: Member Experience", "Toggle 'Show how it works' under any chat message to reveal the live RAG source citation, guardrails passed, and confidence score (0.87).")
add_footer(s8, 8)
add_speaker_notes(s8,
    "This 5-tier architecture explains how our prototype prevents hallucinations. The key is Tier 4: the model cannot answer from generalized training memory. It synthesizes only from vectorized chunks of the member's verified Summary of Benefits. In View 1 of the prototype, clicking 'Show how it works' lets you inspect this in action: it displays the retrieved document section, confirms that PHI and clinical filters passed, and reveals the 0.87 confidence score.",
    "What vector database and chunking strategy is recommended for health plan SBCs?",
    "We recommend table-aware chunking where copay grids and exclusion schedules are stored as distinct relational nodes rather than arbitrary text chunks, stored in an enterprise HIPAA-compliant vector store like Azure AI Search or Pinecone.")

# ==============================================================================
# SLIDE 9: SAFETY, PRIVACY, COMPLIANCE & FAIRNESS (MATCHING EQUITY PANEL)
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
apply_background(s9)
add_header(s9, "Safety, Compliance & Equity", "Direct Telemetry from the Equity and Safety Panels in View 2 Highlights Proactive Governance")

# Left: Equity Monitor Table from View 2
add_card(s9, 0.8, 1.8, 7.2, 4.4)
tb_eq_h = s9.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.8), Inches(0.35))
tb_eq_h.text_frame.paragraphs[0].text = "Equity Monitor — FCR & CSAT by Member Segment (View 2)"
tb_eq_h.text_frame.paragraphs[0].font.size = Pt(11.5)
tb_eq_h.text_frame.paragraphs[0].font.bold = True
tb_eq_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

t9_shape = s9.shapes.add_table(8, 4, Inches(1.0), Inches(2.35), Inches(6.8), Inches(3.7))
t9 = t9_shape.table
t9.columns[0].width = Inches(2.4)
t9.columns[1].width = Inches(1.2)
t9.columns[2].width = Inches(1.1)
t9.columns[3].width = Inches(2.1)

e_hdrs = ["Member Segment", "FCR Rate", "CSAT", "Governance Status & Action"]
for c_idx, htext in enumerate(e_hdrs):
    cell = t9.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Exact data from Equity Panel in Prototype
eq_data = [
    ("Commercial members", "63%", "4.3", "✅ OK (Baseline)"),
    ("Medicare Advantage", "58%", "4.1", "⚠️ GAP (5pp below commercial)"),
    ("Medicaid", "56%", "3.9", "🚨 FLAG (7pp gap — exceeds 5pp threshold)"),
    ("English speakers", "63%", "4.3", "✅ OK (Baseline)"),
    ("Spanish speakers", "59%", "4.0", "⚠️ GAP (Review translation quality)"),
    ("Other languages", "54%", "3.8", "🚨 FLAG (Add plain-language templates)"),
    ("Low literacy proxy", "55%", "3.9", "🚨 FLAG (Activate literacy-adaptive mode)")
]

for r_idx, r_vals in enumerate(eq_data, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t9.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if "FLAG" in val:
            p.font.bold = True
            p.font.color.rgb = C_RED
        elif "GAP" in val:
            p.font.bold = True
            p.font.color.rgb = C_AMBER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right: Quality & Safety Panel from View 2
add_card(s9, 8.2, 1.8, 4.333, 4.4)
tb_qs_h = s9.shapes.add_textbox(Inches(8.4), Inches(1.95), Inches(3.9), Inches(4.1))
tfqs = tb_qs_h.text_frame
tfqs.word_wrap = True

p_qs1 = tfqs.paragraphs[0]
p_qs1.text = "Quality & Safety Panel (View 2)"
p_qs1.font.size = Pt(12)
p_qs1.font.bold = True
p_qs1.font.color.rgb = C_NAVY

qs_items = [
    ("Hallucination Rate (last 500):", "0.4% (✅ Well below <1.0% target). Audited weekly via transcript sampling."),
    ("Unanswered Question Queue:", "23 items pending review (8 items >48 hrs). Flags new unmapped benefit terms."),
    ("Safety Triggers Last 30 Days:", "47 total triggers; 47 escalated to human agents (100% safety escalation recall)."),
    ("Human Review Queue:", "3 pending verification cards (QR-001 to QR-003) with live Accept/Reject controls.")
]
for qsh, qsb in qs_items:
    p_h = tfqs.add_paragraph()
    p_h.text = qsh
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_PURPLE if "Hallucination" in qsh or "Review" in qsh else C_TEAL
    p_h.space_before = Pt(5)
    p_b = tfqs.add_paragraph()
    p_b.text = qsb
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_NAVY
    p_b.space_before = Pt(1)

add_proto_callout(s9, "View 2: Product Dashboard", "Scroll to 'Equity Monitor' and 'Quality & Safety Panel' to see the 🚨 7pp Medicaid flag and the Human Review Queue.")
add_footer(s9, 9)
add_speaker_notes(s9,
    "Here we see why equity is a launch criterion. In View 2 of our dashboard, the Equity Monitor automatically triggers a red flag if the FCR gap between any segment exceeds 5 percentage points. You can see Medicaid is at 56%—a 7-point gap against Commercial's 63%—which immediately surfaced a recommended action to deploy plain-language templates. Meanwhile, our Safety Panel shows zero missed safety triggers across 47 real-world events.",
    "What is the operational protocol when a segment hits a red flag in the Equity Monitor?",
    "A red flag blocks automated rollouts to additional employer groups in that tier. The Product Owner and Lead Content Strategist must conduct a 48-hour root cause analysis on the low-FCR transcripts, deploying simplified prompt templates and testing translation accuracy before scaling resumes.")

# ==============================================================================
# SLIDE 10: BUSINESS CASE & FINANCIAL SCENARIOS
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
apply_background(s10)
add_header(s10, "Business Case & Financial Scenarios", "A Rigorous ROI Model Anchored in the Prototype's $1.84 vs. $10.00 Cost-per-Contact")

# Left: 3 Financial Scenarios Table
add_card(s10, 0.8, 1.8, 7.3, 4.4)
tb_fin_h = s10.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.9), Inches(0.35))
tb_fin_h.text_frame.paragraphs[0].text = "3-Year Pro-Forma Scenarios (Illustrative Assumptions Aligned to Prototype)"
tb_fin_h.text_frame.paragraphs[0].font.size = Pt(11.5)
tb_fin_h.text_frame.paragraphs[0].font.bold = True
tb_fin_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

t10_shape = s10.shapes.add_table(6, 4, Inches(1.0), Inches(2.35), Inches(6.9), Inches(3.7))
t10 = t10_shape.table
t10.columns[0].width = Inches(2.5)
t10.columns[1].width = Inches(1.4)
t10.columns[2].width = Inches(1.5)
t10.columns[3].width = Inches(1.5)

f_hdrs = ["Financial Metric", "Conservative", "Base Case (Prototype)", "Aggressive"]
for c_idx, htext in enumerate(f_hdrs):
    cell = t10.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

f_data = [
    ("Net Call Deflection Rate", "10.0%", "20.0%", "30.0%"),
    ("Annual Calls Deflected", "200,000", "400,000", "600,000"),
    ("Cost Savings ($10 vs $1.84)", "$1,632,000", "$3,264,000", "$4,896,000"),
    ("Year 1 CapEx / Implementation", "$3,200,000", "$3,200,000", "$3,200,000"),
    ("3-Year Cumulative NPV (10% WACC)", "$1,850,000", "$5,240,000", "$8,620,000")
]

for r_idx, r_vals in enumerate(f_data, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t10.cell(r_idx, c_idx)
        cell.fill.solid()
        if r_idx == 5:
            cell.fill.fore_color.rgb = C_PURPLE_LIGHT
        else:
            cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0 or r_idx == 5:
            p.font.bold = True
        if r_idx == 5 and c_idx > 0:
            p.font.color.rgb = C_PURPLE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# Right: Cost-per-Contact Unit Economics
add_card(s10, 8.3, 1.8, 4.233, 4.4)
tb_ue_h = s10.shapes.add_textbox(Inches(8.5), Inches(1.95), Inches(3.8), Inches(4.1))
tfue = tb_ue_h.text_frame
tfue.word_wrap = True

p_ue1 = tfue.paragraphs[0]
p_ue1.text = "Unit Economic Breakdown"
p_ue1.font.size = Pt(12)
p_ue1.font.bold = True
p_ue1.font.color.rgb = C_NAVY

ue_points = [
    ("Baseline Phone Cost:", "$10.00 / call loaded agent cost (7m 00s handle time + loaded benefits)."),
    ("AI Assistant Cost:", "$1.84 / interaction (cloud compute, RAG vector indexing, amortized ops team)."),
    ("Net Unit Savings:", "$8.16 saved for every member question resolved via AI self-service."),
    ("Break-Even Milestone:", "Month 14 under Base Case; Month 22 under Conservative Case.")
]
for ueh, ueb in ue_points:
    p_h = tfue.add_paragraph()
    p_h.text = ueh
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_PURPLE if "Unit" in ueh else C_TEAL
    p_h.space_before = Pt(5)
    p_b = tfue.add_paragraph()
    p_b.text = ueb
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_NAVY
    p_b.space_before = Pt(1)

add_proto_callout(s10, "View 2: Product Dashboard", "Inspect the 'Cost per Contact' KPI card ($1.84 vs $10 phone) and the 'Call Deflection' before-and-after chart.")
add_footer(s10, 10)
add_speaker_notes(s10,
    "Our business case is derived directly from the unit economics proven in the dashboard: each live phone call costs $10.00, while an AI assistant interaction costs $1.84. That is an $8.16 net margin per deflected call. At 400,000 annual deflected calls, gross operational savings reach $3.26M annually. After subtracting $1.4M in ongoing operating expenses, the base case achieves break-even at month 14 and delivers a 3-year NPV of $5.24M.",
    "What if deflection only reaches 10% instead of 20%?",
    "As modeled in our Conservative Scenario on the left, even at 10% deflection (200,000 calls), net annual savings still reach $1.63M, yielding a positive 3-year NPV of $1.85M with break-even at month 22. The product remains financially viable even at half our performance target.")

# ==============================================================================
# SLIDE 11: METRICS, PILOT & SCALE GATES (MIRRORING DASHBOARD)
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
apply_background(s11)
add_header(s11, "Measurement, Pilot & Scale Gates", "Tracking the Resolution Funnel and 12-Week Trends with Rigorous Go / No-Go Decision Gates")

# Top Half: Funnel & 12-Week Trend from View 2
add_card(s11, 0.8, 1.8, 5.7, 2.3)
tb_fn_h = s11.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(5.3), Inches(0.35))
tb_fn_h.text_frame.paragraphs[0].text = "Resolution Funnel Telemetry (View 2)"
tb_fn_h.text_frame.paragraphs[0].font.size = Pt(11)
tb_fn_h.text_frame.paragraphs[0].font.bold = True
tb_fn_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

fn_text = "• Sessions Started: 100% → Intent Identified: 94%\n• Answer Provided: 81% → Member Confirmed Resolved: 61%\n• Escalated to Live Agent: 14% | Session Abandoned: 5%"
tb_fn_b = s11.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(5.3), Inches(1.7))
tb_fn_b.text_frame.paragraphs[0].text = fn_text
tb_fn_b.text_frame.paragraphs[0].font.size = Pt(9.5)
tb_fn_b.text_frame.paragraphs[0].font.color.rgb = C_NAVY

add_card(s11, 6.7, 1.8, 5.833, 2.3)
tb_tr_h = s11.shapes.add_textbox(Inches(6.9), Inches(1.9), Inches(5.4), Inches(0.35))
tb_tr_h.text_frame.paragraphs[0].text = "12-Week Trend: FCR vs Escalation (View 2)"
tb_tr_h.text_frame.paragraphs[0].font.size = Pt(11)
tb_tr_h.text_frame.paragraphs[0].font.bold = True
tb_tr_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

tr_text = "• FCR Rate climbed steadily from 42% (Wk 1) to 61.3% (Wk 12)\n• Escalation Rate dropped from 24% (Wk 1) to 14.7% (Wk 12)\n• Proves that RAG prompt tuning directly expands self-service capacity"
tb_tr_b = s11.shapes.add_textbox(Inches(6.9), Inches(2.25), Inches(5.4), Inches(1.7))
tb_tr_b.text_frame.paragraphs[0].text = tr_text
tb_tr_b.text_frame.paragraphs[0].font.size = Pt(9.5)
tb_tr_b.text_frame.paragraphs[0].font.color.rgb = C_NAVY

# Bottom Half: Decision Gates
add_card(s11, 0.8, 4.25, 5.7, 2.05, bg_color=C_TEAL_LIGHT, border_color=C_TEAL)
tb_sc_h = s11.shapes.add_textbox(Inches(1.0), Inches(4.35), Inches(5.3), Inches(0.35))
tb_sc_h.text_frame.paragraphs[0].text = "🟢 Scale Decision Gates (Proceed to V2)"
tb_sc_h.text_frame.paragraphs[0].font.size = Pt(11)
tb_sc_h.text_frame.paragraphs[0].font.bold = True
tb_sc_h.text_frame.paragraphs[0].font.color.rgb = C_TEAL

sc_items = [
    "• FCR Rate >= 60.0% (Prototype achieving 61.3% ✅)",
    "• Safety Escalation Recall >= 99.0% (Achieving 99.4% ✅)",
    "• Cross-Segment Equity Gap <= 5.0 pp (Commercial vs Medicaid)",
    "• Cost per Contact <= $2.50 (Achieving $1.84 ✅)"
]
tb_sc_b = s11.shapes.add_textbox(Inches(1.0), Inches(4.7), Inches(5.3), Inches(1.5))
tfsc = tb_sc_b.text_frame
for sc in sc_items:
    p = tfsc.add_paragraph() if tfsc.paragraphs[0].text else tfsc.paragraphs[0]
    p.text = sc
    p.font.size = Pt(9)
    p.font.color.rgb = C_NAVY
    p.space_before = Pt(2)

add_card(s11, 6.7, 4.25, 5.833, 2.05, bg_color=RGBColor(254, 242, 242), border_color=C_RED)
tb_kl_h = s11.shapes.add_textbox(Inches(6.9), Inches(4.35), Inches(5.4), Inches(0.35))
tb_kl_h.text_frame.paragraphs[0].text = "🔴 Kill / Pivot Criteria (Pause Rollout)"
tb_kl_h.text_frame.paragraphs[0].font.size = Pt(11)
tb_kl_h.text_frame.paragraphs[0].font.bold = True
tb_kl_h.text_frame.paragraphs[0].font.color.rgb = C_RED

kl_items = [
    "• Safety recall drops below 95.0% (Immediate hard operational stop)",
    "• Hallucination rate on benefits exceeds 1.5% (Prototype at 0.4%)",
    "• Equity gap exceeds 8.0 pp for two consecutive measurement periods",
    "• Live agent handle time (AHT) increases by >10% due to poor handoffs"
]
tb_kl_b = s11.shapes.add_textbox(Inches(6.9), Inches(4.7), Inches(5.4), Inches(1.5))
tfkl = tb_kl_b.text_frame
for kl in kl_items:
    p = tfkl.add_paragraph() if tfkl.paragraphs[0].text else tfkl.paragraphs[0]
    p.text = kl
    p.font.size = Pt(9)
    p.font.color.rgb = C_NAVY
    p.space_before = Pt(2)

add_proto_callout(s11, "View 2: Product Dashboard", "Inspect the 'Resolution Funnel' and the '12-Week Trend — FCR & Escalation Rate' SVG charts.")
add_footer(s11, 11)
add_speaker_notes(s11,
    "In View 2, our Resolution Funnel chart shows that 100% of started sessions translate into 94% identified intent and 61% member-confirmed resolution, with 14% escalating cleanly to agents. Our 12-week trend shows FCR climbing from 42% to 61.3% as prompt embeddings were tuned. On the bottom, notice our hard decision gates: we scale only if FCR stays above 60% and safety recall above 99%. If safety recall drops below 95%, we halt.",
    "Why does the Resolution Funnel drop from 81% answered to 61% confirmed resolved?",
    "That 20% difference represents members who got an answer but either needed additional verification, chose to escalate, or closed the app without clicking a feedback button. We count only verified, completed interactions in FCR to ensure data integrity.")

# ==============================================================================
# SLIDE 12: AGILE EXECUTION (PRD, STORIES, PI PLAN)
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
apply_background(s12)
add_header(s12, "Agile Execution & Sprint Plan", "Delivering MVP in 16 Weeks Across Three Epics Governed by Rigid Acceptance Criteria")

# Left: 5 User Stories from Prototype
add_card(s12, 0.8, 1.8, 7.3, 4.4)
tb_st_h = s12.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.9), Inches(0.35))
tb_st_h.text_frame.paragraphs[0].text = "Core User Stories with Given/When/Then Acceptance Criteria"
tb_st_h.text_frame.paragraphs[0].font.size = Pt(11.5)
tb_st_h.text_frame.paragraphs[0].font.bold = True
tb_st_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

stories_dash = [
    ("US-01: Coverage Q&A (Core Engine)", "Given an authenticated member asks 'Is my MRI covered?', When the RAG engine queries Section 4 with confidence >0.7, Then render plain-language copay and source citation in <3 seconds."),
    ("US-02: Safety & Emergency Intercept (Safety P0)", "Given a member types 'I am having chest pain', When the safety classifier intercepts the input, Then bypass LLM, render red 911 directive, and connect live agent in <1.5 seconds."),
    ("US-03: Real-Time Deductible Tracker (Account)", "Given an HDHP member asks 'How much deductible is left?', When accumulator API responds, Then display progress bar showing $450 paid of $3,000 with remaining balance."),
    ("US-04: Multilingual Parity (Equity P2)", "Given a Spanish-speaking Medicaid member submits a query in Spanish, When the assistant responds, Then output native Spanish with equal factual accuracy (<5pp FCR gap)."),
    ("US-05: Contextual Agent Handoff (Ops P1)", "Given the assistant initiates human escalation, When the live rep accepts the session, Then the CRM screen displays member ID, plan, question, and partial transcript.")
]

tb_st_b = s12.shapes.add_textbox(Inches(1.0), Inches(2.35), Inches(6.9), Inches(3.7))
tfst = tb_st_b.text_frame
tfst.word_wrap = True
for st_t, st_c in stories_dash:
    p_t = tfst.add_paragraph() if tfst.paragraphs[0].text else tfst.paragraphs[0]
    p_t.text = st_t
    p_t.font.size = Pt(9.5)
    p_t.font.bold = True
    p_t.font.color.rgb = C_PURPLE if "Safety" in st_t or "Equity" in st_t else C_TEAL
    p_t.space_before = Pt(3)
    p_c = tfst.add_paragraph()
    p_c.text = st_c
    p_c.font.size = Pt(8.5)
    p_c.font.color.rgb = C_NAVY
    p_c.space_before = Pt(1)

# Right: Sprint Plan matching View 3 exactly
add_card(s12, 8.3, 1.8, 4.233, 4.4)
tb_sp_h = s12.shapes.add_textbox(Inches(8.5), Inches(1.95), Inches(3.8), Inches(4.1))
tfsp = tb_sp_h.text_frame
tfsp.word_wrap = True

p_sp1 = tfsp.paragraphs[0]
p_sp1.text = "Sprint Plan — Program Increment (View 3)"
p_sp1.font.size = Pt(11.5)
p_sp1.font.bold = True
p_sp1.font.color.rgb = C_NAVY

sprints_dash = [
    ("S1–S2 (Wk 1–4): Foundation", "Platform, Auth, PHI compliance setup  [✅ Done]"),
    ("S3–S4 (Wk 5–8): RAG Pipeline", "Coverage Q&A RAG, Confidence scoring  [✅ Done]"),
    ("S5–S6 (Wk 9–12): Integration", "Claims/EOB, Deductible tracker  [🔄 In Progress]"),
    ("S7–S8 (Wk 13–16): Safety & Handoff", "Safety classifier, Escalation, Handoff  [⬜ Planned]"),
    ("S9–S10 (Wk 17–20): Pilot Launch", "Spanish language, Equity testing  [⬜ Planned]"),
    ("S11–S12 (Wk 21–24): Evaluation", "Pilot A/B testing, Monitoring dashboard  [⬜ Planned]"),
    ("S13–S14 (Wk 25–28): Hardening", "Pilot analysis, Bug fixes, Scale Decision  [⬜ Planned]")
]
for sph, spb in sprints_dash:
    p_h = tfsp.add_paragraph()
    p_h.text = sph
    p_h.font.size = Pt(9)
    p_h.font.bold = True
    p_h.font.color.rgb = C_NAVY
    p_h.space_before = Pt(4)
    p_b = tfsp.add_paragraph()
    p_b.text = spb
    p_b.font.size = Pt(8)
    p_b.font.color.rgb = C_MUTED
    p_b.space_before = Pt(1)

add_proto_callout(s12, "View 3: Roadmap & Backlog", "Review the 'Sprint Plan — Program Increment' table showing S1–S2 through S13–S14 statuses.")
add_footer(s12, 12)
add_speaker_notes(s12,
    "Here is our Agile execution plan. On the left are five sample user stories with Given/When/Then acceptance criteria. Notice that Story 2 (Safety Intercept) and Story 4 (Spanish Parity) are treated as first-class backlog tickets. On the right is our 14-sprint Program Increment plan, directly mirroring the table in View 3: S1 through S4 are complete, S5-S6 is currently in progress integrating claims accumulators, leading to pilot launch at Sprint 9.",
    "How do you handle scope creep if stakeholders demand real-time provider directories during Sprint 5?",
    "Under SAFe, Sprint 5 scope is locked to protect accumulator API stabilization. Provider search is already estimated in our backlog as a 6-week V2 feature (RICE score 1,300). I would demonstrate the RICE score trade-off and maintain sprint integrity.")

# ==============================================================================
# SLIDE 13: LIVE DEMO PATH, GOVERNANCE & 90-DAY PLAN
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
apply_background(s13)
add_header(s13, "Live Demo, Governance & 90 Days", "A Structured 3-Minute Demo Path and Five Critical Decisions Needed from Leadership")

# Left: 3-Minute Live Demo Click Path
add_card(s13, 0.8, 1.8, 6.8, 4.4)
tb_dm_h = s13.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.4), Inches(0.35))
tb_dm_h.text_frame.paragraphs[0].text = "3-Minute Live Interview Demo Script (Click Path)"
tb_dm_h.text_frame.paragraphs[0].font.size = Pt(11.5)
tb_dm_h.text_frame.paragraphs[0].font.bold = True
tb_dm_h.text_frame.paragraphs[0].font.color.rgb = C_NAVY

demo_steps_text = [
    ("0:00–0:45 | View 1: Member Experience", "• Click 'Alex' persona → point to plan details sidebar ($3k ded, $450 paid).\n• Click 'Check Coverage' → walk through MRI answer & citation.\n• Toggle 'Show how it works' → prove RAG retrieval & 0.87 confidence.\n• Type 'chest pain' → prove immediate 0ms safety escalation to live nurse."),
    ("0:45–1:30 | View 2: Product Dashboard", "• Point to FCR (61.3%) and Cost per Contact ($1.84 vs $10 live call).\n• Highlight Equity Monitor → show 7pp Medicaid flag and recommended action.\n• Highlight Safety Panel → show 100% safety recall (47/47 triggers escalated)."),
    ("1:30–2:15 | View 3: Roadmap & Backlog", "• Drag Effort slider on 'Coverage Q&A' → click 'Recompute All' to show RICE math.\n• Point to Sprint Table → show S1-S2 Done, S5-S6 In Progress, S9 Pilot Launch."),
    ("2:15–3:00 | View 4: Care Management", "• Show member M-00291 (discharged 18 hrs ago, top priority).\n• Click 'Override ↓' → demonstrate clinician reason-capture modal.")
]

tb_dm_b = s13.shapes.add_textbox(Inches(1.0), Inches(2.35), Inches(6.4), Inches(3.7))
tfdm = tb_dm_b.text_frame
tfdm.word_wrap = True
for dm_t, dm_b in demo_steps_text:
    p_t = tfdm.add_paragraph() if tfdm.paragraphs[0].text else tfdm.paragraphs[0]
    p_t.text = dm_t
    p_t.font.size = Pt(9.5)
    p_t.font.bold = True
    p_t.font.color.rgb = C_PURPLE
    p_t.space_before = Pt(3)
    p_b = tfdm.add_paragraph()
    p_b.text = dm_b
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_NAVY
    p_b.space_before = Pt(1)

# Right: 5 Decisions Needed
add_card(s13, 7.8, 1.8, 4.733, 4.4, bg_color=C_PURPLE_LIGHT, border_color=C_PURPLE)
tb_dc_h = s13.shapes.add_textbox(Inches(8.0), Inches(1.95), Inches(4.3), Inches(4.1))
tfdc = tb_dc_h.text_frame
tfdc.word_wrap = True

p_dc1 = tfdc.paragraphs[0]
p_dc1.text = "5 Decisions Needed from Leadership"
p_dc1.font.size = Pt(12)
p_dc1.font.bold = True
p_dc1.font.color.rgb = C_PURPLE

decs_dash = [
    ("1. Authorize Y1 Implementation CapEx:", "$3.2M budget authorization across engineering, RAG setup, and compliance."),
    ("2. Approve Vendor BAA & Data Lifecycle:", "Legal sign-off on enterprise zero-retention LLM BAA and 90-day transcript purge."),
    ("3. Approve 90-Day Pilot Cohort:", "Selection of 1 Commercial and 1 Medicaid market (100k members total)."),
    ("4. Contact Center Operational Alignment:", "Authorize CRM integration for structured handoff payloads to live agents."),
    ("5. Form Clinical AI Governance Board:", "Charter joint review committee (Clinical, Legal, PM, Ops) for weekly audits.")
]
for dch, dcb in decs_dash:
    p_h = tfdc.add_paragraph()
    p_h.text = dch
    p_h.font.size = Pt(9.5)
    p_h.font.bold = True
    p_h.font.color.rgb = C_NAVY
    p_h.space_before = Pt(4)
    p_b = tfdc.add_paragraph()
    p_b.text = dcb
    p_b.font.size = Pt(8.5)
    p_b.font.color.rgb = C_MUTED
    p_b.space_before = Pt(1)

add_proto_callout(s13, "Prototype Demo Guide", "Click the floating '⏱ Demo Guide' button in the prototype top-right corner to open this exact 3-minute walkthrough script.")
add_footer(s13, 13)
add_speaker_notes(s13,
    "To conclude, this slide provides our exact 3-minute live interview demo script. In 45 seconds we demonstrate member self-service and safety escalation in View 1; in the next 45 seconds we show the FCR and Equity monitors in View 2; at minute 1:30 we show dynamic RICE recalculations in View 3; and at minute 2:15 we show clinical overrides in View 4. On the right are the five specific decisions I need from leadership to deliver our 90-day pilot.",
    "If leadership approves only $2.0M of the requested $3.2M CapEx, how do you adapt?",
    "Under a reduced $2.0M budget, I would preserve the core RAG engine, claims accumulator feeds, and clinical safety classifiers. We would defer the custom CRM handoff integration and utilize standard email/ticket generation for escalations in V1, ensuring data accuracy and safety are not compromised.")

# ==============================================================================
# APPENDIX SLIDES (14 - 18)
# ==============================================================================
# Slide 14: Appendix A - Assumptions Register
s14 = prs.slides.add_slide(blank_layout)
apply_background(s14)
add_header(s14, "Appendix A: Assumptions Register", "Full Transparency on Baseline Data, Sizing Formulas, and Operational Benchmarks")

t14_shape = s14.shapes.add_table(8, 4, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.9))
t14 = t14_shape.table
t14.columns[0].width = Inches(2.8)
t14.columns[1].width = Inches(1.8)
t14.columns[2].width = Inches(4.333)
t14.columns[3].width = Inches(2.8)

for c_idx, htext in enumerate(["Parameter / Metric", "Assumed Value", "Operational Logic & Source", "Sensitivity Impact"]):
    cell = t14.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

assump_list = [
    ("Total Annual Inbound Calls", "5,000,000 calls / yr", "Illustrative mid-size US health plan book of business.", "+/- 1M calls = +/- $800k annual benefit."),
    ("Cost per Live Inbound Call", "$10.00 / call", "Industry loaded benchmark: $8.00–$12.00 / call.", "+/- $2/call = +/- $800k gross savings."),
    ("Cost per AI Assistant Interaction", "$1.84 / interaction", "Derived directly from prototype telemetry KPI.", "+/- $0.50/interaction = +/- $200k OpEx."),
    ("Target Call Deflection Rate", "20.0% of automatable", "Derived from best-in-class healthcare AI implementations.", "+/- 5pp deflection = +/- $1.0M annual savings."),
    ("Year 1 Implementation CapEx", "$3,200,000", "Engineering, RAG pipeline, security/compliance audits.", "Fixed upfront investment amortized over 3 yrs."),
    ("Ongoing Annual OpEx", "$1,400,000 / yr", "Token compute, cloud hosting, 4 FTE review team.", "Scales slightly with token inference volume."),
    ("Equity Disparity Threshold", "Max 5.0 pp gap", "Aligned with NCQA HEDIS health equity standards.", "Strict operational release gate for all cohorts.")
]
for r_idx, r_vals in enumerate(assump_list, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t14.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s14, 14)
add_speaker_notes(s14, "Appendix A documents every numerical assumption, its operational basis, and its sensitivity impact. All metrics align with the dashboard.", "How sensitive is the business case to changes in live call costs?", "Every $1 change in cost-per-call impacts gross savings by $400,000 under our base case.")

# Slide 15: Appendix B - Glossary
s15 = prs.slides.add_slide(blank_layout)
apply_background(s15)
add_header(s15, "Appendix B: US Healthcare Glossary", "Standard Definitions for Payer Operations, Regulatory Rules, and AI Terminology")

gloss_items = [
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
add_card(s15, 0.8, 1.8, 11.733, 4.9)
tb_gl = s15.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.3), Inches(4.5))
tfgl = tb_gl.text_frame
tfgl.word_wrap = True
for term, definition in gloss_items:
    p = tfgl.add_paragraph() if tfgl.paragraphs[0].text else tfgl.paragraphs[0]
    p.text = f"{term} {definition}"
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_NAVY
    p.space_before = Pt(5)

add_footer(s15, 15)
add_speaker_notes(s15, "Appendix B provides standard US healthcare terminology reference definitions for all business and technical terms used in the presentation.", "Why is the distinction between copay and coinsurance so critical in assistant design?", "Over 60% of member billing complaints stem from members confusing flat copays with percentage-based coinsurance after deductible.")

# Slide 16: Appendix C - Detailed Risk Register
s16 = prs.slides.add_slide(blank_layout)
apply_background(s16)
add_header(s16, "Appendix C: Detailed Risk Register", "Comprehensive Risk Assessment Across Clinical, Legal, Operations, and Technical Domains")

t16_shape = s16.shapes.add_table(7, 5, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.9))
t16 = t16_shape.table
t16.columns[0].width = Inches(2.2)
t16.columns[1].width = Inches(1.2)
t16.columns[2].width = Inches(1.2)
t16.columns[3].width = Inches(4.933)
t16.columns[4].width = Inches(2.2)

for c_idx, htext in enumerate(["Risk Description", "Likelihood", "Impact", "Planned Mitigation Strategy", "Accountable Owner"]):
    cell = t16.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

risks_dash = [
    ("Hallucination on covered benefits", "Medium", "High", "RAG restriction to verified SBC text; confidence score threshold (<0.7 escalates).", "Data Science / PM"),
    ("Safety classifier misses acute distress", "Low", "Critical", "Dual-stage regex + NLP safety classifier; monthly 5k-prompt red-team audits.", "Clinical Quality / Lead"),
    ("State insurance AI regulatory shifts", "Medium", "High", "Active monitoring of NAIC model AI bulletin; legal review before new state rollouts.", "Legal & Compliance"),
    ("Call-center agent resistance / fear", "Medium", "Medium", "Change mgmt: AI pitched as agent helper; agents review drafts; time-savings shown.", "Call Center Ops"),
    ("Disparate impact across languages", "Medium", "High", "Dedicated Spanish evaluation suite; strict 5pp FCR parity threshold gate.", "Product Owner (Me)"),
    ("Claims accumulator API downtime", "Low", "High", "Circuit-breaker pattern: display cached accumulator with timestamp & fallback note.", "IT / Engineering")
]
for r_idx, r_vals in enumerate(risks_dash, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t16.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s16, 16)
add_speaker_notes(s16, "Appendix C details our risk register. The highest catastrophic risk is a clinical safety miss, which is mitigated via dual-stage classifiers and monthly red-teaming.", "Who holds the final sign-off authority on safety thresholds?", "The Clinical Quality and Chief Medical Officer representatives hold hard veto power over safety gates.")

# Slide 17: Appendix D - Extended Stories & NFRs
s17 = prs.slides.add_slide(blank_layout)
apply_background(s17)
add_header(s17, "Appendix D: Extended User Stories & NFRs", "Non-Functional Requirements and Secondary Backlog Stories for V2 Enablement")

ext_stories = [
    ("US-06: EOB Denial Plain-Language Explainer (V1)", "Given a member receives a denied claim notice, When they ask 'Why did my insurance not pay for Dr. Smith?', Then the assistant explains the specific denial reason code (e.g. out-of-network provider) in 6th-grade language and outlines the 90-day appeal timeline."),
    ("US-07: Low-Confidence Honest Handshake (V1 Safety)", "Given a member asks an ambiguous question where RAG retrieval confidence is <0.7, When the assistant responds, Then it states 'I cannot verify this specific exception in your plan documents' and offers a 1-click warm transfer."),
    ("NFR-01: Response Latency", "95% of standard conversational Q&A queries must return the complete synthesized response within 2.5 seconds under peak concurrency of 5,000 active sessions."),
    ("NFR-02: Accessibility Standards (WCAG 2.1 AA)", "All digital assistant interfaces must maintain a minimum 4.5:1 color contrast ratio, full keyboard navigability, and full compatibility with screen-readers (JAWS / NVDA)."),
    ("NFR-03: Security & Session Isolation", "Each member chat session must be strictly isolated in memory; session memory is flushed immediately upon logout or after 15 minutes of inactivity.")
]
add_card(s17, 0.8, 1.8, 11.733, 4.9)
tb_s17 = s17.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.3), Inches(4.5))
tfs17 = tb_s17.text_frame
tfs17.word_wrap = True
for title_e, body_e in ext_stories:
    p1 = tfs17.add_paragraph() if tfs17.paragraphs[0].text else tfs17.paragraphs[0]
    p1.text = title_e
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = C_PURPLE if "NFR" in title_e else C_TEAL
    p1.space_before = Pt(5)
    p2 = tfs17.add_paragraph()
    p2.text = body_e
    p2.font.size = Pt(9)
    p2.font.color.rgb = C_NAVY
    p2.space_before = Pt(1)

add_footer(s17, 17)
add_speaker_notes(s17, "Appendix D outlines secondary backlog stories and critical non-functional requirements including 2.5s response latency, WCAG 2.1 AA compliance, and strict memory isolation.", "Why is 2.5-second latency critical in conversational RAG?", "In digital chat, latency exceeding 3 seconds causes perceived friction and a 20% increase in premature member bounce or channel switching.")

# Slide 18: Appendix E - Prototype Data Dictionary & State Schema
s18 = prs.slides.add_slide(blank_layout)
apply_background(s18)
add_header(s18, "Appendix E: Prototype Data Dictionary", "Technical Mapping of Prototype State Variables, Filters, and Interactive Elements")

t18_shape = s18.shapes.add_table(7, 4, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.9))
t18 = t18_shape.table
t18.columns[0].width = Inches(2.2)
t18.columns[1].width = Inches(2.8)
t18.columns[2].width = Inches(4.533)
t18.columns[3].width = Inches(2.2)

for c_idx, htext in enumerate(["Dashboard State Variable", "Prototype Element ID", "Synthetic Data Values & Triggers", "Target View"]):
    cell = t18.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_TEAL
    p = cell.text_frame.paragraphs[0]
    p.text = htext
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

schema_rows = [
    ("currentPersona", "persona-pills .persona-pill", "alex | maria | david | eleanor | priya", "View 1 (Member Chat)"),
    ("chatFlowTrigger", "action-chip, chat-input", "coverage | claim | cost | chest pain | exception", "View 1 (Member Chat)"),
    ("showHowItWorks", "toggle-rag", "retrievedDoc, guardrailsPassed, confidenceScore (0.87)", "View 1 (Member Chat)"),
    ("filterState", "date-filter, plan-filter", "days: 30/60/90; plan: all/commercial/medicare/medicaid", "View 2 (Dashboard)"),
    ("riceRecompute", "rice-table .rice-slider", "Reach × Impact × Confidence% ÷ Effort", "View 3 (Roadmap)"),
    ("careCapacity", "capacity-slider", "1 to 12 members (Priority P1 to P3; Discharged, ED, Gap)", "View 4 (Care Mgmt)")
]
for r_idx, r_vals in enumerate(schema_rows, start=1):
    for c_idx, val in enumerate(r_vals):
        cell = t18.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_CARD_BG if r_idx % 2 == 1 else C_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_NAVY
        if c_idx == 0:
            p.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

add_footer(s18, 18)
add_speaker_notes(s18, "Appendix E provides a complete technical data dictionary mapping every interactive JavaScript state variable in our prototype HTML file to its respective dashboard view.", "How is the prototype packaged?", "It is delivered as a 100% self-contained single HTML file using pure vanilla JavaScript, CSS custom properties, and inline SVG charts, requiring zero external internet or CDN dependencies.")

# Save presentation
out_path = r"D:\Preksha_carelon\Carelon_Prototype_Aligned_Deck.pptx"
prs.save(out_path)
print(f"SUCCESS: Saved 18-slide prototype-aligned presentation to {out_path}")
