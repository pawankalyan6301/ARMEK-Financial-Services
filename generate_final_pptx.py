"""
generate_final_pptx.py
Professional review, correction and finalization of ARMEK PowerPoint Presentation
for CSE Final-Year Project Review at Kuppam Engineering College.
Creates ARMEK_Final_Project_Review.pptx preserving the Gamma fintech theme.
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==========================================
# CONSTANTS & THEME PALETTE
# ==========================================
FONT_HEADING = "Crimson Pro SemiBold"
FONT_BODY = "Heebo"
FONT_BOLD = "Heebo Bold"

COLOR_NAVY_DARK = RGBColor(0, 51, 102)     # #003366 - Deep academic navy
COLOR_NAVY_TITLE = RGBColor(21, 45, 71)    # #152D47 - Slide title navy
COLOR_BLUE_ACCENT = RGBColor(0, 102, 204)  # #0066CC - Fintech blue
COLOR_TEXT_DARK = RGBColor(76, 76, 77)     # #4C4C4D - Slate body text
COLOR_WHITE = RGBColor(255, 255, 255)      # #FFFFFF - Clean white

COLOR_CARD_BG = RGBColor(255, 255, 255)    # Card white
COLOR_CARD_BORDER = RGBColor(226, 232, 240)# Card subtle slate border
COLOR_ALT_ROW = RGBColor(248, 250, 252)    # Subtle table alternating row
COLOR_ACCENT_BG = RGBColor(240, 247, 255)  # Light blue tint box

COLOR_GREEN_BG = RGBColor(236, 253, 245)   # Light emerald green
COLOR_GREEN_BORDER = RGBColor(16, 185, 129)# Emerald border
COLOR_GREEN_TEXT = RGBColor(13, 104, 50)   # Dark emerald text

COLOR_RED_BG = RGBColor(254, 242, 242)     # Light rose red
COLOR_RED_BORDER = RGBColor(239, 68, 68)   # Rose red border
COLOR_RED_TEXT = RGBColor(185, 28, 28)     # Dark red text

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def remove_shapes_below(slide, y_threshold_inches):
    """Removes all shapes whose top position is greater than or equal to y_threshold_inches."""
    shapes_to_remove = [sh for sh in slide.shapes if sh.top.inches >= y_threshold_inches]
    for sh in shapes_to_remove:
        sp = sh._element
        sp.getparent().remove(sp)

def remove_all_shapes(slide):
    """Removes all shapes from a slide (for complete slide rebuilds)."""
    for sh in list(slide.shapes):
        sp = sh._element
        sp.getparent().remove(sp)

def add_header_badge(slide, text, top_inches=0.55):
    """Creates a consistent Gamma-style category badge at top left."""
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(top_inches), Inches(2.4), Inches(0.24))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(241, 245, 249)
    badge.line.color.rgb = RGBColor(203, 213, 225)
    badge.line.width = Pt(0.75)
    
    tf = badge.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.10)
    tf.margin_right = Inches(0.10)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)
    
    p = tf.paragraphs[0]
    p.text = text.upper()
    p.font.name = FONT_BOLD
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_DARK
    p.alignment = PP_ALIGN.CENTER

def add_slide_header(slide, badge_text, title_text, subtitle_text, top_offset=0.52):
    """Creates a standard Gamma slide header with category badge, title and subtitle."""
    # Badge
    add_header_badge(slide, badge_text, top_inches=top_offset)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.72), Inches(top_offset + 0.30), Inches(11.89), Inches(0.48))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = Inches(0)
    tf_t.margin_right = Inches(0)
    tf_t.margin_top = Inches(0)
    tf_t.margin_bottom = Inches(0)
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(25)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_NAVY_TITLE
    
    # Subtitle
    if subtitle_text:
        sub_box = slide.shapes.add_textbox(Inches(0.72), Inches(top_offset + 0.80), Inches(11.89), Inches(0.45))
        tf_s = sub_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = Inches(0)
        tf_s.margin_right = Inches(0)
        tf_s.margin_top = Inches(0)
        tf_s.margin_bottom = Inches(0)
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle_text
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(11.5)
        p_s.font.color.rgb = COLOR_TEXT_DARK

def format_cell(cell, text, bold=False, font_size=10, text_color=COLOR_TEXT_DARK, 
                bg_color=None, align=PP_ALIGN.LEFT, font_name=FONT_BODY):
    """Applies clean styling and padding to a table cell."""
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_WHITE
        
    cell.margin_left = Inches(0.12)
    cell.margin_right = Inches(0.12)
    cell.margin_top = Inches(0.08)
    cell.margin_bottom = Inches(0.08)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    tf = cell.text_frame
    tf.word_wrap = True
    
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
            p.space_before = Pt(2)
            
        p.text = line
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = text_color
        p.alignment = align

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER, border_width=1.0):
    """Creates a stylized card container with border."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)
    else:
        card.line.fill.background()
    return card

def add_flow_step(slide, text, left, top, width, height, step_num=None, is_highlighted=False):
    """Creates an individual flow step card with optional numbering."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    if is_highlighted:
        card.fill.fore_color.rgb = COLOR_NAVY_DARK
        card.line.color.rgb = COLOR_BLUE_ACCENT
        card.line.width = Pt(2.0)
    else:
        card.fill.fore_color.rgb = COLOR_ACCENT_BG
        card.line.color.rgb = RGBColor(191, 219, 254)
        card.line.width = Pt(1.0)
        
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    if step_num:
        p.text = f"{step_num}. {text}"
    else:
        p.text = text
    p.font.name = FONT_BOLD
    p.font.size = Pt(10 if not is_highlighted else 11)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE if is_highlighted else COLOR_NAVY_DARK

def add_arrow_label(slide, left, top, width, height, text="→"):
    """Adds a connecting arrow between steps."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_BOLD
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT

# ==========================================
# MAIN EXECUTION
# ==========================================

def main():
    input_pptx = "ARMEK-Agentic-AI-Powered-Loan-Underwriting-and-Risk-Assessment-System (1).pptx"
    output_pptx = "ARMEK_Final_Project_Review.pptx"
    
    print(f"Loading template: {input_pptx}")
    prs = pptx.Presentation(input_pptx)
    print(f"Loaded successfully. Total slides: {len(prs.slides)}")

    # ----------------------------------------------------
    # SLIDE 1 – TITLE AND ABSTRACT
    # ----------------------------------------------------
    print("Updating Slide 1: Title and Abstract...")
    s1 = prs.slides[0]
    
    # Exact required abstract text
    abstract_text = (
        "ARMEK is an AI-powered loan underwriting and risk assessment system designed to simplify the loan "
        "application process. It uses a multi-agent architecture in which a Master Agent coordinates "
        "specialized agents to perform eligibility checks, verification, fraud detection and risk assessment. "
        "The system integrates React.js, FastAPI and OpenAI GPT-4o-mini to provide a conversational "
        "interface and organized loan processing workflow."
    )
    
    for sh in s1.shapes:
        if sh.name == "Text 5" and sh.has_text_frame:
            tf = sh.text_frame
            p = tf.paragraphs[0]
            p.text = abstract_text
            p.font.name = FONT_BODY
            p.font.size = Pt(14)
            p.font.color.rgb = COLOR_WHITE
            p.space_after = Pt(0)
            p.line_spacing = 1.25
            
            # Remove any extra paragraphs if present
            while len(tf.paragraphs) > 1:
                p_elem = tf.paragraphs[1]._p
                p_elem.getparent().remove(p_elem)

    # ----------------------------------------------------
    # SLIDE 2 – INTRODUCTION
    # ----------------------------------------------------
    print("Updating Slide 2: Introduction...")
    s2 = prs.slides[1]
    remove_shapes_below(s2, 1.95)
    
    # Two Upper Cards
    card_w = Inches(5.82)
    card_h = Inches(2.95)
    top_y = Inches(2.00)
    
    # Left Card: Traditional Underwriting
    create_card(s2, Inches(0.72), top_y, card_w, card_h)
    tb_left = s2.shapes.add_textbox(Inches(0.92), Inches(2.15), Inches(5.42), Inches(2.65))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    
    p = tf_l.paragraphs[0]
    p.text = "Conventional Underwriting Background"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    p.space_after = Pt(8)
    
    points_l = [
        ("Eligibility Evaluation: ", "Loan underwriting evaluates a borrower's eligibility and repayment capacity against defined financial criteria."),
        ("Multi-Stage Verification: ", "Traditional loan processing involves multiple verification stages across departments and personnel."),
        ("Time and Coordination Costs: ", "Manual processing can require considerable time and coordination between different review teams.")
    ]
    for bold_prefix, rest in points_l:
        p = tf_l.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_prefix
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_TITLE
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = COLOR_TEXT_DARK
        
    # Right Card: The Case for AI Multi-Agent Processing
    create_card(s2, Inches(6.78), top_y, card_w, card_h)
    tb_right = s2.shapes.add_textbox(Inches(6.98), Inches(2.15), Inches(5.42), Inches(2.65))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p = tf_r.paragraphs[0]
    p.text = "Motivation for Multi-Agent AI System"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    p.space_after = Pt(8)
    
    points_r = [
        ("Information Inconsistency: ", "Incomplete or inconsistent applicant information can make structured assessment difficult without automated validation."),
        ("Structured AI Workflows: ", "AI can support information processing and structured decision workflows, reducing manual delays."),
        ("ARMEK Multi-Agent Orchestration: ", "ARMEK organizes these underwriting activities through an orchestrated multi-agent architecture.")
    ]
    for bold_prefix, rest in points_r:
        p = tf_r.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_prefix
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_TITLE
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = COLOR_TEXT_DARK

    # Bottom Visual Flow Container
    flow_y = Inches(5.12)
    flow_h = Inches(1.80)
    create_card(s2, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_ALT_ROW, border_color=RGBColor(203, 213, 225))
    
    # Flow Header Label
    lbl_box = s2.shapes.add_textbox(Inches(0.72), Inches(5.20), Inches(11.89), Inches(0.30))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "LOAN UNDERWRITING WORKFLOW STAGES"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT
    p.alignment = PP_ALIGN.CENTER
    
    # 5 Connected Steps:
    # Loan Application -> Information Collection -> Verification -> Risk Assessment -> Loan Decision
    steps_s2 = ["Loan Application", "Information Collection", "Verification", "Risk Assessment", "Loan Decision"]
    step_w = Inches(1.95)
    step_h = Inches(0.85)
    arrow_w = Inches(0.40)
    step_y = Inches(5.70)
    start_x = Inches(0.95)
    
    for i, step_name in enumerate(steps_s2):
        cur_x = start_x + i * (step_w + arrow_w)
        add_flow_step(s2, step_name, cur_x, step_y, step_w, step_h, step_num=i+1)
        if i < len(steps_s2) - 1:
            add_arrow_label(s2, cur_x + step_w, step_y, arrow_w, step_h)

    # ----------------------------------------------------
    # SLIDE 3 – OBJECTIVES
    # ----------------------------------------------------
    print("Updating Slide 3: Objectives...")
    s3 = prs.slides[2]
    remove_shapes_below(s3, 1.30)
    
    # Main Objective Banner
    main_box = create_card(s3, Inches(0.72), Inches(1.35), Inches(11.89), Inches(0.95), 
                           bg_color=COLOR_ACCENT_BG, border_color=COLOR_BLUE_ACCENT, border_width=1.5)
    tb_main = s3.shapes.add_textbox(Inches(0.92), Inches(1.42), Inches(11.49), Inches(0.80))
    tf_m = tb_main.text_frame
    tf_m.word_wrap = True
    
    p = tf_m.paragraphs[0]
    p.text = "Main Objective:"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT
    
    p2 = tf_m.add_paragraph()
    p2.space_before = Pt(3)
    p2.text = "To develop an AI-powered loan underwriting system that organizes applicant information processing, verification workflows and risk assessment through a multi-agent architecture."
    p2.font.name = FONT_BODY
    p2.font.size = Pt(13)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_NAVY_DARK

    # Specific Objectives: 6 Cards in 2 columns x 3 rows
    spec_objs = [
        ("1", "User-Friendly Loan Interface", "Provide a user-friendly loan application interface."),
        ("2", "Applicant Information Processing", "Collect and process applicant information in a structured and validated manner."),
        ("3", "Master Agent Coordination", "Coordinate loan assessment through a Master Agent."),
        ("4", "Eligibility & Verification Checks", "Perform eligibility and verification checks."),
        ("5", "Risk & Fraud Assessment", "Support credit-related risk assessment and fraud checks."),
        ("6", "Document Generation & Status", "Generate loan-related documents and display application status.")
    ]
    
    row_ys = [Inches(2.45), Inches(3.95), Inches(5.45)]
    card_w = Inches(5.82)
    card_h = Inches(1.35)
    
    for idx, (num, title, desc) in enumerate(spec_objs):
        col_x = Inches(0.72) if idx % 2 == 0 else Inches(6.78)
        cur_y = row_ys[idx // 2]
        
        # Container Card
        create_card(s3, col_x, cur_y, card_w, card_h)
        
        # Number Badge
        num_badge = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.18), cur_y + Inches(0.25), Inches(0.40), Inches(0.40))
        num_badge.fill.solid()
        num_badge.fill.fore_color.rgb = COLOR_NAVY_DARK
        num_badge.line.fill.background()
        p_num = num_badge.text_frame.paragraphs[0]
        p_num.text = num
        p_num.font.name = FONT_BOLD
        p_num.font.size = Pt(13)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_WHITE
        p_num.alignment = PP_ALIGN.CENTER
        
        # Content
        tb_obj = s3.shapes.add_textbox(col_x + Inches(0.70), cur_y + Inches(0.15), card_w - Inches(0.85), card_h - Inches(0.25))
        tf_o = tb_obj.text_frame
        tf_o.word_wrap = True
        
        p = tf_o.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_TITLE
        
        p_desc = tf_o.add_paragraph()
        p_desc.space_before = Pt(3)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = COLOR_TEXT_DARK

    # ----------------------------------------------------
    # SLIDE 4 – LITERATURE SURVEY (STRICT 5-COLUMN TABLE)
    # ----------------------------------------------------
    print("Updating Slide 4: Literature Survey (5-Column Table)...")
    s4 = prs.slides[3]
    remove_shapes_below(s4, 1.80)
    
    # 5-Column Table
    table_shape = s4.shapes.add_table(4, 5, Inches(0.72), Inches(2.05), Inches(11.89), Inches(4.95))
    table = table_shape.table
    
    # Column widths (Total = 11.89 in)
    table.columns[0].width = Inches(0.75)   # S. No.
    table.columns[1].width = Inches(2.15)   # Author and Year
    table.columns[2].width = Inches(3.25)   # Research Paper
    table.columns[3].width = Inches(2.74)   # Method / Focus
    table.columns[4].width = Inches(3.00)   # Key Findings
    
    headers_s4 = ["S. No.", "Author and Year", "Research Paper", "Method / Focus", "Key Findings"]
    for c_idx, h_text in enumerate(headers_s4):
        format_cell(table.cell(0, c_idx), h_text, bold=True, font_size=11, 
                    text_color=COLOR_WHITE, bg_color=COLOR_NAVY_DARK, 
                    align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT, font_name=FONT_BOLD)
        
    papers_s4 = [
        (
            "1",
            "Turiel & Aste\n(2020)",
            "Peer-to-peer loan acceptance and default prediction with artificial intelligence\nJournal: Royal Society Open Science, 7(6), 191649\nDOI: https://doi.org/10.1098/rsos.191649",
            "Machine-learning methods for loan acceptance and default prediction.",
            "Evaluated ML models on peer-to-peer lending data; demonstrated improved classification of loan applications and default prediction."
        ),
        (
            "2",
            "Xu, Lu & Xie\n(2021)",
            "Loan default prediction of Chinese P2P market: a machine learning methodology\nJournal: Scientific Reports, 11, 18759\nDOI: https://doi.org/10.1038/s41598-021-98361-6",
            "Machine-learning approaches for loan default prediction.",
            "Analyzed borrower features in P2P lending; showed ensemble and ML techniques enhance default risk identification."
        ),
        (
            "3",
            "Alonso Robisco &\nCarbó Martínez\n(2022)",
            "Measuring the model risk-adjusted performance of machine learning algorithms in credit default prediction\nJournal: Financial Innovation, 8, 70\nDOI: https://doi.org/10.1186/s40854-022-00366-1",
            "Credit-default prediction, machine-learning performance and model risk.",
            "Assessed ML performance under model risk constraints; emphasized balancing predictive accuracy with model stability."
        )
    ]
    
    for r_idx, row_data in enumerate(papers_s4):
        row_bg = COLOR_ALT_ROW if r_idx % 2 == 0 else COLOR_WHITE
        for c_idx, val in enumerate(row_data):
            format_cell(table.cell(r_idx + 1, c_idx), val, bold=(c_idx == 0 or c_idx == 1),
                        font_size=10.5, text_color=COLOR_NAVY_DARK if c_idx == 0 else COLOR_TEXT_DARK,
                        bg_color=row_bg, align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT)

    # ----------------------------------------------------
    # SLIDE 5 – EXISTING SYSTEM (ONLY Existing System!)
    # ----------------------------------------------------
    print("Updating Slide 5: Existing System (Workflow & Characteristics)...")
    s5 = prs.slides[4]
    remove_shapes_below(s5, 1.30)
    
    # Subtitle
    sub_box = s5.shapes.add_textbox(Inches(0.72), Inches(1.22), Inches(11.89), Inches(0.45))
    tf_s = sub_box.text_frame
    tf_s.word_wrap = True
    p_s = tf_s.paragraphs[0]
    p_s.text = "Conventional loan processing relies on sequential, manual stages across separate departments, creating operational overhead and review delays."
    p_s.font.name = FONT_BODY
    p_s.font.size = Pt(12)
    p_s.font.color.rgb = COLOR_TEXT_DARK

    # Two Upper Cards
    top_y = Inches(1.85)
    card_w = Inches(5.82)
    card_h = Inches(3.05)
    
    # Left Card: Characteristics Part 1
    create_card(s5, Inches(0.72), top_y, card_w, card_h)
    tb_el = s5.shapes.add_textbox(Inches(0.92), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_el = tb_el.text_frame
    tf_el.word_wrap = True
    
    p = tf_el.paragraphs[0]
    p.text = "Conventional Application & Checking"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    
    pts_el = [
        ("Application Submission: ", "Applicants submit personal and financial information through manual paper forms or static portals."),
        ("Document Collection: ", "Documents are collected and checked manually for completeness across disparate office channels."),
        ("Eligibility Assessment: ", "Eligibility is evaluated using financial criteria and manual debt calculations.")
    ]
    for bold_prefix, rest in pts_el:
        p = tf_el.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_prefix
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_TITLE
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_DARK
        
    # Right Card: Characteristics Part 2
    create_card(s5, Inches(6.78), top_y, card_w, card_h)
    tb_er = s5.shapes.add_textbox(Inches(6.98), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_er = tb_er.text_frame
    tf_er.word_wrap = True
    
    p = tf_er.paragraphs[0]
    p.text = "Review & Decision Process"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    
    pts_er = [
        ("Credit & Identity Checks: ", "Credit and identity checks are performed as isolated administrative steps."),
        ("Departmental Reviews: ", "Different stages may require separate reviews and handoffs between departments."),
        ("Decision Communication: ", "Final loan decisions are communicated to applicants after notable administrative turnaround times.")
    ]
    for bold_prefix, rest in pts_er:
        p = tf_er.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_prefix
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_TITLE
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_DARK

    # Bottom Workflow Box (REQUIRED)
    flow_y = Inches(5.10)
    flow_h = Inches(1.85)
    create_card(s5, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_ALT_ROW, border_color=RGBColor(203, 213, 225))
    
    lbl_box = s5.shapes.add_textbox(Inches(0.72), Inches(5.18), Inches(11.89), Inches(0.28))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "CONVENTIONAL LOAN PROCESSING WORKFLOW"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT
    p.alignment = PP_ALIGN.CENTER
    
    # 6 Steps: Applicant -> Application Submission -> Document Verification -> Eligibility Assessment -> Risk Assessment -> Loan Decision
    steps_s5 = ["Applicant", "Application Submission", "Document Verification", "Eligibility Assessment", "Risk Assessment", "Loan Decision"]
    step_w = Inches(1.60)
    step_h = Inches(0.85)
    arrow_w = Inches(0.35)
    step_y = Inches(5.68)
    start_x = Inches(0.92)
    
    for i, step_name in enumerate(steps_s5):
        cur_x = start_x + i * (step_w + arrow_w)
        add_flow_step(s5, step_name, cur_x, step_y, step_w, step_h, step_num=i+1)
        if i < len(steps_s5) - 1:
            add_arrow_label(s5, cur_x + step_w, step_y, arrow_w, step_h)

    # ----------------------------------------------------
    # SLIDE 6 – EXISTING SYSTEM: ADVANTAGES AND DISADVANTAGES
    # ----------------------------------------------------
    print("Updating Slide 6: Existing System: Advantages and Disadvantages...")
    s6 = prs.slides[5]
    remove_all_shapes(s6)
    
    add_slide_header(
        s6,
        badge_text="COMPARATIVE EVALUATION",
        title_text="Existing System: Advantages and Disadvantages",
        subtitle_text="Objective evaluation of traditional loan underwriting procedures, highlighting established institutional strengths alongside procedural bottlenecks."
    )
    
    col_w = Inches(5.82)
    col_h = Inches(4.85)
    top_y = Inches(2.00)
    
    # Left Column: Advantages (Green styling)
    create_card(s6, Inches(0.72), top_y, col_w, col_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER, border_width=1.5)
    
    hdr_adv = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(2.15), Inches(2.0), Inches(0.35))
    hdr_adv.fill.solid()
    hdr_adv.fill.fore_color.rgb = COLOR_GREEN_BG
    hdr_adv.line.color.rgb = COLOR_GREEN_BORDER
    hdr_adv.line.width = Pt(1.0)
    p = hdr_adv.text_frame.paragraphs[0]
    p.text = "✓  ADVANTAGES"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    p.alignment = PP_ALIGN.CENTER
    
    tb_adv = s6.shapes.add_textbox(Inches(0.95), Inches(2.65), Inches(5.35), Inches(4.00))
    tf_adv = tb_adv.text_frame
    tf_adv.word_wrap = True
    
    adv_items = [
        ("Established lending procedures.", "Built upon established lending procedures with long-standing regulatory familiarity and standardized compliance."),
        ("Human review for complex applications.", "Direct availability of human judgment for complex, non-standard, or edge-case loan applications."),
        ("Familiar verification processes.", "Verification processes are universally recognized, accepted, and audited across traditional banking institutions."),
        ("Suitable for standard banking workflows.", "Well-suited for standard banking workflows, established audit trails, and legal regulatory frameworks.")
    ]
    for idx, (bold_pfx, rest) in enumerate(adv_items):
        p = tf_adv.paragraphs[0] if idx == 0 else tf_adv.add_paragraph()
        if idx > 0:
            p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = f"{idx+1}. {bold_pfx} "
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_DARK
        
    # Right Column: Disadvantages (Red/Rose styling)
    create_card(s6, Inches(6.78), top_y, col_w, col_h, bg_color=COLOR_CARD_BG, border_color=COLOR_RED_BORDER, border_width=1.5)
    
    hdr_dis = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.01), Inches(2.15), Inches(2.2), Inches(0.35))
    hdr_dis.fill.solid()
    hdr_dis.fill.fore_color.rgb = COLOR_RED_BG
    hdr_dis.line.color.rgb = COLOR_RED_BORDER
    hdr_dis.line.width = Pt(1.0)
    p = hdr_dis.text_frame.paragraphs[0]
    p.text = "⚠  DISADVANTAGES"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED_TEXT
    p.alignment = PP_ALIGN.CENTER
    
    tb_dis = s6.shapes.add_textbox(Inches(7.01), Inches(2.65), Inches(5.35), Inches(4.00))
    tf_dis = tb_dis.text_frame
    tf_dis.word_wrap = True
    
    dis_items = [
        ("Multiple manual processing stages.", "Multi-stage manual handoffs significantly increase overall application turnaround times."),
        ("Repeated document handling.", "Repeated document handling and redundant verification steps across disparate departments."),
        ("Coordination between different departments.", "Coordination friction between separate review teams, branch managers, and credit underwriting desks."),
        ("Delays may occur during application review.", "Delays frequently occur during application review, discrepancy resolution, and manual escalation."),
        ("Limited automation in some workflows.", "Limited automation in routine verification and assessment workflows leads to higher processing effort.")
    ]
    for idx, (bold_pfx, rest) in enumerate(dis_items):
        p = tf_dis.paragraphs[0] if idx == 0 else tf_dis.add_paragraph()
        if idx > 0:
            p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"{idx+1}. {bold_pfx} "
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_NAVY_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_DARK

    # ----------------------------------------------------
    # SLIDE 7 – PROPOSED SYSTEM
    # ----------------------------------------------------
    print("Updating Slide 7: Proposed System (Architecture & Workflow)...")
    s7 = prs.slides[6]
    remove_all_shapes(s7)
    
    add_slide_header(
        s7,
        badge_text="PROPOSED SYSTEM",
        title_text="ARMEK – Agentic AI-Powered Loan Underwriting and Risk Assessment System",
        subtitle_text="A multi-agent loan underwriting architecture coordinating specialized worker agents through a central Master Agent, backed by modern web and AI technologies."
    )
    
    # Upper Section: 4 Core Technology Pillars (y=1.95, h=2.35)
    pillar_w = Inches(2.82)
    pillar_h = Inches(2.35)
    top_p_y = Inches(1.95)
    
    pillars = [
        ("React Frontend", "Applicant Interface", "React-based frontend for applicant interaction, structured information input, conversational form submission, and real-time status display."),
        ("FastAPI Backend", "Application Processing", "Python FastAPI backend for application processing, session management, REST endpoints, and Pydantic-validated data models."),
        ("AI Integration", "OpenAI GPT-4o-mini", "OpenAI GPT-4o-mini integration for AI-enabled conversational assistance, applicant guidance, and structured parameter reasoning."),
        ("Worker Agents", "Modular Verification", "Specialized agents executing eligibility checks, state reconciliation, credit assessment, fraud checks, and sanction document generation.")
    ]
    
    for idx, (title, sub, desc) in enumerate(pillars):
        px = Inches(0.72) + idx * Inches(3.02)
        create_card(s7, px, top_p_y, pillar_w, pillar_h)
        
        tb = s7.shapes.add_textbox(px + Inches(0.12), top_p_y + Inches(0.12), pillar_w - Inches(0.24), pillar_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_TITLE
        
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(2)
        p_sub.text = sub
        p_sub.font.name = FONT_BOLD
        p_sub.font.size = Pt(9.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_BLUE_ACCENT
        
        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(4)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10)
        p_desc.font.color.rgb = COLOR_TEXT_DARK

    # Bottom Section: End-to-End Orchestrated Workflow (REQUIRED)
    flow_y = Inches(4.55)
    flow_h = Inches(2.45)
    create_card(s7, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_ALT_ROW, border_color=COLOR_BLUE_ACCENT, border_width=1.5)
    
    lbl_box = s7.shapes.add_textbox(Inches(0.72), Inches(4.62), Inches(11.89), Inches(0.28))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "END-TO-END ORCHESTRATED WORKFLOW PIPELINE"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT
    p.alignment = PP_ALIGN.CENTER
    
    # Step Row 1: Applicant -> React Frontend -> FastAPI Backend -> AI Integration
    w_r1 = Inches(2.45)
    h_r1 = Inches(0.68)
    arr_w = Inches(0.40)
    r1_y = Inches(5.02)
    start_r1_x = Inches(1.10)
    
    r1_steps = ["Applicant", "React Frontend", "FastAPI Backend", "AI Integration"]
    for i, s_text in enumerate(r1_steps):
        cx = start_r1_x + i * (w_r1 + arr_w)
        add_flow_step(s7, s_text, cx, r1_y, w_r1, h_r1)
        if i < len(r1_steps) - 1:
            add_arrow_label(s7, cx + w_r1, r1_y, arr_w, h_r1)
            
    # Connector from AI Integration down to Master Agent
    arr_down = s7.shapes.add_textbox(Inches(9.65), Inches(5.62), Inches(1.0), Inches(0.35))
    p_d = arr_down.text_frame.paragraphs[0]
    p_d.text = "▼"
    p_d.alignment = PP_ALIGN.CENTER
    p_d.font.name = FONT_BOLD
    p_d.font.size = Pt(14)
    p_d.font.color.rgb = COLOR_BLUE_ACCENT
    
    # Step Row 2: Master Agent -> Specialized Worker Agents -> Assessment & Documents -> Status & Results
    r2_y = Inches(6.00)
    # Master Agent: Prominent Navy Card
    add_flow_step(s7, "★ MASTER AGENT ★\n(Central Coordination)", Inches(1.10), r2_y, Inches(3.20), Inches(0.80), is_highlighted=True)
    
    # Arrow to Worker Agents
    add_arrow_label(s7, Inches(4.35), r2_y, arr_w, Inches(0.80))
    
    # Worker Agents
    add_flow_step(s7, "Specialized Worker Agents\n(Eligibility · KYC · Bureau · Fraud)", Inches(4.80), r2_y, Inches(3.10), Inches(0.80))
    
    # Arrow to Assessment & Document Generation
    add_arrow_label(s7, Inches(7.95), r2_y, arr_w, Inches(0.80))
    
    # Assessment & Document Generation
    add_flow_step(s7, "Assessment & Docs\n(ReportLab Sanction Letter)", Inches(8.40), r2_y, Inches(2.20), Inches(0.80))
    
    # Arrow to Status & Results
    add_arrow_label(s7, Inches(10.65), r2_y, Inches(0.30), Inches(0.80))
    
    # Application Status & Results
    add_flow_step(s7, "Application\nStatus & Results", Inches(11.00), r2_y, Inches(1.35), Inches(0.80))

    # ----------------------------------------------------
    # SLIDE 8 – PROPOSED SYSTEM: ADVANTAGES
    # ----------------------------------------------------
    print("Updating Slide 8: Proposed System: Advantages (6 Advantages)...")
    s8 = prs.slides[7]
    remove_all_shapes(s8)
    
    add_slide_header(
        s8,
        badge_text="SYSTEM BENEFITS",
        title_text="Proposed System: Advantages",
        subtitle_text="Key operational and structural advantages delivered by ARMEK's orchestrated multi-agent underwriting architecture."
    )
    
    adv_6 = [
        ("Organized Multi-Agent Workflow", "Organized multi-agent workflow distributing assessment across specialized agents, preventing processing bottlenecks in any single stage."),
        ("Reduced Manual Coordination", "Reduced repetitive manual coordination steps through automated agent handoffs orchestrated by the central Master Agent."),
        ("Centralized Information Processing", "Centralized applicant information processing in a unified pipeline, eliminating data duplication and state inconsistency across stages."),
        ("Conversational AI-Based Interaction", "Conversational AI-based interaction powered by GPT-4o-mini, providing an intuitive and accessible experience for loan applicants."),
        ("Separate Worker Responsibilities", "Separate responsibilities for specialized worker agents (eligibility, reconciliation, credit checks, fraud detection, and document generation)."),
        ("Clear Status & Document Workflow", "Clear application status and document generation workflow, providing real-time visibility and automated sanction letter generation upon approval.")
    ]
    
    row_ys = [Inches(2.10), Inches(3.75), Inches(5.40)]
    card_w = Inches(5.82)
    card_h = Inches(1.45)
    
    icon_files = [
        "extracted_pptx_images/s7_img5_31.png",
        "extracted_pptx_images/s7_img8_32.png",
        "extracted_pptx_images/s7_img11_33.png",
        "extracted_pptx_images/s7_img14_34.png",
        "extracted_pptx_images/s7_img17_35.png",
        "extracted_pptx_images/s7_img20_36.png"
    ]
    
    for idx, (title, desc) in enumerate(adv_6):
        col_x = Inches(0.72) if idx % 2 == 0 else Inches(6.78)
        cur_y = row_ys[idx // 2]
        
        create_card(s8, col_x, cur_y, card_w, card_h)
        
        if idx < len(icon_files) and os.path.exists(icon_files[idx]):
            s8.shapes.add_picture(icon_files[idx], col_x + Inches(0.20), cur_y + Inches(0.20), Inches(0.40), Inches(0.40))
            text_x = col_x + Inches(0.75)
            text_w = card_w - Inches(0.90)
        else:
            text_x = col_x + Inches(0.20)
            text_w = card_w - Inches(0.40)
            
        tb = s8.shapes.add_textbox(text_x, cur_y + Inches(0.12), text_w, card_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_TITLE
        
        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(4)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = COLOR_TEXT_DARK

    # ----------------------------------------------------
    # SLIDE 9 – SYSTEM ARCHITECTURE DIAGRAM
    # ----------------------------------------------------
    print("Updating Slide 9: System Architecture Diagram (5 Distinct Layers)...")
    s9 = prs.slides[8]
    remove_all_shapes(s9)
    
    add_slide_header(
        s9,
        badge_text="TECHNICAL DESIGN",
        title_text="System Architecture Diagram",
        subtitle_text="Five-tier modular architecture illustrating data flow, agent orchestration, and component responsibilities in ARMEK."
    )
    
    layer_w = Inches(11.89)
    lx = Inches(0.72)
    
    # Layer 1 – User Interface
    l1_y = Inches(1.95)
    l1_h = Inches(0.72)
    create_card(s9, lx, l1_y, layer_w, l1_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_BLUE_ACCENT)
    tb1 = s9.shapes.add_textbox(lx + Inches(0.20), l1_y + Inches(0.10), layer_w - Inches(0.40), l1_h - Inches(0.20))
    p = tb1.text_frame.paragraphs[0]
    p.text = "Layer 1 – User Interface: Applicant  ·  React.js  ·  Bootstrap 5  ·  Axios"
    p.font.name = FONT_BOLD
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    p.alignment = PP_ALIGN.CENTER
    
    # Arrow 1 to 2
    add_arrow_label(s9, Inches(6.30), Inches(2.68), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 2 – Backend
    l2_y = Inches(2.90)
    l2_h = Inches(0.72)
    create_card(s9, lx, l2_y, layer_w, l2_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_BLUE_ACCENT)
    tb2 = s9.shapes.add_textbox(lx + Inches(0.20), l2_y + Inches(0.10), layer_w - Inches(0.40), l2_h - Inches(0.20))
    p = tb2.text_frame.paragraphs[0]
    p.text = "Layer 2 – Backend: Python FastAPI  ·  API Endpoints  ·  Pydantic Validation"
    p.font.name = FONT_BOLD
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    p.alignment = PP_ALIGN.CENTER
    
    # Arrow 2 to 3
    add_arrow_label(s9, Inches(6.30), Inches(3.63), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 3 – AI Orchestration (PROMINENT CENTER)
    l3_y = Inches(3.85)
    l3_h = Inches(0.85)
    create_card(s9, lx, l3_y, layer_w, l3_h, bg_color=COLOR_NAVY_DARK, border_color=COLOR_BLUE_ACCENT, border_width=2.0)
    tb3 = s9.shapes.add_textbox(lx + Inches(0.20), l3_y + Inches(0.12), layer_w - Inches(0.40), l3_h - Inches(0.24))
    p = tb3.text_frame.paragraphs[0]
    p.text = "Layer 3 – AI Orchestration: OpenAI API  ·  GPT-4o-mini  ·  ★ MASTER AGENT ★"
    p.font.name = FONT_BOLD
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER
    p_sub = tb3.text_frame.add_paragraph()
    p_sub.space_before = Pt(2)
    p_sub.text = "Coordinates the 5 specialized worker agents, manages session state, and evaluates decision conditions"
    p_sub.font.name = FONT_BODY
    p_sub.font.size = Pt(10)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)
    p_sub.alignment = PP_ALIGN.CENTER
    
    # Arrow 3 to 4 (Delegation)
    add_arrow_label(s9, Inches(6.30), Inches(4.71), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 4 – Worker Agent Processing (5 Specialized Worker Agents)
    l4_y = Inches(4.92)
    l4_h = Inches(1.10)
    create_card(s9, lx, l4_y, layer_w, l4_h, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER)
    
    lbl_l4 = s9.shapes.add_textbox(lx, l4_y + Inches(0.04), layer_w, Inches(0.25))
    p = lbl_l4.text_frame.paragraphs[0]
    p.text = "Layer 4 – Worker Agent Processing (Master Agent Coordinates the Five Worker Agents)"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE_ACCENT
    p.alignment = PP_ALIGN.CENTER
    
    workers = [
        ("Eligibility Agent", "Income & FOIR checks"),
        ("State Reconciliation Agent", "Cross-field consistency"),
        ("Credit Bureau Agent", "Credit-related risk checks"),
        ("Fraud Detection Agent", "Anomaly & fraud checks"),
        ("Document Generation Agent", "Sanction letter generation")
    ]
    w_card_w = Inches(2.25)
    w_card_h = Inches(0.68)
    w_start_x = Inches(0.85)
    w_gap = Inches(0.12)
    
    for idx, (w_name, w_desc) in enumerate(workers):
        wx = w_start_x + idx * (w_card_w + w_gap)
        create_card(s9, wx, l4_y + Inches(0.32), w_card_w, w_card_h, bg_color=COLOR_ACCENT_BG, border_color=RGBColor(191, 219, 254))
        tb_w = s9.shapes.add_textbox(wx, l4_y + Inches(0.35), w_card_w, w_card_h - Inches(0.06))
        tf_w = tb_w.text_frame
        tf_w.word_wrap = True
        
        p = tf_w.paragraphs[0]
        p.text = w_name
        p.font.name = FONT_BOLD
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_NAVY_DARK
        p.alignment = PP_ALIGN.CENTER
        
        p2 = tf_w.add_paragraph()
        p2.space_before = Pt(1)
        p2.text = w_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = COLOR_TEXT_DARK
        p2.alignment = PP_ALIGN.CENTER

    # Arrow 4 to 5
    add_arrow_label(s9, Inches(6.30), Inches(6.03), Inches(0.72), Inches(0.18), text="▼")
    
    # Layer 5 – Output
    l5_y = Inches(6.20)
    l5_h = Inches(0.55)
    create_card(s9, lx, l5_y, layer_w, l5_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_BLUE_ACCENT)
    tb5 = s9.shapes.add_textbox(lx + Inches(0.20), l5_y + Inches(0.06), layer_w - Inches(0.40), l5_h - Inches(0.12))
    p = tb5.text_frame.paragraphs[0]
    p.text = "Layer 5 – Output: Loan Assessment Result  ·  Loan-Related Documents  ·  Application Status"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_TITLE
    p.alignment = PP_ALIGN.CENTER
    
    # Simulated Verification Note Box
    note_box = create_card(s9, lx, Inches(6.85), layer_w, Inches(0.38), 
                           bg_color=RGBColor(254, 252, 232), border_color=RGBColor(250, 204, 21), border_width=1.0)
    tb_n = s9.shapes.add_textbox(lx + Inches(0.20), Inches(6.87), layer_w - Inches(0.40), Inches(0.34))
    p = tb_n.text_frame.paragraphs[0]
    p.text = "Note: Credit-related assessment and bureau verification are simulated internally within the ARMEK pipeline. The system does not connect to live external government or credit bureau APIs."
    p.font.name = FONT_BODY
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)
    p.alignment = PP_ALIGN.CENTER

    # ----------------------------------------------------
    # SLIDE 10 – REFERENCES (STRICT 5-COLUMN TABLE)
    # ----------------------------------------------------
    print("Updating Slide 10: References (5-Column Table with 7 References)...")
    s10 = prs.slides[9]
    remove_shapes_below(s10, 1.80)
    
    # Table (8 rows x 5 columns: 1 Header + 7 References)
    table_shape_10 = s10.shapes.add_table(8, 5, Inches(0.72), Inches(2.00), Inches(11.89), Inches(5.10))
    table_10 = table_shape_10.table
    
    # Column widths (Total = 11.89 in)
    table_10.columns[0].width = Inches(0.70)   # S. No.
    table_10.columns[1].width = Inches(2.20)   # Author(s) & Year
    table_10.columns[2].width = Inches(3.40)   # Paper / Documentation Title
    table_10.columns[3].width = Inches(2.40)   # Journal / Source
    table_10.columns[4].width = Inches(3.19)   # DOI / URL
    
    headers_s10 = ["S. No.", "Author(s) & Year", "Paper / Documentation Title", "Journal / Source", "DOI / URL"]
    for c_idx, h_text in enumerate(headers_s10):
        format_cell(table_10.cell(0, c_idx), h_text, bold=True, font_size=11, 
                    text_color=COLOR_WHITE, bg_color=COLOR_NAVY_DARK, 
                    align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT, font_name=FONT_BOLD)
        
    refs_s10 = [
        (
            "1",
            "Turiel, J. D., &\nAste, T. (2020)",
            "Peer-to-peer loan acceptance and default prediction with artificial intelligence",
            "Royal Society Open Science, 7(6), 191649",
            "https://doi.org/10.1098/rsos.191649"
        ),
        (
            "2",
            "Xu, J., Lu, Z., &\nXie, Y. (2021)",
            "Loan default prediction of Chinese P2P market: a machine learning methodology",
            "Scientific Reports, 11, 18759",
            "https://doi.org/10.1038/s41598-021-98361-6"
        ),
        (
            "3",
            "Alonso Robisco, A., &\nCarbó Martínez, J. M. (2022)",
            "Measuring the model risk-adjusted performance of machine learning algorithms in credit default prediction",
            "Financial Innovation, 8, 70",
            "https://doi.org/10.1186/s40854-022-00366-1"
        ),
        (
            "4",
            "Meta Open Source\n(2024)",
            "React Official Documentation: The Library for Web and Native User Interfaces",
            "React Technical Documentation\n(https://react.dev/)",
            "https://react.dev/"
        ),
        (
            "5",
            "Sebastián Ramírez\n(2024)",
            "FastAPI Official Documentation: Modern, Fast Web Framework for Python",
            "FastAPI Framework Documentation\n(https://fastapi.tiangolo.com/)",
            "https://fastapi.tiangolo.com/"
        ),
        (
            "6",
            "OpenAI\n(2024)",
            "OpenAI API Documentation and Platform Reference Guides",
            "OpenAI Developer Documentation\n(https://platform.openai.com/docs/)",
            "https://platform.openai.com/docs/"
        ),
        (
            "7",
            "Bootstrap Authors\n(2024)",
            "Bootstrap 5 Documentation: The World's Most Popular Front-End Toolkit",
            "Bootstrap Technical Documentation\n(https://getbootstrap.com/docs/)",
            "https://getbootstrap.com/docs/"
        )
    ]
    
    for r_idx, row_data in enumerate(refs_s10):
        row_bg = COLOR_ALT_ROW if r_idx % 2 == 0 else COLOR_WHITE
        for c_idx, val in enumerate(row_data):
            format_cell(table_10.cell(r_idx + 1, c_idx), val, bold=(c_idx == 0 or c_idx == 1),
                        font_size=9.5, text_color=COLOR_NAVY_DARK if c_idx == 0 else COLOR_TEXT_DARK,
                        bg_color=row_bg, align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT)

    # ----------------------------------------------------
    # SAVE COMPLETED PRESENTATION
    # ----------------------------------------------------
    print(f"\nSaving final presentation to: {output_pptx}")
    prs.save(output_pptx)
    print("SUCCESS: Final presentation saved successfully!")

if __name__ == "__main__":
    main()
