"""
build_college_pptx.py
Generates the updated ARMEK college presentation with 11 slides:
- Follows the Kuppam Engineering College reference image for Slide 1
- Formal, clean green-and-white college presentation styling across all slides
- Strict tables on Slide 4 (Literature Survey) and Slide 10 (References)
- 11th Slide: Conclusion
- Zero occurrences of 'prototype'
Saves output as ARMEK_Final_Project_Review_Updated.pptx
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==========================================
# CONSTANTS & THEME PALETTE (GREEN & WHITE)
# ==========================================
FONT_HEADING = "Trebuchet MS"
FONT_BODY = "Segoe UI"
FONT_BOLD = "Segoe UI"

COLOR_GREEN_DARK = RGBColor(35, 95, 62)     # #235F3E - Forest green
COLOR_GREEN_PRIMARY = RGBColor(44, 112, 75) # #2C704B - College green
COLOR_GREEN_LIGHT = RGBColor(240, 248, 242) # #F0F8F2 - Mint accent background
COLOR_GREEN_BORDER = RGBColor(167, 208, 183)# #A7D0B7 - Subtle green border
COLOR_GREEN_TEXT = RGBColor(27, 57, 49)     # #1B3931 - Dark green/charcoal

COLOR_TEXT_MAIN = RGBColor(33, 37, 41)      # #212529 - Charcoal body
COLOR_TEXT_MUTED = RGBColor(90, 102, 100)   # #5A6664 - Muted gray
COLOR_WHITE = RGBColor(255, 255, 255)

COLOR_CARD_BG = RGBColor(255, 255, 255)
COLOR_CARD_BORDER = RGBColor(218, 228, 222) # Subtle border

COLOR_RED_BG = RGBColor(254, 242, 242)
COLOR_RED_BORDER = RGBColor(239, 68, 68)
COLOR_RED_TEXT = RGBColor(153, 27, 27)

COLOR_ALT_ROW = RGBColor(244, 248, 245)

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def remove_all_shapes(slide):
    """Removes all shapes from a slide."""
    for sh in list(slide.shapes):
        sp = sh._element
        sp.getparent().remove(sp)

def add_college_header(slide, slide_num, total_slides=11):
    """Adds formal college top banner and footer to slides 2 through 11."""
    # Top decorative banner image if exists
    if os.path.exists("college_assets/top_banner.png"):
        slide.shapes.add_picture("college_assets/top_banner.png", Inches(0), Inches(0), Inches(13.333), Inches(0.40))
    else:
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = COLOR_GREEN_PRIMARY
        top_bar.line.fill.background()

    # Footer
    footer_y = Inches(7.10)
    # Left footer: Project title
    tb_fl = slide.shapes.add_textbox(Inches(0.72), footer_y, Inches(7.0), Inches(0.30))
    p_fl = tb_fl.text_frame.paragraphs[0]
    p_fl.text = "ARMEK  /  PROJECT PRESENTATION"
    p_fl.font.name = FONT_BOLD
    p_fl.font.size = Pt(8.5)
    p_fl.font.bold = True
    p_fl.font.color.rgb = COLOR_GREEN_PRIMARY
    
    # Right footer: Slide numbering and college name
    tb_fr = slide.shapes.add_textbox(Inches(7.5), footer_y, Inches(5.11), Inches(0.30))
    p_fr = tb_fr.text_frame.paragraphs[0]
    p_fr.text = f"{slide_num:02d}  -  KUPPAM ENGINEERING COLLEGE"
    p_fr.alignment = PP_ALIGN.RIGHT
    p_fr.font.name = FONT_BODY
    p_fr.font.size = Pt(8.5)
    p_fr.font.bold = True
    p_fr.font.color.rgb = COLOR_TEXT_MUTED

def add_slide_title(slide, badge_text, title_text, subtitle_text=None, top_y=0.55):
    """Adds a standard green-and-white slide header with category badge, title and subtitle."""
    # Category badge pill
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(top_y), Inches(2.5), Inches(0.24))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_GREEN_LIGHT
    badge.line.color.rgb = COLOR_GREEN_BORDER
    badge.line.width = Pt(0.75)
    
    tf = badge.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.10)
    tf.margin_right = Inches(0.10)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    
    p = tf.paragraphs[0]
    p.text = badge_text.upper()
    p.font.name = FONT_BOLD
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    p.alignment = PP_ALIGN.CENTER

    # Main Title
    tb_title = slide.shapes.add_textbox(Inches(0.72), Inches(top_y + 0.28), Inches(11.89), Inches(0.48))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = Inches(0)
    tf_t.margin_right = Inches(0)
    tf_t.margin_top = Inches(0)
    tf_t.margin_bottom = Inches(0)
    
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(24)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_GREEN_TEXT

    # Subtitle
    if subtitle_text:
        tb_sub = slide.shapes.add_textbox(Inches(0.72), Inches(top_y + 0.78), Inches(11.89), Inches(0.42))
        tf_s = tb_sub.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = Inches(0)
        tf_s.margin_right = Inches(0)
        tf_s.margin_top = Inches(0)
        tf_s.margin_bottom = Inches(0)
        
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle_text
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = COLOR_TEXT_MUTED

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER, border_width=1.0):
    """Creates a clean rounded rectangle container."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)
    else:
        card.line.fill.background()
    return card

def add_step_card(slide, text, left, top, width, height, step_num=None, is_highlighted=False):
    """Creates a single step in a process workflow."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    if is_highlighted:
        card.fill.fore_color.rgb = COLOR_GREEN_DARK
        card.line.color.rgb = COLOR_GREEN_PRIMARY
        card.line.width = Pt(1.5)
    else:
        card.fill.fore_color.rgb = COLOR_GREEN_LIGHT
        card.line.color.rgb = COLOR_GREEN_BORDER
        card.line.width = Pt(1.0)
        
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
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
    p.font.color.rgb = COLOR_WHITE if is_highlighted else COLOR_GREEN_TEXT

def add_arrow_label(slide, left, top, width, height, text="→"):
    """Adds a directional connector between cards."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_BOLD
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_PRIMARY

def format_cell(cell, text, bold=False, font_size=10, text_color=COLOR_TEXT_MAIN, 
                bg_color=None, align=PP_ALIGN.LEFT, font_name=FONT_BODY):
    """Formats an editable table cell."""
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_WHITE
        
    cell.margin_left = Inches(0.10)
    cell.margin_right = Inches(0.10)
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

# ==========================================
# MAIN SCRIPT EXECUTION
# ==========================================

def build_presentation():
    input_file = "ARMEK_Final_Project_Review.pptx"
    output_file = "ARMEK_Final_Project_Review_Updated.pptx"
    
    print(f"Loading presentation: {input_file}")
    prs = pptx.Presentation(input_file)
    print(f"Initial slide count: {len(prs.slides)}")

    # Ensure presentation has exactly 11 slides
    while len(prs.slides) < 11:
        layout = prs.slide_layouts[0]
        prs.slides.add_slide(layout)
    while len(prs.slides) > 11:
        r_sp = prs.slides[-1]._element
        r_sp.getparent().remove(r_sp)
    print(f"Adjusted to exactly: {len(prs.slides)} slides.")

    # ----------------------------------------------------
    # SLIDE 1: TITLE PAGE (FOLLOWING REFERENCE IMAGE)
    # ----------------------------------------------------
    print("Building Slide 1: College Reference Format...")
    s1 = prs.slides[0]
    remove_all_shapes(s1)

    # 1. Top Decorative Green Wave Banner
    if os.path.exists("college_assets/top_banner.png"):
        s1.shapes.add_picture("college_assets/top_banner.png", Inches(0), Inches(0), Inches(13.333), Inches(0.48))
    else:
        top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.35))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = COLOR_GREEN_PRIMARY
        top_bar.line.fill.background()

    # 2. College Logo at upper-left
    if os.path.exists("college_assets/kec_logo.png"):
        s1.shapes.add_picture("college_assets/kec_logo.png", Inches(0.85), Inches(0.65), Inches(1.15), Inches(1.15))

    # 3. Institutional Details in header (Middle)
    tb_head = s1.shapes.add_textbox(Inches(2.20), Inches(0.68), Inches(8.50), Inches(1.10))
    tf_h = tb_head.text_frame
    tf_h.word_wrap = True
    
    p = tf_h.paragraphs[0]
    p.text = "KUPPAM ENGINEERING COLLEGE (Autonomous)"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    
    p2 = tf_h.add_paragraph()
    p2.space_before = Pt(3)
    p2.text = "(Recognized by AICTE & Affiliated to JNTUA)"
    p2.font.name = FONT_BODY
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    
    p3 = tf_h.add_paragraph()
    p3.space_before = Pt(1)
    p3.text = "KES Nagar, Kuppam – 517425, Chittoor Dist., A.P."
    p3.font.name = FONT_BODY
    p3.font.size = Pt(11)
    p3.font.color.rgb = COLOR_TEXT_MUTED

    # 4. NAAC Logo at upper-right
    if os.path.exists("college_assets/naac_logo.png"):
        s1.shapes.add_picture("college_assets/naac_logo.png", Inches(11.20), Inches(0.65), Inches(1.25), Inches(1.10))

    # 5. Thin Divider Line
    line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(1.95), Inches(11.60), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_GREEN_BORDER
    line.line.fill.background()

    # 6. PROJECT PRESENTATION heading
    tb_pres = s1.shapes.add_textbox(Inches(0.85), Inches(2.10), Inches(11.60), Inches(0.35))
    p_pr = tb_pres.text_frame.paragraphs[0]
    p_pr.text = "P R O J E C T   P R E S E N T A T I O N"
    p_pr.alignment = PP_ALIGN.CENTER
    p_pr.font.name = FONT_BOLD
    p_pr.font.size = Pt(12)
    p_pr.font.bold = True
    p_pr.font.color.rgb = COLOR_GREEN_PRIMARY

    # 7. Project Title
    tb_title = s1.shapes.add_textbox(Inches(0.85), Inches(2.45), Inches(11.60), Inches(0.55))
    p_ti = tb_title.text_frame.paragraphs[0]
    p_ti.text = "ARMEK – Agentic AI-Powered Loan Underwriting and Risk Assessment System"
    p_ti.alignment = PP_ALIGN.CENTER
    p_ti.font.name = FONT_HEADING
    p_ti.font.size = Pt(21)
    p_ti.font.bold = True
    p_ti.font.color.rgb = COLOR_GREEN_TEXT

    # 8. Abstract as ONE continuous paragraph
    abstract_text = (
        "ARMEK is an AI-powered loan underwriting and risk assessment system designed to simplify the loan "
        "application process. It uses a multi-agent architecture in which a Master Agent coordinates "
        "specialized agents to perform eligibility checks, verification, fraud detection and risk assessment. "
        "The system integrates React.js, FastAPI and OpenAI GPT-4o-mini to provide a conversational "
        "interface and organized loan processing workflow."
    )
    
    # Abstract Card Container
    create_card(s1, Inches(0.85), Inches(3.10), Inches(11.60), Inches(1.05), bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER, border_width=1.0)
    tb_abs = s1.shapes.add_textbox(Inches(0.95), Inches(3.14), Inches(11.40), Inches(0.95))
    tf_a = tb_abs.text_frame
    tf_a.word_wrap = True
    p_a = tf_a.paragraphs[0]
    p_a.text = abstract_text
    p_a.alignment = PP_ALIGN.CENTER
    p_a.font.name = FONT_BODY
    p_a.font.size = Pt(11)
    p_a.font.color.rgb = COLOR_TEXT_MAIN
    p_a.line_spacing = 1.25

    # 9. Two Bottom Cards (SUBMITTED BY & UNDER THE GUIDANCE OF)
    bottom_y = Inches(4.30)
    box_w = Inches(5.65)
    box_h = Inches(2.40)
    
    # Left Card: SUBMITTED BY
    create_card(s1, Inches(0.85), bottom_y, box_w, box_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER, border_width=1.2)
    tb_subm = s1.shapes.add_textbox(Inches(1.05), bottom_y + Inches(0.12), box_w - Inches(0.35), box_h - Inches(0.20))
    tf_sm = tb_subm.text_frame
    tf_sm.word_wrap = True
    
    p = tf_sm.paragraphs[0]
    p.text = "SUBMITTED BY"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    
    p_m = tf_sm.add_paragraph()
    p_m.space_before = Pt(4)
    p_m.text = "Team Members :"
    p_m.font.name = FONT_BOLD
    p_m.font.size = Pt(10.5)
    p_m.font.bold = True
    p_m.font.color.rgb = COLOR_TEXT_MAIN
    
    team_members = [
        "1.  Y N PAWANKALYAN",
        "2.  V RAMYA",
        "3.  C NANDINI",
        "4.  S DILLI RAJESH"
    ]
    for tm in team_members:
        p_t = tf_sm.add_paragraph()
        p_t.space_before = Pt(1)
        p_t.text = "     " + tm
        p_t.font.name = FONT_BODY
        p_t.font.size = Pt(10)
        p_t.font.color.rgb = COLOR_TEXT_MAIN
        
    p_dept = tf_sm.add_paragraph()
    p_dept.space_before = Pt(5)
    p_dept.text = "Department : Computer Science and Engineering"
    p_dept.font.name = FONT_BODY
    p_dept.font.size = Pt(10)
    p_dept.font.color.rgb = COLOR_TEXT_MUTED
    
    p_batch = tf_sm.add_paragraph()
    p_batch.space_before = Pt(1)
    p_batch.text = "Batch            : 2023–2027"
    p_batch.font.name = FONT_BODY
    p_batch.font.size = Pt(10)
    p_batch.font.color.rgb = COLOR_TEXT_MUTED

    # Right Card: UNDER THE GUIDANCE OF
    create_card(s1, Inches(6.80), bottom_y, box_w, box_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER, border_width=1.2)
    tb_guide = s1.shapes.add_textbox(Inches(7.00), bottom_y + Inches(0.12), box_w - Inches(0.35), box_h - Inches(0.20))
    tf_g = tb_guide.text_frame
    tf_g.word_wrap = True
    
    p = tf_g.paragraphs[0]
    p.text = "UNDER THE GUIDANCE OF"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    
    p_gname = tf_g.add_paragraph()
    p_gname.space_before = Pt(6)
    p_gname.text = "Mrs.K.Shandhini, M.E.,"
    p_gname.font.name = FONT_HEADING
    p_gname.font.size = Pt(14)
    p_gname.font.bold = True
    p_gname.font.color.rgb = COLOR_GREEN_TEXT
    
    p_gdes = tf_g.add_paragraph()
    p_gdes.space_before = Pt(3)
    p_gdes.text = "Assistant Professor"
    p_gdes.font.name = FONT_BODY
    p_gdes.font.size = Pt(11)
    p_gdes.font.color.rgb = COLOR_TEXT_MAIN
    
    p_gdept = tf_g.add_paragraph()
    p_gdept.space_before = Pt(3)
    p_gdept.text = "Department of Computer Science and Engineering"
    p_gdept.font.name = FONT_BODY
    p_gdept.font.size = Pt(10.5)
    p_gdept.font.color.rgb = COLOR_TEXT_MUTED
    
    p_gyear = tf_g.add_paragraph()
    p_gyear.space_before = Pt(3)
    p_gyear.text = "Academic Year : 2026–2027"
    p_gyear.font.name = FONT_BODY
    p_gyear.font.size = Pt(10.5)
    p_gyear.font.color.rgb = COLOR_TEXT_MUTED

    # 10. Footer on Slide 1
    tb_fl1 = s1.shapes.add_textbox(Inches(0.85), Inches(6.85), Inches(5.0), Inches(0.30))
    p_fl1 = tb_fl1.text_frame.paragraphs[0]
    p_fl1.text = "ARMEK  /  PROJECT PRESENTATION"
    p_fl1.font.name = FONT_BOLD
    p_fl1.font.size = Pt(8.5)
    p_fl1.font.bold = True
    p_fl1.font.color.rgb = COLOR_GREEN_PRIMARY
    
    tb_fr1 = s1.shapes.add_textbox(Inches(6.80), Inches(6.85), Inches(5.65), Inches(0.30))
    p_fr1 = tb_fr1.text_frame.paragraphs[0]
    p_fr1.text = "01  -  KUPPAM ENGINEERING COLLEGE"
    p_fr1.alignment = PP_ALIGN.RIGHT
    p_fr1.font.name = FONT_BODY
    p_fr1.font.size = Pt(8.5)
    p_fr1.font.bold = True
    p_fr1.font.color.rgb = COLOR_TEXT_MUTED

    # ----------------------------------------------------
    # SLIDE 2: INTRODUCTION
    # ----------------------------------------------------
    print("Building Slide 2: Introduction...")
    s2 = prs.slides[1]
    remove_all_shapes(s2)
    add_college_header(s2, 2)
    add_slide_title(
        s2,
        badge_text="BACKGROUND",
        title_text="Introduction",
        subtitle_text="Loan underwriting evaluates a borrower's eligibility and repayment capacity against defined financial criteria."
    )
    
    card_w = Inches(5.82)
    card_h = Inches(3.10)
    top_y = Inches(1.85)
    
    # Left Card: Conventional Underwriting
    create_card(s2, Inches(0.72), top_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_left = s2.shapes.add_textbox(Inches(0.92), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    
    p = tf_l.paragraphs[0]
    p.text = "Conventional Loan Underwriting"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    
    points_l = [
        ("Meaning of Underwriting: ", "Loan underwriting evaluates a borrower's eligibility and repayment capacity against defined institutional risk parameters."),
        ("Traditional Multi-Stage Process: ", "Traditional loan processing involves multiple verification stages across disparate departments and manual review queues."),
        ("Time and Coordination: ", "Manual processing can require considerable time and coordination, leading to delays and administrative friction.")
    ]
    for bold_pfx, rest in points_l:
        p = tf_l.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + bold_pfx
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_GREEN_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_MAIN
        
    # Right Card: Multi-Agent AI Approach
    create_card(s2, Inches(6.78), top_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_right = s2.shapes.add_textbox(Inches(6.98), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    
    p = tf_r.paragraphs[0]
    p.text = "Motivation for Multi-Agent Architecture"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    
    points_r = [
        ("Data Consistency Challenges: ", "Incomplete or inconsistent applicant information makes manual assessment difficult without automated validation."),
        ("Structured AI Workflows: ", "AI can support information processing and structured decision workflows, organizing tasks with greater consistency."),
        ("ARMEK Multi-Agent Solution: ", "ARMEK organizes these underwriting activities through a multi-agent architecture coordinating eligibility, KYC, and risk checks.")
    ]
    for bold_pfx, rest in points_r:
        p = tf_r.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + bold_pfx
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_GREEN_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_MAIN

    # Bottom Visual Flow
    flow_y = Inches(5.12)
    flow_h = Inches(1.80)
    create_card(s2, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
    
    lbl_box = s2.shapes.add_textbox(Inches(0.72), Inches(5.20), Inches(11.89), Inches(0.28))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "LOAN UNDERWRITING WORKFLOW PROCESS"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    p.alignment = PP_ALIGN.CENTER
    
    steps_s2 = ["Loan Application", "Information Collection", "Verification", "Risk Assessment", "Loan Decision"]
    step_w = Inches(1.95)
    step_h = Inches(0.85)
    arrow_w = Inches(0.40)
    step_y = Inches(5.68)
    start_x = Inches(0.95)
    
    for i, step_name in enumerate(steps_s2):
        cur_x = start_x + i * (step_w + arrow_w)
        add_step_card(s2, step_name, cur_x, step_y, step_w, step_h, step_num=i+1)
        if i < len(steps_s2) - 1:
            add_arrow_label(s2, cur_x + step_w, step_y, arrow_w, step_h)

    # ----------------------------------------------------
    # SLIDE 3: OBJECTIVES
    # ----------------------------------------------------
    print("Building Slide 3: Objectives...")
    s3 = prs.slides[2]
    remove_all_shapes(s3)
    add_college_header(s3, 3)
    add_slide_title(
        s3,
        badge_text="PROJECT GOALS",
        title_text="Objectives",
        subtitle_text="Primary aim and specific engineering goals established for the ARMEK multi-agent loan underwriting system."
    )
    
    # Main Objective Card
    create_card(s3, Inches(0.72), Inches(1.85), Inches(11.89), Inches(1.00), bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_PRIMARY, border_width=1.5)
    tb_main = s3.shapes.add_textbox(Inches(0.92), Inches(1.92), Inches(11.49), Inches(0.85))
    tf_m = tb_main.text_frame
    tf_m.word_wrap = True
    
    p = tf_m.paragraphs[0]
    p.text = "Main Objective:"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    
    p2 = tf_m.add_paragraph()
    p2.space_before = Pt(3)
    p2.text = "To develop an AI-powered loan underwriting system that organizes applicant information processing, verification workflows and risk assessment through a multi-agent architecture."
    p2.font.name = FONT_BODY
    p2.font.size = Pt(13)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_GREEN_TEXT

    # Specific Objectives: 6 Cards in 2 columns x 3 rows
    spec_objs = [
        ("1", "User-Friendly Interface", "Provide a user-friendly loan application interface."),
        ("2", "Information Processing", "Collect and process applicant information in a structured and validated manner."),
        ("3", "Master Agent Coordination", "Coordinate assessment through a Master Agent."),
        ("4", "Eligibility & Verification", "Perform eligibility and verification checks."),
        ("5", "Risk & Fraud Checks", "Support credit-related risk assessment and fraud checks."),
        ("6", "Document Generation", "Generate loan-related documents and display application status.")
    ]
    
    row_ys = [Inches(3.05), Inches(4.35), Inches(5.65)]
    card_w = Inches(5.82)
    card_h = Inches(1.15)
    
    for idx, (num, title, desc) in enumerate(spec_objs):
        col_x = Inches(0.72) if idx % 2 == 0 else Inches(6.78)
        cur_y = row_ys[idx // 2]
        
        create_card(s3, col_x, cur_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
        
        num_badge = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.18), cur_y + Inches(0.20), Inches(0.36), Inches(0.36))
        num_badge.fill.solid()
        num_badge.fill.fore_color.rgb = COLOR_GREEN_DARK
        num_badge.line.fill.background()
        p_num = num_badge.text_frame.paragraphs[0]
        p_num.text = num
        p_num.font.name = FONT_BOLD
        p_num.font.size = Pt(12)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_WHITE
        p_num.alignment = PP_ALIGN.CENTER
        
        tb_obj = s3.shapes.add_textbox(col_x + Inches(0.68), cur_y + Inches(0.10), card_w - Inches(0.80), card_h - Inches(0.18))
        tf_o = tb_obj.text_frame
        tf_o.word_wrap = True
        
        p = tf_o.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_GREEN_TEXT
        
        p_desc = tf_o.add_paragraph()
        p_desc.space_before = Pt(2)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = COLOR_TEXT_MAIN

    # ----------------------------------------------------
    # SLIDE 4: LITERATURE SURVEY (NATIVE EDITABLE TABLE)
    # ----------------------------------------------------
    print("Building Slide 4: Literature Survey (Native Table)...")
    s4 = prs.slides[3]
    remove_all_shapes(s4)
    add_college_header(s4, 4)
    add_slide_title(
        s4,
        badge_text="ACADEMIC FOUNDATION",
        title_text="Literature Survey",
        subtitle_text="Comparative review of foundational peer-reviewed research papers on AI and machine-learning credit risk prediction."
    )
    
    # 5-Column Table
    table_shape = s4.shapes.add_table(4, 5, Inches(0.72), Inches(1.85), Inches(11.89), Inches(5.00))
    table = table_shape.table
    
    table.columns[0].width = Inches(0.75)   # S. No.
    table.columns[1].width = Inches(2.15)   # Author and Year
    table.columns[2].width = Inches(3.25)   # Research Paper
    table.columns[3].width = Inches(2.74)   # Method / Focus
    table.columns[4].width = Inches(3.00)   # Key Findings
    
    headers_s4 = ["S. No.", "Author and Year", "Research Paper", "Method / Focus", "Key Findings"]
    for c_idx, h_text in enumerate(headers_s4):
        format_cell(table.cell(0, c_idx), h_text, bold=True, font_size=11, 
                    text_color=COLOR_WHITE, bg_color=COLOR_GREEN_DARK, 
                    align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT, font_name=FONT_BOLD)
        
    papers_s4 = [
        (
            "1",
            "Turiel and Aste\n(2020)",
            "Peer-to-peer loan acceptance and default prediction with artificial intelligence\nJournal: Royal Society Open Science, 7(6), 191649\nDOI: https://doi.org/10.1098/rsos.191649",
            "Machine-learning methods (logistic regression, SVM, deep neural networks) for loan acceptance and default prediction.",
            "Demonstrated that AI and ML models effectively classify loan applications and improve default-risk prediction in peer-to-peer lending."
        ),
        (
            "2",
            "Xu, Lu and Xie\n(2021)",
            "Loan default prediction of Chinese P2P market: a machine learning methodology\nJournal: Scientific Reports, 11, 18759\nDOI: https://doi.org/10.1038/s41598-021-98361-6",
            "Machine-learning approaches (Random Forest, Gradient Boosting, Neural Networks) for loan default prediction.",
            "Analyzed borrower features in P2P lending; showed ensemble and tree-based ML methods significantly enhance default identification."
        ),
        (
            "3",
            "Alonso Robisco and\nCarbó Martínez\n(2022)",
            "Measuring the model risk-adjusted performance of machine learning algorithms in credit default prediction\nJournal: Financial Innovation, 8, 70\nDOI: https://doi.org/10.1186/s40854-022-00366-1",
            "Credit-default prediction, machine-learning performance and model risk assessment.",
            "Assessed ML performance under model risk constraints; emphasized balancing predictive accuracy with model stability and validation."
        )
    ]
    
    for r_idx, row_data in enumerate(papers_s4):
        row_bg = COLOR_ALT_ROW if r_idx % 2 == 0 else COLOR_WHITE
        for c_idx, val in enumerate(row_data):
            format_cell(table.cell(r_idx + 1, c_idx), val, bold=(c_idx == 0 or c_idx == 1),
                        font_size=10, text_color=COLOR_GREEN_DARK if c_idx == 0 else COLOR_TEXT_MAIN,
                        bg_color=row_bg, align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT)

    # ----------------------------------------------------
    # SLIDE 5: EXISTING SYSTEM
    # ----------------------------------------------------
    print("Building Slide 5: Existing System...")
    s5 = prs.slides[4]
    remove_all_shapes(s5)
    add_college_header(s5, 5)
    add_slide_title(
        s5,
        badge_text="STATUS QUO",
        title_text="Existing System",
        subtitle_text="Conventional loan processing relies on sequential, manual stages across separate departments, creating operational overhead."
    )
    
    top_y = Inches(1.85)
    card_w = Inches(5.82)
    card_h = Inches(3.05)
    
    create_card(s5, Inches(0.72), top_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_el = s5.shapes.add_textbox(Inches(0.92), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_el = tb_el.text_frame
    tf_el.word_wrap = True
    
    p = tf_el.paragraphs[0]
    p.text = "Conventional Processing Stages"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    
    pts_el = [
        ("Application Submission: ", "Applicants submit personal, financial, and employment information through manual paper forms or static portals."),
        ("Document Collection: ", "Physical and digital documents (pay slips, tax returns, bank statements) are manually collected and examined."),
        ("Eligibility Assessment: ", "Loan officers manually verify borrower eligibility and calculate debt-to-income against rigid financial criteria.")
    ]
    for bold_pfx, rest in pts_el:
        p = tf_el.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_pfx
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_GREEN_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_MAIN
        
    create_card(s5, Inches(6.78), top_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
    tb_er = s5.shapes.add_textbox(Inches(6.98), Inches(1.98), Inches(5.42), Inches(2.80))
    tf_er = tb_er.text_frame
    tf_er.word_wrap = True
    
    p = tf_er.paragraphs[0]
    p.text = "Verification & Departmental Reviews"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    
    pts_er = [
        ("Credit & Identity Checks: ", "Credit bureau scores and identity credentials are queried as separate, uncoordinated administrative tasks."),
        ("Departmental Reviews: ", "Different verification stages require separate departmental reviews, manual escalations, and cross-team handoffs."),
        ("Final Loan Decision: ", "Final loan approval or rejection decisions are communicated to applicants after notable processing turnaround times.")
    ]
    for bold_pfx, rest in pts_er:
        p = tf_er.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = "• " + bold_pfx
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_GREEN_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = COLOR_TEXT_MAIN

    # Bottom Workflow Box
    flow_y = Inches(5.10)
    flow_h = Inches(1.85)
    create_card(s5, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
    
    lbl_box = s5.shapes.add_textbox(Inches(0.72), Inches(5.18), Inches(11.89), Inches(0.28))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "CONVENTIONAL LOAN PROCESSING WORKFLOW"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    p.alignment = PP_ALIGN.CENTER
    
    steps_s5 = ["Applicant", "Application Submission", "Document Verification", "Eligibility Assessment", "Risk Assessment", "Loan Decision"]
    step_w = Inches(1.60)
    step_h = Inches(0.85)
    arrow_w = Inches(0.35)
    step_y = Inches(5.68)
    start_x = Inches(0.92)
    
    for i, step_name in enumerate(steps_s5):
        cur_x = start_x + i * (step_w + arrow_w)
        add_step_card(s5, step_name, cur_x, step_y, step_w, step_h, step_num=i+1)
        if i < len(steps_s5) - 1:
            add_arrow_label(s5, cur_x + step_w, step_y, arrow_w, step_h)

    # ----------------------------------------------------
    # SLIDE 6: EXISTING SYSTEM: ADVANTAGES AND DISADVANTAGES
    # ----------------------------------------------------
    print("Building Slide 6: Existing System: Advantages and Disadvantages...")
    s6 = prs.slides[5]
    remove_all_shapes(s6)
    add_college_header(s6, 6)
    add_slide_title(
        s6,
        badge_text="COMPARATIVE EVALUATION",
        title_text="Existing System: Advantages and Disadvantages",
        subtitle_text="Objective evaluation of traditional loan underwriting procedures, highlighting established institutional strengths alongside procedural bottlenecks."
    )
    
    col_w = Inches(5.82)
    col_h = Inches(4.90)
    top_y = Inches(1.85)
    
    # Left Column: Advantages (Green styling)
    create_card(s6, Inches(0.72), top_y, col_w, col_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER, border_width=1.5)
    
    hdr_adv = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(2.00), Inches(2.0), Inches(0.35))
    hdr_adv.fill.solid()
    hdr_adv.fill.fore_color.rgb = COLOR_GREEN_LIGHT
    hdr_adv.line.color.rgb = COLOR_GREEN_BORDER
    hdr_adv.line.width = Pt(1.0)
    p = hdr_adv.text_frame.paragraphs[0]
    p.text = "✓  ADVANTAGES"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    p.alignment = PP_ALIGN.CENTER
    
    tb_adv = s6.shapes.add_textbox(Inches(0.95), Inches(2.50), Inches(5.35), Inches(4.00))
    tf_adv = tb_adv.text_frame
    tf_adv.word_wrap = True
    
    adv_items = [
        ("Established lending procedures.", "Built on established lending procedures with long-standing regulatory familiarity and standardized compliance."),
        ("Human review for complex applications.", "Direct availability of human judgment for complex, edge-case, or non-standard loan applications."),
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
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_GREEN_DARK
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT_MAIN
        
    # Right Column: Disadvantages (Red/Rose styling)
    create_card(s6, Inches(6.78), top_y, col_w, col_h, bg_color=COLOR_CARD_BG, border_color=COLOR_RED_BORDER, border_width=1.5)
    
    hdr_dis = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.01), Inches(2.00), Inches(2.2), Inches(0.35))
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
    
    tb_dis = s6.shapes.add_textbox(Inches(7.01), Inches(2.50), Inches(5.35), Inches(4.00))
    tf_dis = tb_dis.text_frame
    tf_dis.word_wrap = True
    
    dis_items = [
        ("Multiple manual processing stages.", "Multi-stage manual handoffs significantly increase overall application turnaround times."),
        ("Repeated document handling.", "Repeated document handling and redundant verification steps across disparate departments."),
        ("Coordination between departments.", "Coordination friction between separate review teams, branch managers, and credit underwriting desks."),
        ("Delays during application review.", "Delays frequently occur during application review, discrepancy resolution, and manual escalation."),
        ("Limited automation in some workflows.", "Limited automation in routine verification and assessment workflows leads to higher processing effort.")
    ]
    for idx, (bold_pfx, rest) in enumerate(dis_items):
        p = tf_dis.paragraphs[0] if idx == 0 else tf_dis.add_paragraph()
        if idx > 0:
            p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"{idx+1}. {bold_pfx} "
        r1.font.name = FONT_BOLD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_RED_TEXT
        r2 = p.add_run()
        r2.text = rest
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT_MAIN

    # ----------------------------------------------------
    # SLIDE 7: PROPOSED SYSTEM
    # ----------------------------------------------------
    print("Building Slide 7: Proposed System...")
    s7 = prs.slides[6]
    remove_all_shapes(s7)
    add_college_header(s7, 7)
    add_slide_title(
        s7,
        badge_text="PROPOSED SYSTEM",
        title_text="ARMEK – Agentic AI-Powered Loan Underwriting and Risk Assessment System",
        subtitle_text="A multi-agent loan underwriting architecture coordinating specialized worker agents through a central Master Agent, backed by modern web and AI technologies."
    )
    
    pillar_w = Inches(2.82)
    pillar_h = Inches(2.35)
    top_p_y = Inches(1.85)
    
    pillars = [
        ("React.js Frontend", "Applicant Interface", "React-based frontend for applicant interaction, structured information input, conversational form submission, and real-time status display."),
        ("FastAPI Backend", "Application Processing", "Python FastAPI backend for application processing, session management, REST endpoints, and Pydantic-validated data models."),
        ("OpenAI Integration", "GPT-4o-mini", "OpenAI GPT-4o-mini integration for AI-enabled conversational assistance, applicant guidance, and structured parameter reasoning."),
        ("Specialized Worker Agents", "Modular Verification", "Dedicated agents executing eligibility checks, state reconciliation, credit assessment, fraud checks, and sanction document generation.")
    ]
    
    for idx, (title, sub, desc) in enumerate(pillars):
        px = Inches(0.72) + idx * Inches(3.02)
        create_card(s7, px, top_p_y, pillar_w, pillar_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
        
        tb = s7.shapes.add_textbox(px + Inches(0.12), top_p_y + Inches(0.12), pillar_w - Inches(0.24), pillar_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_GREEN_TEXT
        
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(2)
        p_sub.text = sub
        p_sub.font.name = FONT_BOLD
        p_sub.font.size = Pt(9.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_GREEN_PRIMARY
        
        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(4)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10)
        p_desc.font.color.rgb = COLOR_TEXT_MAIN

    # Bottom Section: End-to-End Orchestrated Workflow
    flow_y = Inches(4.45)
    flow_h = Inches(2.50)
    create_card(s7, Inches(0.72), flow_y, Inches(11.89), flow_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_PRIMARY, border_width=1.5)
    
    lbl_box = s7.shapes.add_textbox(Inches(0.72), Inches(4.52), Inches(11.89), Inches(0.28))
    p = lbl_box.text_frame.paragraphs[0]
    p.text = "END-TO-END ORCHESTRATED WORKFLOW PIPELINE"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
    p.alignment = PP_ALIGN.CENTER
    
    # Row 1: Applicant -> React Frontend -> FastAPI Backend -> AI Integration
    w_r1 = Inches(2.45)
    h_r1 = Inches(0.68)
    arr_w = Inches(0.40)
    r1_y = Inches(4.90)
    start_r1_x = Inches(1.10)
    
    r1_steps = ["Applicant", "React Frontend", "FastAPI Backend", "AI Integration"]
    for i, s_text in enumerate(r1_steps):
        cx = start_r1_x + i * (w_r1 + arr_w)
        add_step_card(s7, s_text, cx, r1_y, w_r1, h_r1)
        if i < len(r1_steps) - 1:
            add_arrow_label(s7, cx + w_r1, r1_y, arr_w, h_r1)
            
    # Connector down to Master Agent
    arr_down = s7.shapes.add_textbox(Inches(9.65), Inches(5.50), Inches(1.0), Inches(0.35))
    p_d = arr_down.text_frame.paragraphs[0]
    p_d.text = "▼"
    p_d.alignment = PP_ALIGN.CENTER
    p_d.font.name = FONT_BOLD
    p_d.font.size = Pt(14)
    p_d.font.color.rgb = COLOR_GREEN_PRIMARY
    
    # Row 2: Master Agent -> Specialized Worker Agents -> Assessment & Documents -> Status & Results
    r2_y = Inches(5.90)
    add_step_card(s7, "★ MASTER AGENT ★\n(Central Coordination)", Inches(1.10), r2_y, Inches(3.20), Inches(0.85), is_highlighted=True)
    add_arrow_label(s7, Inches(4.35), r2_y, arr_w, Inches(0.85))
    
    add_step_card(s7, "Specialized Worker Agents\n(Eligibility · KYC · Bureau · Fraud)", Inches(4.80), r2_y, Inches(3.10), Inches(0.85))
    add_arrow_label(s7, Inches(7.95), r2_y, arr_w, Inches(0.85))
    
    add_step_card(s7, "Assessment & Docs\n(ReportLab Sanction Letter)", Inches(8.40), r2_y, Inches(2.20), Inches(0.85))
    add_arrow_label(s7, Inches(10.65), r2_y, Inches(0.30), Inches(0.85))
    
    add_step_card(s7, "Application\nStatus & Results", Inches(11.00), r2_y, Inches(1.35), Inches(0.85))

    # ----------------------------------------------------
    # SLIDE 8: PROPOSED SYSTEM: ADVANTAGES
    # ----------------------------------------------------
    print("Building Slide 8: Proposed System: Advantages...")
    s8 = prs.slides[7]
    remove_all_shapes(s8)
    add_college_header(s8, 8)
    add_slide_title(
        s8,
        badge_text="SYSTEM BENEFITS",
        title_text="Proposed System: Advantages",
        subtitle_text="Structural and operational improvements delivered by ARMEK's orchestrated multi-agent underwriting architecture."
    )
    
    adv_6 = [
        ("Organized multi-agent workflow.", "Organized multi-agent workflow distributing assessment across specialized agents, preventing processing bottlenecks in any single stage."),
        ("Reduced manual coordination.", "Reduced repetitive manual coordination steps through automated agent handoffs orchestrated by the central Master Agent."),
        ("Centralized information processing.", "Centralized applicant information processing in a unified pipeline, eliminating data duplication and state inconsistency across stages."),
        ("Conversational AI-based interaction.", "Conversational AI-based interaction powered by GPT-4o-mini, providing an intuitive and accessible experience for loan applicants."),
        ("Separate responsibilities for worker agents.", "Separate responsibilities for specialized worker agents (eligibility, reconciliation, credit checks, fraud detection, and document generation)."),
        ("Clear application status and document workflow.", "Clear application status and document generation workflow, providing real-time visibility and automated sanction letter generation upon approval.")
    ]
    
    row_ys = [Inches(1.85), Inches(3.55), Inches(5.25)]
    card_w = Inches(5.82)
    card_h = Inches(1.50)
    
    for idx, (title, desc) in enumerate(adv_6):
        col_x = Inches(0.72) if idx % 2 == 0 else Inches(6.78)
        cur_y = row_ys[idx // 2]
        
        create_card(s8, col_x, cur_y, card_w, card_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
        
        # Green icon pill / check badge
        chk = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.18), cur_y + Inches(0.18), Inches(0.35), Inches(0.35))
        chk.fill.solid()
        chk.fill.fore_color.rgb = COLOR_GREEN_LIGHT
        chk.line.color.rgb = COLOR_GREEN_BORDER
        p_c = chk.text_frame.paragraphs[0]
        p_c.text = "✓"
        p_c.font.name = FONT_BOLD
        p_c.font.size = Pt(12)
        p_c.font.bold = True
        p_c.font.color.rgb = COLOR_GREEN_DARK
        p_c.alignment = PP_ALIGN.CENTER
        
        tb = s8.shapes.add_textbox(col_x + Inches(0.65), cur_y + Inches(0.12), card_w - Inches(0.78), card_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_GREEN_TEXT
        
        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(3)
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = COLOR_TEXT_MAIN

    # ----------------------------------------------------
    # SLIDE 9: SYSTEM ARCHITECTURE DIAGRAM
    # ----------------------------------------------------
    print("Building Slide 9: System Architecture Diagram...")
    s9 = prs.slides[8]
    remove_all_shapes(s9)
    add_college_header(s9, 9)
    add_slide_title(
        s9,
        badge_text="TECHNICAL DESIGN",
        title_text="System Architecture Diagram",
        subtitle_text="Five-tier modular architecture illustrating data flow, agent orchestration, and component responsibilities in ARMEK."
    )
    
    layer_w = Inches(11.89)
    lx = Inches(0.72)
    
    # Layer 1 – User Interface
    l1_y = Inches(1.85)
    l1_h = Inches(0.72)
    create_card(s9, lx, l1_y, layer_w, l1_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
    tb1 = s9.shapes.add_textbox(lx + Inches(0.20), l1_y + Inches(0.10), layer_w - Inches(0.40), l1_h - Inches(0.20))
    p = tb1.text_frame.paragraphs[0]
    p.text = "Layer 1 – User Interface: Applicant  ·  React.js  ·  Bootstrap  ·  Axios"
    p.font.name = FONT_BOLD
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    p.alignment = PP_ALIGN.CENTER
    
    add_arrow_label(s9, Inches(6.30), Inches(2.58), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 2 – Backend
    l2_y = Inches(2.80)
    l2_h = Inches(0.72)
    create_card(s9, lx, l2_y, layer_w, l2_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
    tb2 = s9.shapes.add_textbox(lx + Inches(0.20), l2_y + Inches(0.10), layer_w - Inches(0.40), l2_h - Inches(0.20))
    p = tb2.text_frame.paragraphs[0]
    p.text = "Layer 2 – Backend: Python FastAPI  ·  API endpoints  ·  Pydantic validation"
    p.font.name = FONT_BOLD
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    p.alignment = PP_ALIGN.CENTER
    
    add_arrow_label(s9, Inches(6.30), Inches(3.53), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 3 – AI Orchestration (PROMINENT CENTER)
    l3_y = Inches(3.75)
    l3_h = Inches(0.85)
    create_card(s9, lx, l3_y, layer_w, l3_h, bg_color=COLOR_GREEN_DARK, border_color=COLOR_GREEN_PRIMARY, border_width=2.0)
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
    p_sub.text = "Coordinates the specialized worker agents, manages session state, and evaluates decision conditions"
    p_sub.font.name = FONT_BODY
    p_sub.font.size = Pt(10)
    p_sub.font.color.rgb = COLOR_GREEN_LIGHT
    p_sub.alignment = PP_ALIGN.CENTER
    
    add_arrow_label(s9, Inches(6.30), Inches(4.61), Inches(0.72), Inches(0.22), text="▼")
    
    # Layer 4 – Worker Agent Processing (5 Specialized Worker Agents)
    l4_y = Inches(4.82)
    l4_h = Inches(1.10)
    create_card(s9, lx, l4_y, layer_w, l4_h, bg_color=COLOR_CARD_BG, border_color=COLOR_GREEN_BORDER)
    
    lbl_l4 = s9.shapes.add_textbox(lx, l4_y + Inches(0.04), layer_w, Inches(0.25))
    p = lbl_l4.text_frame.paragraphs[0]
    p.text = "Layer 4 – Worker Agent Processing (Master Agent Coordinates the Five Worker Agents)"
    p.font.name = FONT_BOLD
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_DARK
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
        create_card(s9, wx, l4_y + Inches(0.32), w_card_w, w_card_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
        tb_w = s9.shapes.add_textbox(wx, l4_y + Inches(0.35), w_card_w, w_card_h - Inches(0.06))
        tf_w = tb_w.text_frame
        tf_w.word_wrap = True
        
        p = tf_w.paragraphs[0]
        p.text = w_name
        p.font.name = FONT_BOLD
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GREEN_DARK
        p.alignment = PP_ALIGN.CENTER
        
        p2 = tf_w.add_paragraph()
        p2.space_before = Pt(1)
        p2.text = w_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(8.5)
        p2.font.color.rgb = COLOR_TEXT_MUTED
        p2.alignment = PP_ALIGN.CENTER

    add_arrow_label(s9, Inches(6.30), Inches(5.93), Inches(0.72), Inches(0.18), text="▼")
    
    # Layer 5 – Output
    l5_y = Inches(6.10)
    l5_h = Inches(0.55)
    create_card(s9, lx, l5_y, layer_w, l5_h, bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_BORDER)
    tb5 = s9.shapes.add_textbox(lx + Inches(0.20), l5_y + Inches(0.06), layer_w - Inches(0.40), l5_h - Inches(0.12))
    p = tb5.text_frame.paragraphs[0]
    p.text = "Layer 5 – Output: Loan assessment result  ·  Loan-related documents  ·  Application status"
    p.font.name = FONT_BOLD
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_TEXT
    p.alignment = PP_ALIGN.CENTER
    
    # Simulated Verification Note Box
    note_box = create_card(s9, lx, Inches(6.72), layer_w, Inches(0.34), 
                           bg_color=RGBColor(254, 252, 232), border_color=RGBColor(250, 204, 21), border_width=1.0)
    tb_n = s9.shapes.add_textbox(lx + Inches(0.20), Inches(6.73), layer_w - Inches(0.40), Inches(0.30))
    p = tb_n.text_frame.paragraphs[0]
    p.text = "Note: Credit-related assessment and bureau verification are simulated internally. The system does not connect to live external government or credit-bureau APIs."
    p.font.name = FONT_BODY
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = RGBColor(133, 77, 14)
    p.alignment = PP_ALIGN.CENTER

    # ----------------------------------------------------
    # SLIDE 10: REFERENCES (NATIVE EDITABLE TABLE)
    # ----------------------------------------------------
    print("Building Slide 10: References (Native Table)...")
    s10 = prs.slides[9]
    remove_all_shapes(s10)
    add_college_header(s10, 10)
    add_slide_title(
        s10,
        badge_text="ACADEMIC & TECHNICAL SOURCES",
        title_text="References",
        subtitle_text="Comprehensive research literature and official technical documentation supporting the ARMEK loan underwriting system."
    )
    
    table_shape_10 = s10.shapes.add_table(8, 5, Inches(0.72), Inches(1.85), Inches(11.89), Inches(5.10))
    table_10 = table_shape_10.table
    
    table_10.columns[0].width = Inches(0.70)   # S. No.
    table_10.columns[1].width = Inches(2.20)   # Author(s) and Year
    table_10.columns[2].width = Inches(3.40)   # Paper / Documentation Title
    table_10.columns[3].width = Inches(2.40)   # Journal / Source
    table_10.columns[4].width = Inches(3.19)   # DOI / URL
    
    headers_s10 = ["S. No.", "Author(s) and Year", "Paper / Documentation Title", "Journal / Source", "DOI / URL"]
    for c_idx, h_text in enumerate(headers_s10):
        format_cell(table_10.cell(0, c_idx), h_text, bold=True, font_size=11, 
                    text_color=COLOR_WHITE, bg_color=COLOR_GREEN_DARK, 
                    align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT, font_name=FONT_BOLD)
        
    refs_s10 = [
        (
            "1",
            "Turiel and Aste\n(2020)",
            "Peer-to-peer loan acceptance and default prediction with artificial intelligence",
            "Royal Society Open Science, 7(6), 191649",
            "https://doi.org/10.1098/rsos.191649"
        ),
        (
            "2",
            "Xu, Lu and Xie\n(2021)",
            "Loan default prediction of Chinese P2P market: a machine learning methodology",
            "Scientific Reports, 11, 18759",
            "https://doi.org/10.1038/s41598-021-98361-6"
        ),
        (
            "3",
            "Alonso Robisco and\nCarbó Martínez (2022)",
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
                        font_size=9.5, text_color=COLOR_GREEN_DARK if c_idx == 0 else COLOR_TEXT_MAIN,
                        bg_color=row_bg, align=PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT)

    # ----------------------------------------------------
    # SLIDE 11: CONCLUSION
    # ----------------------------------------------------
    print("Building Slide 11: Conclusion...")
    s11 = prs.slides[10]
    remove_all_shapes(s11)
    add_college_header(s11, 11)
    add_slide_title(
        s11,
        badge_text="PROJECT SUMMARY",
        title_text="Conclusion",
        subtitle_text="Summary of system achievements, multi-agent workflow benefits, and project outcomes."
    )
    
    # Exact required conclusion paragraph
    conclusion_text = (
        "ARMEK is an AI-powered loan underwriting system that makes the loan application and assessment "
        "process more organized. It uses a Master Agent to coordinate different worker agents for eligibility "
        "checking, verification, fraud detection and risk assessment. The system combines React.js, FastAPI "
        "and OpenAI GPT-4o-mini to support applicant interaction and loan processing. Overall, ARMEK "
        "demonstrates how multi-agent AI can help simplify loan assessment through a structured workflow."
    )
    
    # Central clean presentation card
    create_card(s11, Inches(0.72), Inches(2.20), Inches(11.89), Inches(3.20), bg_color=COLOR_GREEN_LIGHT, border_color=COLOR_GREEN_PRIMARY, border_width=1.5)
    
    tb_c = s11.shapes.add_textbox(Inches(1.10), Inches(2.55), Inches(11.13), Inches(2.50))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    
    p = tf_c.paragraphs[0]
    p.text = conclusion_text
    p.font.name = FONT_BODY
    p.font.size = Pt(16)
    p.font.color.rgb = COLOR_GREEN_TEXT
    p.line_spacing = 1.35
    p.alignment = PP_ALIGN.CENTER

    # Save to ARMEK_Final_Project_Review_Updated.pptx
    print(f"\nSaving final presentation to: {output_file}")
    prs.save(output_file)
    print("SUCCESS: Presentation saved successfully!")

if __name__ == "__main__":
    build_presentation()

