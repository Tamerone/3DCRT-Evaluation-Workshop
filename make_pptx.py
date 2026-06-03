"""
3DCRT Workshop Deck v2 — PowerPoint generator
Run: python3 make_pptx.py
Output: 3DCRT_Workshop_Deck_v2.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import pptx.util as util
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Colour palette ──────────────────────────────────────────────
C_NAVY      = RGBColor(0x0E, 0x28, 0x41)
C_DARK_NAVY = RGBColor(0x0a, 0x1f, 0x33)
C_CYAN      = RGBColor(0x0F, 0x9E, 0xD5)
C_ORANGE    = RGBColor(0xE9, 0x71, 0x32)
C_GREEN     = RGBColor(0x29, 0xA3, 0x5B)
C_RED       = RGBColor(0xD4, 0x40, 0x40)
C_WHITE     = RGBColor(0xE8, 0xED, 0xF2)
C_DIM       = RGBColor(0x9A, 0xAC, 0xBE)
C_YELLOW    = RGBColor(0xF5, 0xC5, 0x18)

# ── Slide dimensions: 1920×1080 px ≈ 33.87×19.05 cm ─────────────
W = Cm(33.87)
H = Cm(19.05)

def new_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def blank_layout(prs):
    return prs.slide_layouts[6]   # blank

def set_bg(slide, color):
    """Fill slide background with a solid colour."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, left, top, width, height, text,
                font_size=18, bold=False, color=C_WHITE,
                align=PP_ALIGN.LEFT, italic=False, font_name="Calibri",
                word_wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.name  = font_name
    run.font.italic = italic
    return txBox

def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=0):
    shape = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO_SHAPE if False else 1,  # 1 = rectangle
        left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def add_slide_number(slide, slide_num):
    """Add slide number in bottom-right corner."""
    add_textbox(slide,
        left=W - Cm(3), top=H - Cm(1.0),
        width=Cm(2.5), height=Cm(0.7),
        text=str(slide_num),
        font_size=10, color=C_DIM,
        align=PP_ALIGN.RIGHT)

def add_footer_bar(slide):
    """Thin cyan line at bottom."""
    add_rect(slide,
        left=0, top=H - Cm(0.25),
        width=W, height=Cm(0.25),
        fill_color=C_CYAN)

def make_title_slide(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_footer_bar(slide)

    # Decorative left accent bar
    add_rect(slide, left=0, top=0, width=Cm(0.6), height=H, fill_color=C_CYAN)

    # Main title
    add_textbox(slide,
        left=Cm(2), top=Cm(4.5), width=Cm(28), height=Cm(4),
        text="What Makes a Plan Good?",
        font_size=54, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)

    # Subtitle
    add_textbox(slide,
        left=Cm(2), top=Cm(9.2), width=Cm(28), height=Cm(2),
        text="3DCRT Plan Evaluation Workshop — Theory & Practical Module",
        font_size=22, bold=False, color=C_CYAN, align=PP_ALIGN.LEFT)

    # Cyan rule
    add_rect(slide, left=Cm(2), top=Cm(11.5), width=Cm(12), height=Cm(0.1), fill_color=C_CYAN)

    # Author
    add_textbox(slide,
        left=Cm(2), top=Cm(12.0), width=Cm(28), height=Cm(1.5),
        text="Tamerone Manasse  ·  UCT/GSH Registrars 2026",
        font_size=16, color=C_DIM, align=PP_ALIGN.LEFT)

    add_slide_number(slide, slide_num)
    return slide

def make_session_map(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "Session Map", font_size=32, bold=True, color=C_CYAN)

    add_textbox(slide, Cm(1.5), Cm(2.0), Cm(30), Cm(0.9),
        "2 Hours · 5 Cases · One framework.",
        font_size=18, bold=False, color=C_WHITE)

    rows = [
        ("0:00", "Theory",              "18 min",  C_ORANGE),
        ("0:18", "Case 1: Breast",      "22 min",  C_CYAN),
        ("0:40", "Case 2: Gynaecology", "10 min",  C_CYAN),
        ("0:50", "Case 3: Prostate",    "10 min",  C_CYAN),
        ("1:00", "Case 4: Pituitary",   "10 min",  C_CYAN),
        ("1:10", "Case 5: H&N",         "10 min",  C_CYAN),
        ("1:20", "Buffer",              "30 min",  C_DIM),
        ("1:50", "Wrap-up",             "10 min",  C_GREEN),
    ]

    y = Cm(3.2)
    row_h = Cm(1.4)
    for i, (time, label, dur, col) in enumerate(rows):
        bg = C_DARK_NAVY if i % 2 == 0 else RGBColor(0x0c, 0x25, 0x3d)
        add_rect(slide, Cm(1.5), y, Cm(28), row_h, fill_color=bg)
        add_textbox(slide, Cm(1.7), y + Cm(0.15), Cm(3), row_h - Cm(0.3),
            time, font_size=14, bold=True, color=col, font_name="Courier New")
        add_textbox(slide, Cm(5.5), y + Cm(0.15), Cm(19), row_h - Cm(0.3),
            label, font_size=14, bold=False, color=C_WHITE)
        add_textbox(slide, Cm(25), y + Cm(0.15), Cm(4), row_h - Cm(0.3),
            dur, font_size=14, bold=False, color=C_DIM, align=PP_ALIGN.RIGHT)
        y += row_h

    # Learning outcomes
    add_textbox(slide, Cm(1.5), y + Cm(0.3), Cm(28), Cm(2.2),
        "By the end you can:  Apply FCB-CHOPS · Read a DVH · Explain CI & HI · Apply OAR constraints · Triage a plan",
        font_size=13, bold=False, color=C_DIM, italic=True)

    add_slide_number(slide, slide_num)
    return slide

def make_four_questions(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "The Four Central Questions", font_size=32, bold=True, color=C_CYAN)

    add_textbox(slide, Cm(1.5), Cm(2.0), Cm(30), Cm(0.9),
        "A good plan answers four questions — in order.",
        font_size=18, color=C_WHITE)

    qs = [
        ("Q1", "Does it treat the target?",           C_GREEN),
        ("Q2", "Does it spare normal tissue?",        C_CYAN),
        ("Q3", "Is it technically sound?",            C_ORANGE),
        ("Q4", "Is it deliverable & robust?",         C_DIM),
    ]

    y = Cm(3.3)
    for qn_str, text, col in qs:
        add_rect(slide, Cm(1.5), y, Cm(28), Cm(2.2), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.5), Cm(2.2), fill_color=col)
        add_textbox(slide, Cm(2.4), y + Cm(0.1), Cm(5), Cm(2.0),
            qn_str, font_size=26, bold=True, color=col)
        add_textbox(slide, Cm(5.5), y + Cm(0.3), Cm(23), Cm(1.6),
            text, font_size=22, bold=False, color=C_WHITE)
        y += Cm(2.5)

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"A plan is never perfect — it is the best defensible compromise."',
        font_size=14, italic=True, color=C_DIM, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_plan_quality_spectrum(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "Plan Quality Spectrum", font_size=32, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(1.5), Cm(2.0), Cm(30), Cm(0.9),
        "Bad, acceptable, trade-off — or good?",
        font_size=18, color=C_WHITE)

    items = [
        ("CLEARLY BAD",  "Unsafe / unacceptable",                                 C_RED),
        ("ACCEPTABLE",   "Passes constraints but better achievable",               C_YELLOW),
        ("TRADE-OFF",    "Competing priorities — documented clinical decision",    C_ORANGE),
        ("GOOD",         "Balanced, robust, documented",                           C_GREEN),
    ]

    x_positions = [Cm(1.5), Cm(8.5), Cm(16.5), Cm(24.5)]
    box_w = Cm(6.5)
    box_top = Cm(3.2)
    box_h = Cm(7.5)

    for i, ((label, desc, col), xpos) in enumerate(zip(items, x_positions)):
        add_rect(slide, xpos, box_top, box_w, box_h, fill_color=C_DARK_NAVY)
        add_rect(slide, xpos, box_top, box_w, Cm(0.5), fill_color=col)
        add_textbox(slide, xpos + Cm(0.2), box_top + Cm(0.7), box_w - Cm(0.4), Cm(1.8),
            label, font_size=16, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_textbox(slide, xpos + Cm(0.2), box_top + Cm(2.6), box_w - Cm(0.4), Cm(4.5),
            desc, font_size=13, color=C_WHITE, align=PP_ALIGN.CENTER, word_wrap=True)

    add_rect(slide, Cm(1.5), Cm(11.5), Cm(30), Cm(1.8), fill_color=RGBColor(0x12, 0x32, 0x52))
    add_textbox(slide, Cm(1.7), Cm(11.6), Cm(29.6), Cm(1.6),
        'GUT-CHECK: "Would I confidently treat my own family member with this plan?"',
        font_size=15, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_fcbchops_evolution(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "FCB-CHOPS Evolution", font_size=32, bold=True, color=C_CYAN)

    add_rect(slide, Cm(1.5), Cm(2.5), Cm(13.5), Cm(6.5), fill_color=C_DARK_NAVY)
    add_textbox(slide, Cm(1.7), Cm(2.7), Cm(13), Cm(1.0),
        "CB-CHOP (2017)", font_size=20, bold=True, color=C_DIM)
    add_textbox(slide, Cm(1.7), Cm(3.8), Cm(13), Cm(4.5),
        "Jimenez et al.\nOriginal six-checkpoint framework\nfor 3DCRT plan evaluation",
        font_size=15, color=C_WHITE, word_wrap=True)

    add_rect(slide, Cm(16.5), Cm(2.5), Cm(15), Cm(6.5), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(16.5), Cm(2.5), Cm(15), Cm(0.4), fill_color=C_CYAN)
    add_textbox(slide, Cm(16.7), Cm(3.0), Cm(14.5), Cm(1.0),
        "FCB-CHOPS (2024)", font_size=20, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(16.7), Cm(4.1), Cm(14.5), Cm(4.5),
        "Weisman et al.  ·  PMID 40017913\n\nAdded:  F (Fusion/Imaging)\n          S (Summation)",
        font_size=15, color=C_WHITE, word_wrap=True)

    # Arrow
    add_textbox(slide, Cm(14.5), Cm(4.8), Cm(2.5), Cm(1.2),
        "→", font_size=48, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    # Letters display
    letters = [("F", C_ORANGE), ("C", C_WHITE), ("B", C_WHITE),
               ("–", C_DIM), ("C", C_WHITE), ("H", C_WHITE),
               ("O", C_WHITE), ("P", C_WHITE), ("S", C_ORANGE)]
    x = Cm(3)
    for letter, col in letters:
        add_textbox(slide, x, Cm(10.0), Cm(3.0), Cm(2.5),
            letter, font_size=44, bold=True, color=col, align=PP_ALIGN.CENTER)
        x += Cm(3.2)

    add_textbox(slide, Cm(1.5), Cm(12.5), Cm(10), Cm(1.0),
        "F · S = NEW additions", font_size=13, color=C_ORANGE, italic=True)

    add_slide_number(slide, slide_num)
    return slide

def make_fcbchops_framework(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "FCB-CHOPS Framework", font_size=32, bold=True, color=C_CYAN)

    checkpoints = [
        ("F", "Fusion / Imaging",      "Correct image registration, tumour delineation from best imaging"),
        ("C", "Contours",              "Target volumes and OAR delineation accurate, margins appropriate"),
        ("B", "Beam Arrangement",      "Isocentre, geometry, MLC shaping, technical quality"),
        ("C", "Coverage",              "Target coverage metrics: D95, D2, HI, CI, cold spots"),
        ("H", "Heterogeneity/Hotspots","Hotspot location — inside PTV? D2% <107%?"),
        ("O", "Organs at Risk",        "Serial vs parallel; QUANTEC limits; PRV applied"),
        ("P", "Prescription",          "Dose/fractionation correct; normalisation point valid; laterality"),
        ("S", "Summation",             "Prior RT, cumulative dose, brachytherapy phases"),
    ]

    y = Cm(2.2)
    for i, (letter, name, desc) in enumerate(checkpoints):
        bg = C_DARK_NAVY if i % 2 == 0 else RGBColor(0x0c, 0x25, 0x3d)
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(1.7), fill_color=bg)
        add_textbox(slide, Cm(1.7), y + Cm(0.1), Cm(1.2), Cm(1.5),
            letter, font_size=20, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
        add_textbox(slide, Cm(3.2), y + Cm(0.1), Cm(8), Cm(1.5),
            name, font_size=15, bold=True, color=C_WHITE)
        add_textbox(slide, Cm(11.5), y + Cm(0.1), Cm(19.5), Cm(1.5),
            desc, font_size=13, color=C_DIM)
        y += Cm(1.75)

    add_textbox(slide, Cm(1.5), H - Cm(2.2), Cm(18), Cm(0.8),
        '"Start with Prescription."', font_size=13, italic=True, color=C_ORANGE)
    add_textbox(slide, Cm(15), H - Cm(2.2), Cm(17), Cm(0.8),
        '"Letters are a safety net, not a script."', font_size=13, italic=True, color=C_DIM)

    add_slide_number(slide, slide_num)
    return slide

def make_block_b_beams(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_rect(slide, Cm(1.5), Cm(0.5), Cm(6), Cm(1.5), fill_color=C_ORANGE)
    add_textbox(slide, Cm(1.7), Cm(0.55), Cm(5.5), Cm(1.4),
        "Block B: Beams & Technical Quality", font_size=18, bold=True, color=C_WHITE)

    add_textbox(slide, Cm(8.5), Cm(0.8), Cm(10), Cm(1.0),
        "⏱ 12 minutes", font_size=16, color=C_ORANGE)

    topics = [
        ("Isocentre placement",    "Central within PTV; not in lung, air, or density interface"),
        ("Beam geometry",          "Gantry/collimator angles; coplanar vs non-coplanar; BEV check"),
        ("MLC shaping",            "Adequate margin around PTV; no tight clipping"),
        ("Wedges / FiF",           "Compensate sloping surfaces; field-in-field sub-segments for hotspots"),
        ("Monitor unit check",     "MU values plausible; no anomalous field weights"),
    ]

    y = Cm(2.5)
    for title, detail in topics:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.4), fill_color=C_DARK_NAVY)
        add_textbox(slide, Cm(1.7), y + Cm(0.1), Cm(11), Cm(1.0),
            title, font_size=16, bold=True, color=C_ORANGE)
        add_textbox(slide, Cm(1.7), y + Cm(1.1), Cm(28), Cm(1.1),
            detail, font_size=14, color=C_WHITE)
        y += Cm(2.6)

    add_slide_number(slide, slide_num)
    return slide

def make_field_arrangement(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Field Arrangement Diagram", font_size=32, bold=True, color=C_CYAN)

    add_textbox(slide, Cm(1.5), Cm(1.9), Cm(30), Cm(0.9),
        "Every beam should reach the target by the most direct, safest route",
        font_size=16, color=C_WHITE, italic=True)

    scenarios = [
        ("UNAVOIDABLE",
         "4-field box — OARs in beam path\nDocument and manage",
         C_YELLOW,
         "AP + PA + 2 laterals. Rectum and bladder in beam path for pelvis.\nDocument DVH values; optimize weighting to minimize dose."),
        ("AVOIDABLE",
         "Beam transits OAR before PTV\nRe-angle the field",
         C_RED,
         "Beam enters through critical OAR on its way to target.\nRe-angle gantry to use OAR-free entry path."),
        ("SOLUTION",
         "Ipsilateral entry — OAR-free side\nBeam avoids contralateral OAR",
         C_GREEN,
         "Choose entry angle from the ipsilateral side.\nAvoids transit through uninvolved structures."),
    ]

    x = Cm(1.5)
    box_w = Cm(9.5)
    for label, headline, col, detail in scenarios:
        add_rect(slide, x, Cm(3.2), box_w, Cm(9.5), fill_color=C_DARK_NAVY)
        add_rect(slide, x, Cm(3.2), box_w, Cm(0.5), fill_color=col)
        add_textbox(slide, x + Cm(0.2), Cm(3.9), box_w - Cm(0.4), Cm(1.0),
            label, font_size=16, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + Cm(0.2), Cm(5.1), box_w - Cm(0.4), Cm(2.0),
            headline, font_size=14, bold=False, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + Cm(0.2), Cm(7.5), box_w - Cm(0.4), Cm(4.8),
            detail, font_size=12, color=C_DIM, word_wrap=True)
        x += Cm(10.5)

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"Before accepting any OAR transit — ask: is there an arrangement that avoids it?"',
        font_size=14, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_technical_quality(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "Technical Quality", font_size=32, bold=True, color=C_CYAN)

    # Red flags column
    add_textbox(slide, Cm(1.5), Cm(2.2), Cm(14), Cm(0.9),
        "Technical Red Flags", font_size=18, bold=True, color=C_RED)

    flags = [
        ("RISK",  "Isocentre in lung / air / density interface"),
        ("RISK",  "Beam transiting critical OAR"),
        ("RISK",  "Hotspot outside PTV"),
        ("MAJOR", "MLC clipped tight to PTV edge"),
        ("MAJOR", "Normalisation point in unstable tissue"),
    ]

    y = Cm(3.3)
    for tag, text in flags:
        col = C_RED if tag == "RISK" else C_ORANGE
        add_rect(slide, Cm(1.5), y, Cm(14), Cm(1.5), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.4), Cm(1.5), fill_color=col)
        add_textbox(slide, Cm(2.1), y + Cm(0.1), Cm(3), Cm(1.3),
            tag, font_size=11, bold=True, color=col)
        add_textbox(slide, Cm(5.2), y + Cm(0.1), Cm(10), Cm(1.3),
            text, font_size=13, color=C_WHITE)
        y += Cm(1.7)

    # Checks column
    add_textbox(slide, Cm(17), Cm(2.2), Cm(15), Cm(0.9),
        "Technical Checks", font_size=18, bold=True, color=C_GREEN)

    checks = [
        "Isocentre — central in PTV, not in air/lung",
        "Beam weighting — balanced; hotspot central",
        "Wedges — appropriate heel/toe; check direction",
        "Field-in-field — sub-segments block peak hotspots",
        "Monitor units — plausible; no anomalous values",
    ]

    y = Cm(3.3)
    for check in checks:
        add_rect(slide, Cm(17), y, Cm(15), Cm(1.5), fill_color=C_DARK_NAVY)
        add_textbox(slide, Cm(17.3), y + Cm(0.15), Cm(14.4), Cm(1.2),
            "✓  " + check, font_size=13, color=C_WHITE)
        y += Cm(1.7)

    add_slide_number(slide, slide_num)
    return slide

def make_hotspots_diagram(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Hot Spots — Location Matters", font_size=32, bold=True, color=C_CYAN)

    add_textbox(slide, Cm(1.5), Cm(1.9), Cm(30), Cm(0.9),
        "Hot spots, dose gradient & integral dose",
        font_size=16, color=C_WHITE)

    # Left box — inside PTV
    add_rect(slide, Cm(1.5), Cm(3.2), Cm(14), Cm(9.0), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(1.5), Cm(3.2), Cm(14), Cm(0.5), fill_color=C_GREEN)
    add_textbox(slide, Cm(1.7), Cm(3.9), Cm(13.5), Cm(1.2),
        "Hotspot INSIDE PTV", font_size=18, bold=True, color=C_GREEN)
    add_textbox(slide, Cm(1.7), Cm(5.3), Cm(13.5), Cm(6.5),
        "= ACCEPTABLE\n\n107% inside the target is expected in 3DCRT.\n\nThe prescription point is normalised to ~100%.\nA modest hotspot in the tumour itself is clinically acceptable.\n\nMonitor: ensure D2% < 107%",
        font_size=14, color=C_WHITE, word_wrap=True)

    # Right box — outside PTV
    add_rect(slide, Cm(17), Cm(3.2), Cm(14), Cm(9.0), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(17), Cm(3.2), Cm(14), Cm(0.5), fill_color=C_RED)
    add_textbox(slide, Cm(17.2), Cm(3.9), Cm(13.5), Cm(1.2),
        "Hotspot OUTSIDE PTV", font_size=18, bold=True, color=C_RED)
    add_textbox(slide, Cm(17.2), Cm(5.3), Cm(13.5), Cm(6.5),
        "= CONCERN\n\nIn OAR or skin = PROBLEM.\n\nCauses: deep isocentre, unweighted beams, density heterogeneity.\n\nAction: move iso, adjust weighting, check BEV.",
        font_size=14, color=C_WHITE, word_wrap=True)

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"The same hot spot is acceptable inside the PTV and a problem outside it"',
        font_size=14, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_block_c_coverage(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_rect(slide, Cm(1.5), Cm(0.5), Cm(6), Cm(1.5), fill_color=C_CYAN)
    add_textbox(slide, Cm(1.7), Cm(0.55), Cm(5.6), Cm(1.4),
        "Block C: Coverage", font_size=18, bold=True, color=C_DARK_NAVY)
    add_textbox(slide, Cm(8.5), Cm(0.8), Cm(10), Cm(1.0),
        "⏱ 10 minutes", font_size=16, color=C_CYAN)

    items = [
        ("Volume Hierarchy", "GTV → CTV → ITV → PTV", C_CYAN),
        ("DVH Reading",      "D-metrics and V-metrics; DVH screens, isodose slices diagnose", C_WHITE),
        ("D95 / D2",         "D95 ≥ 95%  ·  D2% < 107% (ICRU 83 near-maximum)", C_WHITE),
        ("HI",               "Homogeneity Index = (D2−D98)/D50  ·  <0.1 good  <0.15 acceptable", C_WHITE),
        ("CI",               "Conformity Index = PIV/TV  ·  Ideal 1.0–1.2  ·  Higher expected in 3DCRT", C_WHITE),
        ("Cold Spots",       "Under-dosed CTV region — most important failure to detect", C_RED),
    ]

    y = Cm(2.3)
    for title, detail, col in items:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.2), fill_color=C_DARK_NAVY)
        add_textbox(slide, Cm(1.7), y + Cm(0.1), Cm(8), Cm(1.0),
            title, font_size=16, bold=True, color=col)
        add_textbox(slide, Cm(10.5), y + Cm(0.25), Cm(20.5), Cm(1.6),
            detail, font_size=13, color=C_DIM if col == C_WHITE else C_WHITE)
        y += Cm(2.35)

    add_slide_number(slide, slide_num)
    return slide

def make_dvh_volume_hierarchy(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "DVH + Volume Hierarchy", font_size=32, bold=True, color=C_CYAN)

    volumes = [
        ("GTV", "Gross Tumour Volume",    "Visible/palpable tumour on imaging",       C_RED),
        ("CTV", "Clinical Target Volume", "GTV + microscopic spread (clinical target)", C_ORANGE),
        ("ITV", "Internal Target Volume", "CTV + internal motion margin",              C_YELLOW),
        ("PTV", "Planning Target Volume", "ITV/CTV + set-up uncertainty",              C_CYAN),
    ]

    y = Cm(2.3)
    for abbr, name, desc, col in volumes:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.2), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.5), Cm(2.2), fill_color=col)
        add_textbox(slide, Cm(2.2), y + Cm(0.1), Cm(3), Cm(1.0),
            abbr, font_size=20, bold=True, color=col)
        add_textbox(slide, Cm(5.5), y + Cm(0.1), Cm(9), Cm(1.0),
            name, font_size=15, bold=True, color=C_WHITE)
        add_textbox(slide, Cm(15), y + Cm(0.25), Cm(16), Cm(1.6),
            desc, font_size=13, color=C_DIM)
        y += Cm(2.4)

    add_textbox(slide, Cm(1.5), Cm(12.5), Cm(30), Cm(1.0),
        '"We prescribe to PTV but the clinical aim is to cover the CTV"',
        font_size=15, bold=True, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_rect(slide, Cm(1.5), Cm(14.0), Cm(30), Cm(2.5), fill_color=C_DARK_NAVY)
    add_textbox(slide, Cm(1.7), Cm(14.1), Cm(29.6), Cm(2.3),
        "DVH: screens for problems at a glance  ·  Isodose slices: diagnose exactly where",
        font_size=15, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_coverage_metrics(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "Coverage Metrics", font_size=32, bold=True, color=C_CYAN)

    metrics = [
        ("D95%",
         "Standard aim: D95 ≥ 95%",
         "Dose received by 95% of the PTV. Key coverage benchmark."),
        ("D2%",
         "Near-maximum · ICRU 83: D2% < 107%",
         "Dose to the hottest 2% of the volume. Replaces Dmax in modern reporting."),
        ("HI",
         "Homogeneity Index = (D2 − D98) / D50",
         "<0.1 good   <0.15 acceptable   >0.2 poor"),
        ("CI",
         "Conformity Index = PIV / TV",
         "Ideal 1.0–1.2   Higher values expected in 3DCRT (no IMRT shaping)"),
    ]

    y = Cm(2.5)
    for abbr, headline, detail in metrics:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(3.0), fill_color=C_DARK_NAVY)
        add_textbox(slide, Cm(1.7), y + Cm(0.15), Cm(3.5), Cm(1.2),
            abbr, font_size=24, bold=True, color=C_CYAN, font_name="Courier New")
        add_textbox(slide, Cm(5.5), y + Cm(0.15), Cm(24.5), Cm(1.2),
            headline, font_size=17, bold=True, color=C_WHITE)
        add_textbox(slide, Cm(5.5), y + Cm(1.4), Cm(24.5), Cm(1.4),
            detail, font_size=14, color=C_DIM)
        y += Cm(3.2)

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"Coverage is negotiable where PTV abuts serial OAR"',
        font_size=14, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_block_o_oars(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_rect(slide, Cm(1.5), Cm(0.5), Cm(6), Cm(1.5), fill_color=C_GREEN)
    add_textbox(slide, Cm(1.7), Cm(0.55), Cm(5.6), Cm(1.4),
        "Block O: Organs at Risk", font_size=18, bold=True, color=C_DARK_NAVY)
    add_textbox(slide, Cm(8.5), Cm(0.8), Cm(10), Cm(1.0),
        "⏱ 10 minutes", font_size=16, color=C_GREEN)

    topics = [
        ("Serial vs Parallel",   "Organ architecture determines the relevant dose metric"),
        ("Dose Metrics",         "Serial → Dmax or D1cc  ·  Parallel → Dmean or Vx"),
        ("QUANTEC Constraints",  "Evidence-based dose–volume limits for normal tissue complication"),
        ("PRV",
         "Planning Risk Volume — structural margin around serial OAR (analogous to PTV for targets)"),
    ]

    y = Cm(2.5)
    for title, detail in topics:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.8), fill_color=C_DARK_NAVY)
        add_textbox(slide, Cm(1.7), y + Cm(0.1), Cm(10), Cm(1.2),
            title, font_size=17, bold=True, color=C_GREEN)
        add_textbox(slide, Cm(1.7), y + Cm(1.3), Cm(28.5), Cm(1.3),
            detail, font_size=14, color=C_WHITE, word_wrap=True)
        y += Cm(3.0)

    add_slide_number(slide, slide_num)
    return slide

def make_serial_parallel(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.6), Cm(30), Cm(1.4),
        "Serial vs Parallel OARs", font_size=32, bold=True, color=C_CYAN)

    # Serial
    add_rect(slide, Cm(1.5), Cm(2.4), Cm(14.5), Cm(11.0), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(1.5), Cm(2.4), Cm(14.5), Cm(0.5), fill_color=C_RED)
    add_textbox(slide, Cm(1.7), Cm(3.0), Cm(14), Cm(1.2),
        "SERIAL OARs", font_size=20, bold=True, color=C_RED)
    add_textbox(slide, Cm(1.7), Cm(4.3), Cm(14), Cm(1.0),
        "Metric: Dmax / D1cc", font_size=15, bold=True, color=C_WHITE)
    add_textbox(slide, Cm(1.7), Cm(5.3), Cm(14), Cm(1.5),
        "One failed sub-unit = organ failure", font_size=13, color=C_DIM)
    add_textbox(slide, Cm(1.7), Cm(6.8), Cm(14), Cm(4.5),
        "Cord:       ≤ 45 Gy\nBrainstem:  ≤ 54 Gy\nChiasm:     ≤ 54 Gy\n\nApply PRV margin (structural buffer)",
        font_size=14, color=C_WHITE, font_name="Courier New")

    # Parallel
    add_rect(slide, Cm(17), Cm(2.4), Cm(14.5), Cm(11.0), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(17), Cm(2.4), Cm(14.5), Cm(0.5), fill_color=C_GREEN)
    add_textbox(slide, Cm(17.2), Cm(3.0), Cm(14), Cm(1.2),
        "PARALLEL OARs", font_size=20, bold=True, color=C_GREEN)
    add_textbox(slide, Cm(17.2), Cm(4.3), Cm(14), Cm(1.0),
        "Metric: Mean dose / Vx", font_size=15, bold=True, color=C_WHITE)
    add_textbox(slide, Cm(17.2), Cm(5.3), Cm(14), Cm(1.5),
        "Redundant sub-units — partial loss tolerated", font_size=13, color=C_DIM)
    add_textbox(slide, Cm(17.2), Cm(6.8), Cm(14), Cm(4.5),
        "Lung:    V20 < 30%\nParotid: mean < 26 Gy\nRectum:  V70 < 20%",
        font_size=14, color=C_WHITE, font_name="Courier New")

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"Constraints are not absolutes — re-irradiation, chemo, comorbidities all move the line"',
        font_size=13, italic=True, color=C_DIM, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_block_ps(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_rect(slide, Cm(1.5), Cm(0.5), Cm(8), Cm(1.5), fill_color=C_ORANGE)
    add_textbox(slide, Cm(1.7), Cm(0.55), Cm(7.8), Cm(1.4),
        "Block P·S: Prescription & Summation", font_size=18, bold=True, color=C_WHITE)
    add_textbox(slide, Cm(10.5), Cm(0.8), Cm(10), Cm(1.0),
        "⏱ 8 minutes", font_size=16, color=C_ORANGE)

    sections = [
        ("Prescription Verification",
         ["Dose per fraction correct?",
          "Total dose = dose/fraction × number of fractions",
          "Normalisation point in stable representative tissue",
          "Laterality confirmed (left vs right side)"]),
        ("Summation Checklist",
         ["Prior RT to same region — cumulative dose calculated?",
          "Brachytherapy boost phase — EBRT is Phase 1 only",
          "EQD2 summation for re-irradiation context",
          "All phases documented in plan directive"]),
        ("Triage Tiers",
         ["CRITICAL → Re-plan immediately",
          "MAJOR → Correct or justify with documentation",
          "MINOR → Comment; proceed with review"]),
    ]

    y = Cm(2.5)
    for heading, bullets in sections:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(0.9), fill_color=C_ORANGE)
        add_textbox(slide, Cm(1.7), y + Cm(0.05), Cm(29), Cm(0.8),
            heading, font_size=16, bold=True, color=C_WHITE)
        by = y + Cm(1.0)
        for b in bullets:
            add_textbox(slide, Cm(2.5), by, Cm(29), Cm(0.75),
                "· " + b, font_size=13, color=C_DIM)
            by += Cm(0.75)
        y = by + Cm(0.3)

    add_slide_number(slide, slide_num)
    return slide

def make_tiers_scorecard(prs, slide_num):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Triage Tiers + Evaluation Scorecard", font_size=30, bold=True, color=C_CYAN)

    tiers = [
        ("CRITICAL — Re-plan",
         "Geographic miss  ·  Serial OAR above hard limit  ·  Hotspot in critical organ",
         C_RED),
        ("MAJOR — Correct or justify",
         "Parallel OAR over constraint  ·  V95 borderline  ·  Avoidable beam through OAR",
         C_ORANGE),
        ("MINOR — Comment only",
         "Minor HI elevation  ·  Documentation gap  ·  Suboptimal but acceptable",
         C_YELLOW),
    ]

    y = Cm(2.2)
    for label, detail, col in tiers:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.2), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.5), Cm(2.2), fill_color=col)
        add_textbox(slide, Cm(2.2), y + Cm(0.1), Cm(12), Cm(1.0),
            label, font_size=16, bold=True, color=col)
        add_textbox(slide, Cm(2.2), y + Cm(1.1), Cm(28), Cm(0.9),
            detail, font_size=13, color=C_DIM)
        y += Cm(2.4)

    # Reject on sight
    add_rect(slide, Cm(1.5), y + Cm(0.2), Cm(13.5), Cm(4.5), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(1.5), y + Cm(0.2), Cm(13.5), Cm(0.5), fill_color=C_RED)
    add_textbox(slide, Cm(1.7), y + Cm(0.8), Cm(13), Cm(1.0),
        "Reject on Sight:", font_size=14, bold=True, color=C_RED)
    add_textbox(slide, Cm(1.7), y + Cm(1.8), Cm(13), Cm(2.5),
        "1) Cold CTV\n2) Serial OAR hard limit breach\n3) Unsafe / undeliverable",
        font_size=13, color=C_WHITE)

    # Evaluation routine
    add_rect(slide, Cm(16.5), y + Cm(0.2), Cm(15.5), Cm(4.5), fill_color=C_DARK_NAVY)
    add_rect(slide, Cm(16.5), y + Cm(0.2), Cm(15.5), Cm(0.5), fill_color=C_CYAN)
    add_textbox(slide, Cm(16.7), y + Cm(0.8), Cm(15), Cm(1.0),
        "Evaluation Routine:", font_size=14, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(16.7), y + Cm(1.8), Cm(15), Cm(2.5),
        "Prescription → Contours → Beams → Isodose\n→ Coverage → Hotspots → DVH → OARs → Decision",
        font_size=12, color=C_WHITE)

    add_slide_number(slide, slide_num)
    return slide

# ── Case opener helper ──────────────────────────────────────────
def make_case_opener(prs, slide_num, case_num, case_title, rx, technique, oars, quote=None):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_footer_bar(slide)

    # Case number pill
    add_rect(slide, Cm(1.5), Cm(0.6), Cm(4.5), Cm(1.5), fill_color=C_CYAN)
    add_textbox(slide, Cm(1.7), Cm(0.65), Cm(4.2), Cm(1.4),
        f"Case {case_num}", font_size=18, bold=True, color=C_DARK_NAVY, align=PP_ALIGN.CENTER)

    add_textbox(slide, Cm(7), Cm(0.8), Cm(24), Cm(1.2),
        case_title, font_size=28, bold=True, color=C_WHITE)

    add_rect(slide, Cm(1.5), Cm(2.8), Cm(30), Cm(0.1), fill_color=C_CYAN)

    add_textbox(slide, Cm(1.5), Cm(3.2), Cm(5), Cm(0.8),
        "Prescription", font_size=12, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(1.5), Cm(4.0), Cm(30), Cm(0.9),
        rx, font_size=16, color=C_WHITE)

    add_textbox(slide, Cm(1.5), Cm(5.2), Cm(5), Cm(0.8),
        "Technique", font_size=12, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(1.5), Cm(6.0), Cm(30), Cm(0.9),
        technique, font_size=15, color=C_DIM)

    add_textbox(slide, Cm(1.5), Cm(7.2), Cm(5), Cm(0.8),
        "OARs", font_size=12, bold=True, color=C_CYAN)
    add_textbox(slide, Cm(1.5), Cm(8.0), Cm(30), Cm(0.9),
        "  ·  ".join(oars), font_size=15, color=C_DIM)

    if quote:
        add_rect(slide, Cm(1.5), Cm(10.5), Cm(30), Cm(2.5), fill_color=C_NAVY)
        add_textbox(slide, Cm(1.7), Cm(10.7), Cm(29.6), Cm(2.1),
            f'"{quote}"', font_size=16, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

def make_case_review(prs, slide_num, case_name, guide_items):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_rect(slide, Cm(1.5), Cm(0.5), Cm(8), Cm(1.5), fill_color=C_DARK_NAVY)
    add_textbox(slide, Cm(1.7), Cm(0.6), Cm(7.6), Cm(1.3),
        f"{case_name} — Guide Your Eye", font_size=20, bold=True, color=C_CYAN)

    y = Cm(2.5)
    for i, item in enumerate(guide_items):
        bg = C_DARK_NAVY if i % 2 == 0 else RGBColor(0x0c, 0x25, 0x3d)
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(1.8), fill_color=bg)
        add_textbox(slide, Cm(2.0), y + Cm(0.2), Cm(1.0), Cm(1.4),
            f"{i+1}.", font_size=15, bold=True, color=C_CYAN)
        add_textbox(slide, Cm(3.2), y + Cm(0.2), Cm(28), Cm(1.4),
            item, font_size=14, color=C_WHITE)
        y += Cm(1.9)

    add_slide_number(slide, slide_num)
    return slide

def make_case_findings(prs, slide_num, case_name, findings):
    """findings = list of (tag, text, severity_color)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        f"{case_name} — Findings", font_size=28, bold=True, color=C_CYAN)

    y = Cm(2.2)
    for tag, text, col in findings:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(1.7), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.4), Cm(1.7), fill_color=col)
        add_textbox(slide, Cm(2.1), y + Cm(0.1), Cm(4.5), Cm(0.8),
            tag, font_size=11, bold=True, color=col)
        add_textbox(slide, Cm(7.0), y + Cm(0.2), Cm(23), Cm(1.3),
            text, font_size=13, color=C_WHITE)
        y += Cm(1.85)

    add_slide_number(slide, slide_num)
    return slide

def make_case_solutions(prs, slide_num, case_name, solutions, footer=None):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        f"{case_name} — Solutions", font_size=28, bold=True, color=C_CYAN)

    y = Cm(2.2)
    for i, sol in enumerate(solutions):
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(1.7), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.5), Cm(1.7), fill_color=C_GREEN)
        add_textbox(slide, Cm(2.2), y + Cm(0.1), Cm(1.5), Cm(1.5),
            str(i+1), font_size=18, bold=True, color=C_GREEN, align=PP_ALIGN.CENTER)
        add_textbox(slide, Cm(4.0), y + Cm(0.2), Cm(26.5), Cm(1.3),
            sol, font_size=13, color=C_WHITE)
        y += Cm(1.85)

    if footer:
        add_rect(slide, Cm(1.5), y + Cm(0.3), Cm(30), Cm(2.0), fill_color=RGBColor(0x12, 0x32, 0x52))
        add_textbox(slide, Cm(1.7), y + Cm(0.4), Cm(29.6), Cm(1.8),
            f'"{footer}"', font_size=15, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, slide_num)
    return slide

# ── Build all slides ─────────────────────────────────────────────
def build_deck():
    prs = new_prs()

    # 1. Title
    make_title_slide(prs, 1)

    # 2. Session Map
    make_session_map(prs, 2)

    # 3. Four Central Questions
    make_four_questions(prs, 3)

    # 4. Plan Quality Spectrum
    make_plan_quality_spectrum(prs, 4)

    # 5. FCB-CHOPS Evolution
    make_fcbchops_evolution(prs, 5)

    # 6. FCB-CHOPS Framework
    make_fcbchops_framework(prs, 6)

    # 7. Block B: Beams
    make_block_b_beams(prs, 7)

    # 8. Field Arrangement Diagram
    make_field_arrangement(prs, 8)

    # 9. Technical Quality
    make_technical_quality(prs, 9)

    # 10. Hot Spots Diagram
    make_hotspots_diagram(prs, 10)

    # 11. Block C: Coverage
    make_block_c_coverage(prs, 11)

    # 12. DVH + Volume Hierarchy
    make_dvh_volume_hierarchy(prs, 12)

    # 13. Coverage Metrics
    make_coverage_metrics(prs, 13)

    # 14. Block O: OARs
    make_block_o_oars(prs, 14)

    # 15. Serial vs Parallel OARs
    make_serial_parallel(prs, 15)

    # 16. Block P·S
    make_block_ps(prs, 16)

    # 17. Tiers + Scorecard
    make_tiers_scorecard(prs, 17)

    # 18. Case 1 - Breast Opener
    make_case_opener(prs, 18,
        case_num=1,
        case_title="Breast — Whole-Breast Tangents",
        rx="2.67 Gy × 15 = 40.05 Gy  ·  Whole breast, right side",
        technique="Tangential photon fields (medial + lateral), FiF",
        oars=["Heart", "Ipsilateral lung", "Contralateral breast"],
        quote="5 minutes — score Plan A")

    # 19. Breast Plan A Review
    make_case_review(prs, 19, "Breast Plan A", [
        "Isocentre position — is it central in the breast tissue?",
        "Tangent borders — are they crossing the midline?",
        "Flash — at least 2 cm beyond skin?",
        "Normalisation point — in breast tissue (not lung)?",
        "Hotspot < 115%?",
        "Heart + lung DVH — within tolerance?",
    ])

    # 20. Breast - Bad Plan Findings
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)
    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Breast — Plan A is the BAD Plan", font_size=28, bold=True, color=C_RED)

    col1_items = [
        ("CRITICAL", "Norm point in lung",                  C_RED),
        ("CRITICAL", "Hotspot > 115%",                      C_RED),
        ("CRITICAL", "Heart over tolerance",                 C_RED),
        ("MAJOR",    "Isocentre too deep (lung/heart)",     C_ORANGE),
        ("MAJOR",    "Fields crossing midline",              C_ORANGE),
        ("MAJOR",    "No flash",                             C_ORANGE),
        ("MINOR",    "No wedges / shaping",                  C_YELLOW),
    ]
    y = Cm(2.2)
    for tag, text, col in col1_items:
        add_rect(slide, Cm(1.5), y, Cm(18), Cm(1.5), fill_color=C_DARK_NAVY)
        add_rect(slide, Cm(1.5), y, Cm(0.4), Cm(1.5), fill_color=col)
        add_textbox(slide, Cm(2.1), y + Cm(0.1), Cm(4.5), Cm(0.8),
            tag, font_size=11, bold=True, color=col)
        add_textbox(slide, Cm(6.5), y + Cm(0.15), Cm(12.5), Cm(1.2),
            text, font_size=13, color=C_WHITE)
        y += Cm(1.65)

    add_textbox(slide, Cm(21), Cm(2.2), Cm(10), Cm(1.0),
        "What must change:", font_size=14, bold=True, color=C_CYAN)
    changes = [
        "Move iso to breast tissue",
        "Norm point in breast (not lung)",
        "No midline crossing",
        "≥ 2 cm flash beyond skin",
        "Add wedges / FiF for shaping",
    ]
    y2 = Cm(3.4)
    for c in changes:
        add_textbox(slide, Cm(21), y2, Cm(11), Cm(0.85),
            "→  " + c, font_size=13, color=C_WHITE)
        y2 += Cm(0.95)

    add_slide_number(slide, 20)

    # 21. Breast - Acceptable vs Good
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)
    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Breast — Acceptable vs Good", font_size=28, bold=True, color=C_CYAN)

    plans = [
        ("Plan B", "Acceptable — Wedges",
         ["Good isocentre placement",
          "Norm point in breast tissue",
          "Flash ≥ 2 cm",
          "Stays midline",
          "Hotspot < 115%",
          "HI < 0.15"],
         C_YELLOW),
        ("Plan C", "Good — Field-in-Field",
         ["All Plan B features PLUS:",
          "Hotspot < 110%",
          "HI < 0.10",
          "FiF sub-segments block peak hotspots"],
         C_GREEN),
    ]
    x = Cm(1.5)
    for plan_id, plan_label, bullets, col in plans:
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(12.5), fill_color=C_DARK_NAVY)
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(0.5), fill_color=col)
        add_textbox(slide, x + Cm(0.2), Cm(2.8), Cm(14), Cm(0.9),
            plan_id, font_size=20, bold=True, color=col)
        add_textbox(slide, x + Cm(0.2), Cm(3.8), Cm(14), Cm(0.9),
            plan_label, font_size=15, bold=False, color=C_WHITE)
        by = Cm(5.0)
        for b in bullets:
            add_textbox(slide, x + Cm(0.4), by, Cm(14), Cm(0.85),
                "✓  " + b, font_size=13, color=C_DIM)
            by += Cm(0.95)
        x += Cm(16)

    add_slide_number(slide, 21)

    # 22. Breast - Teaching Moment
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)
    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Breast — Teaching Moment: Homogeneity", font_size=28, bold=True, color=C_CYAN)

    cols3 = [
        ("BAD\n(No shaping)", "HI > 0.2\nDose piles at beam entry\nHotspot outside PTV possible", C_RED),
        ("ACCEPTABLE\n(Wedges)", "HI < 0.15\nCompensates sloping surface\nImproved homogeneity", C_YELLOW),
        ("GOOD\n(Field-in-Field)", "HI < 0.10\nSub-segments block peak hotspots\nOptimal homogeneity", C_GREEN),
    ]
    x = Cm(1.5)
    for label, detail, col in cols3:
        add_rect(slide, x, Cm(2.2), Cm(9.5), Cm(10.0), fill_color=C_DARK_NAVY)
        add_rect(slide, x, Cm(2.2), Cm(9.5), Cm(0.5), fill_color=col)
        add_textbox(slide, x + Cm(0.2), Cm(2.9), Cm(9.1), Cm(2.5),
            label, font_size=18, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + Cm(0.2), Cm(5.5), Cm(9.1), Cm(6.5),
            detail, font_size=14, color=C_WHITE, align=PP_ALIGN.CENTER, word_wrap=True)
        x += Cm(10.5)

    add_textbox(slide, Cm(1.5), H - Cm(2.0), Cm(30), Cm(1.2),
        '"The three plans differ mainly in dose shaping — same geometry, different homogeneity tools"',
        font_size=14, italic=True, color=C_DIM, align=PP_ALIGN.CENTER)
    add_slide_number(slide, 22)

    # 23. Case 2 - Gynae Opener
    make_case_opener(prs, 23,
        case_num=2,
        case_title="Gynaecology — Whole-Pelvis Irradiation",
        rx="1.8 Gy × 25 = 45 Gy  ·  Whole pelvis, cervix + parametria + regional nodes",
        technique="4-field box AP/PA + 2 laterals, FiF",
        oars=["Rectum", "Small bowel", "Bladder", "Femoral heads"])

    # 24. Gynae Review
    make_case_review(prs, 24, "Gynae", [
        "PTV coverage — V95 and D95 on DVH",
        "OARs on DVH — rectum, small bowel, bladder, femoral heads",
        "Rectum — high-dose isodose line on axial/sagittal",
        "Small bowel loops — anterior beam enters through small bowel?",
        "Bladder filling protocol documented?",
        "Femoral heads — Dmax within QUANTEC limit?",
    ])

    # 25. Gynae Findings
    make_case_findings(prs, 25, "Gynae", [
        ("✓ Coverage",   "PTV well covered  ·  V95 and D95 within targets",          C_GREEN),
        ("↑ Rectum",     "V45/V30 elevated — above QUANTEC for this dose level",       C_ORANGE),
        ("~ Small bowel","V45 elevated — loops in beam path",                          C_YELLOW),
        ("✓ Bladder",    "Within QUANTEC limits",                                      C_GREEN),
        ("✓ Fem heads",  "Within QUANTEC tolerance",                                   C_GREEN),
        ("Summation",    "Brachy boost to follow — EBRT is Phase 1 only",              C_CYAN),
        ("QUANTEC ref",  "Rectum: V70 < 20%  ·  V50 < 50%",                           C_DIM),
    ])

    # 26. Gynae Solutions
    make_case_solutions(prs, 26, "Gynae", [
        "Adjust AP/PA vs lateral field weighting to reduce posterior dose",
        "Patient positioning & preparation — full bladder, belly board, small bowel displacement",
        "Reduce posterior PTV margin at rectum-prostate interface",
        "Adjust prescription if clinically appropriate (discuss with attending)",
        "Accept with documentation — geometric constraint, not a planning error",
    ], footer="EBRT is phase 1 — brachytherapy adds rectal and bladder dose. Track cumulative from the start.")

    # 27. Case 3 - Prostate Opener
    make_case_opener(prs, 27,
        case_num=3,
        case_title="Prostate — 3-Field Plan",
        rx="2.5 Gy × 27 = 67.5 Gy  ·  Prostate + proximal SVs",
        technique="AP + 2 posterior obliques  ·  Full bladder  ·  Empty rectum",
        oars=["Rectum", "Penile bulb", "Bladder", "Femoral heads", "Small bowel"])

    # 28. Prostate Review
    make_case_review(prs, 28, "Prostate", [
        "D95 / D98 — is the CTV fully covered?",
        "Rectum — V70, V50 vs QUANTEC tolerance",
        "Penile bulb — contoured? Mean dose? QUANTEC 52 Gy",
        "Bladder and femoral heads — within tolerance?",
        "Posterior PTV margin vs rectum — where is the interface?",
        "Isocentre — in prostate tissue (not rectum / air)?",
    ])

    # 29. Prostate Findings
    make_case_findings(prs, 29, "Prostate", [
        ("✓ Coverage",   "CTV D98 ≈ 99.2%  ·  PTV D98 ≈ 96.6%  ·  Coverage good",   C_GREEN),
        ("↑ Rectum",     "V55 ≈ 37.9%  ·  V59 ≈ 32.9%  — over QUANTEC tolerance",    C_RED),
        ("↑ Penile bulb","Mean ≈ 62.8 Gy  — above QUANTEC limit of 52 Gy",            C_RED),
        ("~ Small bowel","Minor focal involvement  ·  Within tolerance",               C_YELLOW),
        ("✓ Bladder",    "Within tolerance",                                            C_GREEN),
        ("✓ Fem heads",  "Within tolerance",                                            C_GREEN),
        ("QUANTEC ref",  "Rectum: V70 <20%  V50 <50%  ·  Penile bulb: Dmean <52 Gy  D90 <50 Gy", C_DIM),
    ])

    # 30. Prostate Solutions
    make_case_solutions(prs, 30, "Prostate", [
        "Reduce inferior margin — decrease penile bulb dose",
        "Reduce posterior PTV margin at prostate–rectum interface",
        "Adjust posterior oblique field weighting",
        "Adjust prescription if clinically appropriate",
        "Accept with documentation — geometric proximity, not a planning error",
    ], footer="Coverage was good — the problem was the OARs. Name it, don't ignore it.")

    # 31. Case 4 - Pituitary Opener
    make_case_opener(prs, 31,
        case_num=4,
        case_title="Pituitary Adenoma — Coplanar 3-Field Plan",
        rx="1.8 Gy × 30 = 54 Gy  ·  Pituitary fossa + margin",
        technique="3-field coplanar MLC shaped",
        oars=["Optic chiasm", "Optic nerves", "Brainstem", "Brain"],
        quote="Pay very close attention to the optic chiasm dose")

    # 32. Pituitary Review
    make_case_review(prs, 32, "Pituitary", [
        "D95 / D98 / Dmin — is the PTV fully covered?",
        "Optic chiasm Dmax — QUANTEC tolerance is 54 Gy (hard limit)",
        "Optic nerves — bilateral Dmax",
        "Brainstem — Dmax and mean dose",
        "Inferior undercoverage — sphenoid sinus physics effect?",
        "Non-coplanar option — does it improve coverage or OAR sparing?",
    ])

    # 33. Pituitary Findings
    make_case_findings(prs, 33, "Pituitary", [
        ("✓ Coverage",    "PTV D95 adequate",                                          C_GREEN),
        ("↑ Chiasm",      "Dmax 54.9 Gy — ABOVE 54 Gy QUANTEC hard limit (mandatory sign-off)", C_RED),
        ("~ Inferior PTV","PTV min 41.75 Gy — reduced inferior coverage",              C_YELLOW),
        ("✓ Brainstem",   "Dmax 52.5 Gy  ·  Mean ~11 Gy  ·  Within tolerance",        C_GREEN),
        ("Note",          "Sphenoid air sinus undercoverage 88.7% — physics-driven, expected",   C_DIM),
        ("Architecture",  "Chiasm is serial: Dmax governs — even small breach = tolerance violation", C_ORANGE),
    ])

    # 34. Pituitary Solutions
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)
    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Pituitary — Solutions: Coplanar vs Non-Coplanar", font_size=26, bold=True, color=C_CYAN)

    plan_data = [
        ("Plan A — Coplanar",
         ["Chiasm Dmax:  54.9 Gy  ↑  (above tolerance)",
          "PTV min:       41.75 Gy",
          "Brainstem mean: ~11 Gy"],
         C_ORANGE),
        ("Plan B — Non-Coplanar",
         ["Chiasm Dmax:  ~54.9 Gy  (SAME — cannot move chiasm)",
          "PTV min:       43.9 Gy  (+2.1 Gy improved)  ↑",
          "Brainstem mean: 9.7 Gy  (↓ 1.6 Gy improved)"],
         C_GREEN),
    ]
    x = Cm(1.5)
    for plan_id, bullets, col in plan_data:
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(11.0), fill_color=C_DARK_NAVY)
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(0.5), fill_color=col)
        add_textbox(slide, x + Cm(0.2), Cm(2.8), Cm(14.1), Cm(1.0),
            plan_id, font_size=18, bold=True, color=col)
        by = Cm(4.0)
        for b in bullets:
            add_textbox(slide, x + Cm(0.2), by, Cm(14.1), Cm(0.9),
                b, font_size=13, color=C_WHITE, font_name="Courier New")
            by += Cm(1.05)
        x += Cm(16)

    add_textbox(slide, Cm(1.5), H - Cm(2.2), Cm(30), Cm(1.8),
        '"Non-coplanar fields spread entrance angles — but cannot move the chiasm. Documentation required in both."',
        font_size=14, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    add_slide_number(slide, 34)

    # 35. Case 5 - H&N Opener
    make_case_opener(prs, 35,
        case_num=5,
        case_title="Head & Neck — Parotid Tumour",
        rx="2 Gy × 33 = 66 Gy  ·  Parotid tumour + regional nodes",
        technique="4-field 3DCRT  ·  Photon + possible electron match",
        oars=["Contralateral parotid", "Spinal cord", "Brainstem", "Oral cavity"],
        quote="Focus especially on the contralateral parotid dose")

    # 36. H&N Review
    make_case_review(prs, 36, "H&N", [
        "PTV coverage — D95, D98",
        "Cord Dmax — hard limit 45 Gy (PRV) / 50 Gy (cord)",
        "Brainstem Dmax",
        "Contralateral parotid mean — QUANTEC < 26 Gy",
        "Ipsilateral parotid — likely sacrifice; document",
        "Beam entry / exit path — does right lateral exit through contralateral parotid?",
    ])

    # 37. H&N Findings
    make_case_findings(prs, 37, "H&N", [
        ("✓ Coverage",    "PTV mean ≈ 66.0 Gy",                                        C_GREEN),
        ("↑ Contra parotid","Mean ≈ 3.30 Gy — higher than necessary (Plan A, 4-field)", C_ORANGE),
        ("✓ Cord",        "Dmax 19.9 Gy  ·  Within tolerance",                         C_GREEN),
        ("✓ Brainstem",   "Dmax 18.2 Gy  ·  Within tolerance",                         C_GREEN),
        ("~ Submandibular","R mean ≈ 40.4 Gy — elevated",                               C_YELLOW),
        ("Summation",     "Electron/photon match junction — dose at junction must be summed", C_CYAN),
        ("Cause",         "Right lateral field exits through contralateral parotid — geometry problem", C_DIM),
        ("Late effect",   "Xerostomia: most common serious late effect of H&N RT",       C_DIM),
    ])

    # 38. H&N Solutions
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer_bar(slide)
    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "H&N — Solutions: 4-Field vs 3-Field", font_size=26, bold=True, color=C_CYAN)

    hn_plans = [
        ("Plan A — 4-Field",
         ["PTV mean:          ≈ 66.0 Gy",
          "Contra parotid:    3.30 Gy  ↑  (above optimal)",
          "Cord Dmax:         19.9 Gy",
          "Brainstem Dmax:    18.2 Gy"],
         C_ORANGE),
        ("Plan B — 3-Field\n(right lateral removed)",
         ["PTV mean:          ≈ 66.0 Gy  (maintained)",
          "Contra parotid:    1.14 Gy  ↓  (↓ 2.2 Gy improvement)",
          "Cord Dmax:         17.8 Gy  (↓ 2.1 Gy)",
          "Brainstem Dmax:    16.1 Gy  (↓ 2.1 Gy)"],
         C_GREEN),
    ]
    x = Cm(1.5)
    for plan_id, bullets, col in hn_plans:
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(11.0), fill_color=C_DARK_NAVY)
        add_rect(slide, x, Cm(2.2), Cm(14.5), Cm(0.5), fill_color=col)
        add_textbox(slide, x + Cm(0.2), Cm(2.8), Cm(14.1), Cm(1.0),
            plan_id, font_size=17, bold=True, color=col)
        by = Cm(4.0)
        for b in bullets:
            add_textbox(slide, x + Cm(0.2), by, Cm(14.1), Cm(0.9),
                b, font_size=13, color=C_WHITE, font_name="Courier New")
            by += Cm(1.05)
        x += Cm(16)

    add_textbox(slide, Cm(1.5), H - Cm(2.2), Cm(30), Cm(1.8),
        '"Fewer fields ≠ worse plan. Remove the beam that exits through the contralateral parotid."',
        font_size=15, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    add_slide_number(slide, 38)

    # 39. Wrap-up
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_rect(slide, Cm(0), Cm(0), Cm(0.6), H, fill_color=C_CYAN)
    add_footer_bar(slide)

    add_textbox(slide, Cm(1.5), Cm(0.5), Cm(30), Cm(1.3),
        "Wrap-up — 5 Key Takeaways", font_size=30, bold=True, color=C_CYAN)

    takeaways = [
        ("1", "FCB-CHOPS every time",                              "Use the framework as a safety net — don't skip letters"),
        ("2", "DVH screens; slices diagnose",                      "Use both — DVH flags the problem, isodose slices locate it"),
        ("3", "Match metric to organ architecture",                 "Serial → Dmax / D1cc   ·   Parallel → mean / Vx"),
        ("4", "Normalisation point matters",                       "In lung or air = unreliable; must be in stable representative tissue"),
        ("5", "Document every trade-off",                          "If you accept a compromise, write it down — name it, don't ignore it"),
    ]
    y = Cm(2.2)
    for num, headline, detail in takeaways:
        add_rect(slide, Cm(1.5), y, Cm(30), Cm(2.2), fill_color=C_NAVY)
        add_textbox(slide, Cm(1.7), y + Cm(0.1), Cm(1.2), Cm(2.0),
            num, font_size=24, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
        add_textbox(slide, Cm(3.3), y + Cm(0.1), Cm(14), Cm(1.0),
            headline, font_size=16, bold=True, color=C_WHITE)
        add_textbox(slide, Cm(3.3), y + Cm(1.1), Cm(26), Cm(0.9),
            detail, font_size=13, color=C_DIM)
        y += Cm(2.4)

    # References
    add_textbox(slide, Cm(1.5), H - Cm(2.5), Cm(30), Cm(0.7),
        "References: FCB-CHOPS Weisman 2024 (PMID 40017913)  ·  ICRU 50/62/83  ·  QUANTEC 2010  ·  Emami 1991  ·  Khan textbook",
        font_size=11, color=C_DIM)
    add_textbox(slide, Cm(1.5), H - Cm(1.8), Cm(30), Cm(1.0),
        '"A good plan is a documented compromise"',
        font_size=16, italic=True, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_number(slide, 39)

    return prs

# ── Save ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    out = "/home/user/3DCRT-Evaluation-Workshop/3DCRT_Workshop_Deck_v2.pptx"
    prs = build_deck()
    prs.save(out)
    print(f"Saved → {out}")
    print(f"Slides: {len(prs.slides)}")
