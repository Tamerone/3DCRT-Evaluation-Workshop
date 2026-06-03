"""
3DCRT Workshop Deck v2 — PowerPoint generator (visual redesign)
Run: python3 make_pptx.py
Output: 3DCRT_Workshop_Deck_v2.pptx
"""

from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pptx

# ── Colour palette ───────────────────────────────────────────────
C_NAVY      = RGBColor(0x0E, 0x28, 0x41)
C_NAVY2     = RGBColor(0x16, 0x35, 0x54)
C_DARK_NAVY = RGBColor(0x0a, 0x1f, 0x33)
C_CYAN      = RGBColor(0x0F, 0x9E, 0xD5)
C_ORANGE    = RGBColor(0xE9, 0x71, 0x32)
C_GREEN     = RGBColor(0x29, 0xA3, 0x5B)
C_RED       = RGBColor(0xD4, 0x40, 0x40)
C_WHITE     = RGBColor(0xE8, 0xED, 0xF2)
C_DIM       = RGBColor(0x9A, 0xAC, 0xBE)
C_YELLOW    = RGBColor(0xF5, 0xC5, 0x18)
C_BLOCK_B   = RGBColor(0x99, 0x33, 0xBB)
C_BLOCK_C   = RGBColor(0xCC, 0x33, 0x99)
C_BLOCK_O   = RGBColor(0xFF, 0x88, 0x22)
C_BLOCK_PS  = RGBColor(0x77, 0xCC, 0x33)

# Ghost colors for block transition slides (darkened bg ~12%)
C_GHOST_B   = RGBColor(0x7A, 0x28, 0x96)
C_GHOST_C   = RGBColor(0xA3, 0x28, 0x7A)
C_GHOST_O   = RGBColor(0xCC, 0x6A, 0x18)
C_GHOST_PS  = RGBColor(0x5E, 0xA3, 0x28)

# Block label colors (light version ~60% toward white)
C_LABEL_B   = RGBColor(0xCC, 0xAA, 0xDD)
C_LABEL_C   = RGBColor(0xDD, 0xAA, 0xCC)
C_LABEL_O   = RGBColor(0xFF, 0xCC, 0x99)
C_LABEL_PS  = RGBColor(0xBB, 0xEE, 0x88)

W = Cm(33.87)
H = Cm(19.05)
PL = Cm(1.7)   # left pad
PR = Cm(1.7)   # right pad
CW = W - PL - PR  # content width = 30.47 cm

def darken(color, factor=0.25):
    """Return a darkened version of an RGBColor."""
    h = str(color)
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return RGBColor(int(r * factor), int(g * factor), int(b * factor))

# ── Core primitives ──────────────────────────────────────────────
def new_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def blank_layout(prs):
    return prs.slide_layouts[6]

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill_color,
             line_color=None, line_width=0):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, left, top, width, height, text,
                font_size=18, bold=False, color=C_WHITE,
                align=PP_ALIGN.LEFT, italic=False,
                font_name="Calibri", word_wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.color.rgb = color
    run.font.name   = font_name
    run.font.italic = italic
    return tb

def add_multiline_textbox(slide, left, top, width, height, lines,
                          font_size=18, color=C_WHITE, font_name="Calibri",
                          line_space_pt=None):
    """Add a textbox with multiple paragraphs."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (line_text, line_bold, line_color, line_size) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line_text
        run.font.size = Pt(line_size or font_size)
        run.font.bold = line_bold
        run.font.color.rgb = line_color or color
        run.font.name = font_name
    return tb

def add_slide_num(slide, n):
    add_textbox(slide, W - Cm(3.0), H - Cm(0.95),
                Cm(2.5), Cm(0.7), str(n),
                font_size=11, color=C_DIM, align=PP_ALIGN.RIGHT,
                font_name="Courier New")

def add_footer(slide):
    add_rect(slide, 0, H - Cm(0.12), W, Cm(0.12), fill_color=C_CYAN)

def lbl(slide, text, color=C_CYAN, top=None):
    """Uppercase monospace label — the HTML .lbl class."""
    y = top if top is not None else Cm(1.1)
    add_textbox(slide, PL, y, CW, Cm(0.85),
                text.upper(), font_size=16, bold=True, color=color,
                font_name="Courier New")
    return y + Cm(0.9)

def rule(slide, y, color=C_CYAN):
    """Short 56px-wide rule — the HTML .rule class."""
    add_rect(slide, PL, y, Cm(1.5), Cm(0.1), fill_color=color)
    return y + Cm(0.6)

def h2_text(slide, text, y, color=C_WHITE, size=38, width=None):
    w = width or CW
    add_textbox(slide, PL, y, w, Cm(2.6),
                text, font_size=size, bold=True, color=color)
    return y + Cm(2.6)

def card_box(slide, left, top, width, height,
             accent=C_CYAN, title=None, body=None,
             title_size=18, body_size=18):
    """Dark card with colored top accent strip."""
    add_rect(slide, left, top, width, height, fill_color=C_NAVY2)
    add_rect(slide, left, top, width, Cm(0.18), fill_color=accent)
    cy = top + Cm(0.5)
    if title:
        add_textbox(slide, left + Cm(0.45), cy, width - Cm(0.9), Cm(0.9),
                    title, font_size=title_size, bold=True, color=accent,
                    font_name="Courier New")
        cy += Cm(0.85)
    if body:
        add_textbox(slide, left + Cm(0.45), cy,
                    width - Cm(0.9), top + height - cy - Cm(0.2),
                    body, font_size=body_size, color=C_DIM, word_wrap=True)

def hbox(slide, left, top, width, height,
         accent=C_CYAN, text="", size=18, bg=None):
    """Horizontal info box — HTML .hbox class."""
    bg_col = bg or C_NAVY2
    add_rect(slide, left, top, width, height, fill_color=bg_col)
    add_rect(slide, left, top, Cm(0.18), height, fill_color=accent)
    add_textbox(slide, left + Cm(0.55), top + Cm(0.2),
                width - Cm(0.75), height - Cm(0.4),
                text, font_size=size, color=C_DIM, word_wrap=True)

def find_row(slide, left, top, width, tag, text, tag_color):
    """Finding row — HTML .find class with .tag pill."""
    rh = Cm(1.75)
    add_rect(slide, left, top, width, rh, fill_color=C_NAVY2)
    add_rect(slide, left, top, Cm(0.18), rh, fill_color=tag_color)
    # pill background
    add_rect(slide, left + Cm(0.4), top + Cm(0.3),
             Cm(5.2), Cm(1.1), fill_color=C_DARK_NAVY)
    add_textbox(slide, left + Cm(0.5), top + Cm(0.3),
                Cm(5.1), Cm(1.1),
                tag, font_size=14, bold=True, color=tag_color,
                font_name="Courier New")
    add_textbox(slide, left + Cm(6.2), top + Cm(0.35),
                width - Cm(6.5), Cm(1.1),
                text, font_size=18, color=C_WHITE, word_wrap=True)
    return top + rh + Cm(0.12)

def tier_row(slide, left, top, width, tag, detail, tag_color, bg_color):
    """Tier row — HTML .tier class."""
    rh = Cm(2.0)
    add_rect(slide, left, top, width, rh, fill_color=bg_color)
    # pill
    add_rect(slide, left + Cm(0.4), top + Cm(0.4),
             Cm(5.5), Cm(1.1), fill_color=darken(tag_color, 0.5))
    add_textbox(slide, left + Cm(0.5), top + Cm(0.4),
                Cm(5.4), Cm(1.1),
                tag, font_size=15, bold=True, color=tag_color,
                font_name="Courier New")
    add_textbox(slide, left + Cm(6.3), top + Cm(0.5),
                width - Cm(6.7), Cm(1.1),
                detail, font_size=17, color=C_DIM, word_wrap=True)
    return top + rh + Cm(0.25)

# ── Block transition helper ──────────────────────────────────────
def make_block_transition(prs, slide_num, letter, title, block_num_str,
                          block_label, timer_text, bullets, bg_color,
                          ghost_color, label_color):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, bg_color)

    # Ghost block number — huge, bottom-left, very faint
    add_textbox(slide, Cm(0.5), H - Cm(10.5), Cm(20), Cm(11.0),
                block_num_str,
                font_size=240, bold=True, color=ghost_color,
                font_name="Courier New", align=PP_ALIGN.LEFT)

    # "BLOCK X" label at bottom-left
    add_textbox(slide, Cm(1.5), H - Cm(1.8), Cm(15), Cm(1.2),
                block_label,
                font_size=18, bold=True, color=label_color,
                font_name="Courier New")

    # Left column: giant letter
    add_textbox(slide, Cm(1.5), Cm(2.0), Cm(9.5), Cm(12.0),
                letter,
                font_size=105, bold=True, color=C_WHITE,
                font_name="Courier New")

    # Vertical divider
    add_rect(slide, Cm(11.2), Cm(1.8), Cm(0.07), H - Cm(3.8),
             fill_color=RGBColor(0xCC, 0xCC, 0xCC))

    # Right column: title
    add_textbox(slide, Cm(12.0), Cm(2.5), Cm(21.0), Cm(4.5),
                title, font_size=54, bold=True, color=C_WHITE)

    # Timer badge (rect outline + text)
    add_rect(slide, Cm(12.0), Cm(7.2), Cm(9.5), Cm(1.3),
             fill_color=darken(bg_color, 0.8),
             line_color=C_WHITE, line_width=1)
    add_textbox(slide, Cm(12.2), Cm(7.3), Cm(9.2), Cm(1.1),
                "⏱ " + timer_text,
                font_size=18, bold=True, color=C_WHITE,
                font_name="Courier New")

    # Bullets
    y = Cm(9.2)
    for bullet in bullets:
        add_textbox(slide, Cm(12.0), y, Cm(21.0), Cm(1.2),
                    "●  " + bullet, font_size=19, color=RGBColor(0xEE, 0xEE, 0xEE))
        y += Cm(1.3)

    # Footer label
    add_textbox(slide, W - Cm(22), H - Cm(1.0), Cm(20.5), Cm(0.8),
                "3DCRT Plan Evaluation · Theory Module",
                font_size=13, color=RGBColor(0x99, 0x99, 0x99),
                font_name="Courier New", align=PP_ALIGN.RIGHT)

    add_slide_num(slide, slide_num)
    return slide

# ═══════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════════

def make_title_slide(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_footer(slide)

    # Decorative background circles
    add_rect(slide, W - Cm(14.0), Cm(-4.0), Cm(18.0), Cm(18.0),
             fill_color=RGBColor(0x12, 0x2F, 0x4A))
    shape = slide.shapes[-1]
    shape.line.fill.background()

    add_rect(slide, Cm(-2.0), H - Cm(7.0), Cm(8.0), Cm(8.0),
             fill_color=RGBColor(0x0F, 0x26, 0x3D))
    slide.shapes[-1].line.fill.background()

    # Eyebrow
    add_textbox(slide, PL, Cm(3.5), CW, Cm(0.9),
                "RADIATION ONCOLOGY  ·  UCT / GSH REGISTRAR TRAINING  ·  2026",
                font_size=16, bold=True, color=C_CYAN, font_name="Courier New")

    # Main title
    add_textbox(slide, PL, Cm(4.8), Cm(24), Cm(6.5),
                "What Makes\na Plan Good?",
                font_size=75, bold=True, color=C_WHITE)

    # Subtitle
    add_textbox(slide, PL, Cm(11.8), Cm(24), Cm(1.5),
                "3DCRT Plan Evaluation Workshop — Theory & Practical Module",
                font_size=25, color=C_DIM)

    # Rule + author
    add_rect(slide, PL, Cm(13.8), Cm(0.18), Cm(2.0), fill_color=C_CYAN)
    add_textbox(slide, PL + Cm(0.65), Cm(13.9), Cm(20), Cm(1.1),
                "Tamerone Manasse", font_size=24, bold=True, color=C_WHITE)
    add_textbox(slide, PL + Cm(0.65), Cm(15.0), Cm(20), Cm(0.9),
                "UCT / GSH Registrars  ·  2026", font_size=20, color=C_DIM)

    add_slide_num(slide, n)
    return slide


def make_session_map(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "How this session runs")
    y = rule(slide, y)
    y = h2_text(slide, "2 Hours · 5 Cases · One framework.", y, size=34)
    y -= Cm(0.4)

    rows = [
        ("0:00", "Theory — framework, DVH, OARs, red flags",         "18 min", C_ORANGE),
        ("0:18", "Case 1: Breast — full scorecard exercise",          "22 min", C_CYAN),
        ("0:40", "Case 2: Gynaecology — trade-off",                   "10 min", C_CYAN),
        ("0:50", "Case 3: Prostate — trade-off",                      "10 min", C_CYAN),
        ("1:00", "Case 4: Pituitary — coplanar vs non-coplanar",      "10 min", C_CYAN),
        ("1:10", "Case 5: H&N — 4-field vs 3-field",                  "10 min", C_CYAN),
        ("1:20", "Buffer / questions / overflow",                     "30 min", C_DIM),
        ("1:50", "Wrap-up + key takeaways",                           "10 min", C_GREEN),
    ]
    rh = Cm(1.35)
    for i, (time, label, dur, col) in enumerate(rows):
        bg = C_DARK_NAVY if i % 2 == 0 else C_NAVY2
        add_rect(slide, PL, y, CW, rh, fill_color=bg)
        add_rect(slide, PL, y, Cm(0.18), rh, fill_color=col)
        add_textbox(slide, PL + Cm(0.5), y + Cm(0.15), Cm(3.0), rh - Cm(0.2),
                    time, font_size=16, bold=True, color=col, font_name="Courier New")
        add_textbox(slide, PL + Cm(4.0), y + Cm(0.15), Cm(22.0), rh - Cm(0.2),
                    label, font_size=16, color=C_WHITE)
        add_textbox(slide, W - PR - Cm(4.0), y + Cm(0.15), Cm(3.8), rh - Cm(0.2),
                    dur, font_size=16, color=C_DIM, font_name="Courier New",
                    align=PP_ALIGN.RIGHT)
        y += rh

    y += Cm(0.4)
    hbox(slide, PL, y, CW, Cm(2.0), accent=C_CYAN,
         text="By the end you can:  Apply FCB-CHOPS · Read a DVH · Explain CI & HI · Apply OAR constraints · Triage a plan",
         size=17)

    add_slide_num(slide, n)
    return slide


def make_four_questions(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    # Decorative bg circles
    add_rect(slide, W - Cm(11), Cm(-3), Cm(14), Cm(14),
             fill_color=RGBColor(0x12, 0x2E, 0x4A))
    slide.shapes[-1].line.fill.background()
    add_rect(slide, Cm(-1.5), H - Cm(6), Cm(7), Cm(7),
             fill_color=RGBColor(0x11, 0x2B, 0x45))
    slide.shapes[-1].line.fill.background()

    y = lbl(slide, "The foundation", color=RGBColor(0x9A, 0xAC, 0xBE))
    add_textbox(slide, PL, y, Cm(24), Cm(3.0),
                "A good plan answers four questions — in order.",
                font_size=46, bold=True, color=C_WHITE)
    y += Cm(3.0)

    qs = [
        ("01", "Does it treat the target?",
         "Adequate dose throughout the target — no geographic miss, no cold spot where tumour lives.",
         RGBColor(0x33, 0x55, 0xCC)),
        ("02", "Does it spare normal tissue?",
         "Every OAR within tolerance — judged against the right dose-volume metric for that organ's architecture.",
         RGBColor(0x99, 0x33, 0xBB)),
        ("03", "Is it technically sound?",
         "Sensible isocentre, beam geometry, weighting — the plan is intentionally built, not just generated.",
         RGBColor(0xEE, 0x33, 0x33)),
        ("04", "Is it deliverable & robust?",
         "Achievable on the machine, stable to setup and anatomy variation.",
         RGBColor(0x4E, 0xA7, 0x2E)),
    ]
    col_w = (CW - Cm(1.5)) / 2
    col_h = Cm(4.8)
    positions = [(PL, y), (PL + col_w + Cm(1.5), y),
                 (PL, y + col_h + Cm(0.4)), (PL + col_w + Cm(1.5), y + col_h + Cm(0.4))]

    for (cx, cy), (num, q, detail, col) in zip(positions, qs):
        add_rect(slide, cx, cy, col_w, col_h, fill_color=C_NAVY2)
        add_rect(slide, cx, cy, col_w, Cm(0.12),
                 fill_color=col)
        # Circle number icon
        add_rect(slide, cx + Cm(0.4), cy + Cm(0.5), Cm(2.2), Cm(2.2),
                 fill_color=darken(col, 0.25))
        add_textbox(slide, cx + Cm(0.4), cy + Cm(0.6), Cm(2.2), Cm(2.0),
                    num, font_size=24, bold=True, color=col,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_textbox(slide, cx + Cm(3.0), cy + Cm(0.6), col_w - Cm(3.3), Cm(1.2),
                    q, font_size=21, bold=True, color=C_WHITE)
        add_textbox(slide, cx + Cm(0.4), cy + Cm(2.9), col_w - Cm(0.8),
                    col_h - Cm(3.2),
                    detail, font_size=17, color=C_DIM, word_wrap=True)

    # Footer quote
    add_rect(slide, PL, H - Cm(1.8), CW, Cm(1.4), fill_color=C_DARK_NAVY)
    add_textbox(slide, PL + Cm(0.3), H - Cm(1.7), CW - Cm(0.6), Cm(1.2),
                '"A plan is never perfect — it is the best defensible compromise."',
                font_size=18, italic=True, color=C_DIM, align=PP_ALIGN.CENTER)

    add_slide_num(slide, n)
    return slide


def make_plan_quality_spectrum(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "The reality of plan evaluation")
    y = rule(slide, y)
    y = h2_text(slide, "Bad, acceptable, trade-off — or good?", y, size=34)
    y -= Cm(0.3)

    items = [
        ("CLEARLY BAD",  "Unsafe or unacceptable. Easy to spot, easy to reject. Geographic miss, serial OAR breach, undeliverable MU.",
         C_RED),
        ("ACCEPTABLE",   "Passes all constraints — but a better plan is clearly achievable with reasonable effort. Not wrong, but not optimal.",
         C_ORANGE),
        ("TRADE-OFF",    "Competing priorities pull against each other. Fixing one worsens another. Needs clinical judgement and documentation.",
         C_CYAN),
        ("GOOD",         "Balanced, robust, clinically appropriate. Target covered, OARs spared, technically sound, documented.",
         C_GREEN),
    ]
    col_w = (CW - Cm(1.5)) / 4
    for i, (label, detail, col) in enumerate(items):
        cx = PL + i * (col_w + Cm(0.5))
        ch = Cm(7.5)
        add_rect(slide, cx, y, col_w, ch, fill_color=C_NAVY2)
        add_rect(slide, cx, y, col_w, Cm(0.18), fill_color=col)
        add_textbox(slide, cx + Cm(0.3), y + Cm(0.5), col_w - Cm(0.6), Cm(1.4),
                    label, font_size=20, bold=True, color=col,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_textbox(slide, cx + Cm(0.3), y + Cm(2.1), col_w - Cm(0.6),
                    ch - Cm(2.3),
                    detail, font_size=17, color=C_DIM,
                    align=PP_ALIGN.LEFT, word_wrap=True)

    gy = y + Cm(8.0)
    hbox(slide, PL, gy, CW, Cm(1.9), accent=C_CYAN,
         text='THE GUT-CHECK: "Would I confidently treat my own family member with this plan?" — if no, name exactly what would have to change.',
         size=19)

    add_slide_num(slide, n)
    return slide


def make_fcbchops_evolution(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "The framework — origins and evolution")
    y = rule(slide, y)

    col_w = (CW - Cm(2.5)) / 2
    panel_h = H - y - Cm(1.5)

    # Left panel: CB-CHOP 2017
    lx = PL
    add_rect(slide, lx, y, col_w, panel_h, fill_color=C_NAVY2)
    add_textbox(slide, lx + Cm(0.5), y + Cm(0.4), col_w - Cm(1.0), Cm(0.8),
                "2017 — ORIGINAL", font_size=16, bold=True, color=C_DIM,
                font_name="Courier New")
    add_textbox(slide, lx + Cm(0.5), y + Cm(1.2), col_w - Cm(1.0), Cm(1.5),
                "CB-CHOP", font_size=38, bold=True, color=C_WHITE)
    add_textbox(slide, lx + Cm(0.5), y + Cm(2.8), col_w - Cm(1.0), Cm(0.8),
                "Jimenez et al. · Pract Radiat Oncol 2017",
                font_size=17, italic=True, color=C_DIM)

    orig_letters = [
        ("C", RGBColor(0x66, 0x44, 0xBB), "Contours"),
        ("B", RGBColor(0x99, 0x33, 0xBB), "Beam Arrangement"),
        ("C", RGBColor(0xCC, 0x33, 0x99), "Coverage"),
        ("H", RGBColor(0xEE, 0x33, 0x33), "Heterogeneity / Hotspots"),
        ("O", RGBColor(0xFF, 0x88, 0x22), "Organs at Risk"),
        ("P", RGBColor(0xFF, 0xDD, 0x22), "Prescription"),
    ]
    ly = y + Cm(3.8)
    for lt, lc, lname in orig_letters:
        add_rect(slide, lx + Cm(0.5), ly, Cm(1.8), Cm(1.8),
                 fill_color=darken(lc, 0.2))
        add_textbox(slide, lx + Cm(0.5), ly, Cm(1.8), Cm(1.8),
                    lt, font_size=24, bold=True, color=lc,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_textbox(slide, lx + Cm(2.7), ly + Cm(0.35), col_w - Cm(3.2), Cm(1.1),
                    lname, font_size=18, color=RGBColor(0xBB, 0xBB, 0xBB))
        ly += Cm(2.0)

    # Arrow
    add_textbox(slide, PL + col_w + Cm(0.5), Cm(8.5), Cm(1.5), Cm(2.0),
                "→", font_size=48, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
    add_textbox(slide, PL + col_w + Cm(0.3), Cm(10.5), Cm(1.9), Cm(0.8),
                "+F  +S", font_size=16, bold=True, color=C_CYAN,
                font_name="Courier New", align=PP_ALIGN.CENTER)

    # Right panel: FCB-CHOPS 2024
    rx = PL + col_w + Cm(2.5)
    add_rect(slide, rx, y, col_w, panel_h, fill_color=C_NAVY2)
    add_rect(slide, rx, y, col_w, Cm(0.18), fill_color=C_CYAN)
    add_textbox(slide, rx + Cm(0.5), y + Cm(0.4), col_w - Cm(1.0), Cm(0.8),
                "2024 — EVOLUTION", font_size=16, bold=True, color=C_CYAN,
                font_name="Courier New")
    add_textbox(slide, rx + Cm(0.5), y + Cm(1.2), col_w - Cm(1.0), Cm(1.5),
                "FCB-CHOPS", font_size=38, bold=True, color=C_WHITE)
    add_textbox(slide, rx + Cm(0.5), y + Cm(2.8), col_w - Cm(1.0), Cm(0.9),
                "Weisman et al. · Adv Radiat Oncol 2024  ·  PMID 40017913",
                font_size=16, italic=True, color=C_DIM)

    new_letters = [
        ("F", RGBColor(0x33, 0x55, 0xCC), "Fusion / Imaging", True),
        ("C", RGBColor(0x66, 0x44, 0xBB), "Contours", False),
        ("B", RGBColor(0x99, 0x33, 0xBB), "Beam Arrangement", False),
        ("C", RGBColor(0xCC, 0x33, 0x99), "Coverage", False),
        ("H", RGBColor(0xEE, 0x33, 0x33), "Heterogeneity / Hotspots", False),
        ("O", RGBColor(0xFF, 0x88, 0x22), "Organs at Risk", False),
        ("P", RGBColor(0xFF, 0xDD, 0x22), "Prescription", False),
        ("S", RGBColor(0x77, 0xCC, 0x33), "Summation", True),
    ]
    ry = y + Cm(3.8)
    for lt, lc, lname, is_new in new_letters:
        add_rect(slide, rx + Cm(0.5), ry, Cm(1.8), Cm(1.8),
                 fill_color=darken(lc, 0.2))
        add_textbox(slide, rx + Cm(0.5), ry, Cm(1.8), Cm(1.8),
                    lt, font_size=24, bold=True, color=lc,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_textbox(slide, rx + Cm(2.7), ry + Cm(0.35), col_w - Cm(5.5), Cm(1.1),
                    lname, font_size=18, color=RGBColor(0xBB, 0xBB, 0xBB))
        if is_new:
            add_rect(slide, rx + col_w - Cm(3.0), ry + Cm(0.35),
                     Cm(2.5), Cm(1.0),
                     fill_color=darken(lc, 0.2))
            add_textbox(slide, rx + col_w - Cm(3.0), ry + Cm(0.3),
                        Cm(2.5), Cm(1.1),
                        "NEW", font_size=15, bold=True, color=lc,
                        font_name="Courier New", align=PP_ALIGN.CENTER)
        ry += Cm(1.9)

    add_slide_num(slide, n)
    return slide


def make_fcbchops_framework(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "A systematic framework")
    y = rule(slide, y)
    # Coloured letter heading
    add_textbox(slide, PL, y, CW, Cm(1.8),
                "FCB-CHOPS — eight checkpoints, every plan",
                font_size=30, bold=True, color=C_WHITE)
    y += Cm(1.9)

    checkpoints = [
        ("F", RGBColor(0x33, 0x55, 0xCC), "Fusion / Imaging",
         "Verify image fusion accuracy — correct target & OAR anatomy"),
        ("C", RGBColor(0x66, 0x44, 0xBB), "Contours",
         "All target & OAR contours complete and anatomically correct"),
        ("B", RGBColor(0x99, 0x33, 0xBB), "Beam Arrangement",
         "Beam geometry and entry points appropriate for the target"),
        ("C", RGBColor(0xCC, 0x33, 0x99), "Coverage",
         "Prescribed dose adequately covers the target volume"),
        ("H", RGBColor(0xEE, 0x33, 0x33), "Heterogeneity/Hotspots",
         "Internal dose variation — acceptable high and low dose regions"),
        ("O", RGBColor(0xFF, 0x88, 0x22), "Organs at Risk",
         "OAR doses respect clinical constraints for each organ"),
        ("P", RGBColor(0xFF, 0xDD, 0x22), "Prescription",
         "Correct dose, fractionation, normalisation, laterality documented"),
        ("S", RGBColor(0x77, 0xCC, 0x33), "Summation",
         "Cumulative dose from all courses — or explicitly acknowledge none"),
    ]

    col_w = (CW - Cm(1.5)) / 4
    cph = Cm(3.2)
    for i, (lt, lc, name, desc) in enumerate(checkpoints):
        cx = PL + (i % 4) * (col_w + Cm(0.5))
        cy = y + (i // 4) * (cph + Cm(0.3))
        add_rect(slide, cx, cy, col_w, cph, fill_color=C_NAVY2)
        add_rect(slide, cx, cy, col_w, Cm(0.12),
                 fill_color=lc)
        add_textbox(slide, cx + Cm(0.4), cy + Cm(0.3), col_w - Cm(0.8), Cm(1.2),
                    lt, font_size=40, bold=True, color=lc, font_name="Courier New")
        add_textbox(slide, cx + Cm(0.4), cy + Cm(1.4), col_w - Cm(0.8), Cm(0.8),
                    name, font_size=17, bold=True, color=C_WHITE)
        add_textbox(slide, cx + Cm(0.4), cy + Cm(2.2), col_w - Cm(0.8), Cm(0.9),
                    desc, font_size=15, color=C_DIM, word_wrap=True)

    gy = y + 2 * (cph + Cm(0.3)) + Cm(0.3)
    col_w2 = (CW - Cm(0.8)) / 2
    hbox(slide, PL, gy, col_w2, Cm(1.6), accent=C_CYAN,
         text="Start with Prescription. It defines treatment intent — confirm it before anything else.",
         size=17)
    hbox(slide, PL + col_w2 + Cm(0.8), gy, col_w2, Cm(1.6), accent=C_ORANGE,
         text="The letters are a safety net, not a script. Work in whatever order the plan demands.",
         size=17)

    add_slide_num(slide, n)
    return slide


def make_field_arrangement(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block B · Beam Arrangement")
    y = h2_text(slide, "Every beam should reach the target by the most direct, safest route",
                y, size=30)
    y -= Cm(0.2)

    scenarios = [
        ("UNAVOIDABLE", "ACCEPTABLE",
         "4-field box — beams transit OARs",
         "Whole pelvis, bilateral treatment — OAR transit is geometry-driven. Document field weighting; manage with DVH.",
         C_ORANGE),
        ("AVOIDABLE", "RE-ANGLE",
         "Beam transits OAR before reaching PTV",
         "Left lateral must pass through OAR en route to PTV. This is avoidable — enter from the contralateral side.",
         C_RED),
        ("SOLUTION", "IPSILATERAL",
         "Enter from the OAR-free side",
         "AP + ipsilateral lateral — OAR on the contralateral side receives no direct beam dose.",
         C_GREEN),
    ]
    col_w = (CW - Cm(1.0)) / 3
    sh = H - y - Cm(2.8)
    for i, (tag1, tag2, headline, detail, col) in enumerate(scenarios):
        cx = PL + i * (col_w + Cm(0.5))
        add_rect(slide, cx, y, col_w, sh, fill_color=C_NAVY2)
        add_rect(slide, cx, y, col_w, Cm(0.18), fill_color=col)
        # Tags
        tw = Cm(4.5)
        add_rect(slide, cx + Cm(0.3), y + Cm(0.4), tw, Cm(1.0),
                 fill_color=darken(col, 0.2))
        add_textbox(slide, cx + Cm(0.3), y + Cm(0.4), tw, Cm(1.0),
                    tag1, font_size=14, bold=True, color=col,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_rect(slide, cx + tw + Cm(0.5), y + Cm(0.4), tw, Cm(1.0),
                 fill_color=C_DARK_NAVY)
        add_textbox(slide, cx + tw + Cm(0.5), y + Cm(0.4), tw, Cm(1.0),
                    tag2, font_size=14, bold=True, color=C_DIM,
                    font_name="Courier New", align=PP_ALIGN.CENTER)
        add_textbox(slide, cx + Cm(0.3), y + Cm(1.7), col_w - Cm(0.6), Cm(1.2),
                    headline, font_size=19, bold=True, color=col if col != C_RED else RGBColor(0xF0, 0x6A, 0x6A))
        add_textbox(slide, cx + Cm(0.3), y + Cm(3.1), col_w - Cm(0.6),
                    sh - Cm(3.4),
                    detail, font_size=17, color=C_DIM, word_wrap=True)

    # Footer callout
    add_rect(slide, PL, H - Cm(2.2), CW, Cm(1.6),
             fill_color=RGBColor(0x1A, 0x2E, 0x14))
    add_textbox(slide, PL + Cm(0.4), H - Cm(2.1), CW - Cm(0.8), Cm(1.4),
                "Before accepting any OAR transit — ask: is there a beam arrangement that avoids it? If yes, use it. If no, document why.",
                font_size=18, color=RGBColor(0xFF, 0xDD, 0x44), align=PP_ALIGN.CENTER)

    add_slide_num(slide, n)
    return slide


def make_technical_quality(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block B · Technical Quality")
    y = rule(slide, y)

    col_w = (CW - Cm(2.0)) / 2
    add_textbox(slide, PL, y, col_w, Cm(1.4),
                "Technical red flags", font_size=28, bold=True, color=C_RED)
    add_textbox(slide, PL + col_w + Cm(2.0), y, col_w, Cm(1.4),
                "Technical quality check", font_size=28, bold=True, color=C_GREEN)
    y += Cm(1.6)

    flags = [
        ("RISK",  "Isocentre in lung, air, or density interface — dose calculations unreliable", C_RED),
        ("RISK",  "Beam transiting directly through a critical serial OAR — re-angle if any workable alternative exists", C_RED),
        ("RISK",  "Hot spot outside the PTV — in skin, bowel, or OAR it is unacceptable regardless of magnitude", C_RED),
        ("MAJOR", "MLC clipped tight to PTV edge — no penumbra margin means a cold rim on every beam", C_ORANGE),
        ("MAJOR", "Normalisation point in unstable tissue — lung, air, interface — every DVH number unreliable", C_ORANGE),
    ]
    fy = y
    for tag, text, col in flags:
        add_rect(slide, PL, fy, col_w, Cm(1.7), fill_color=C_NAVY2)
        add_rect(slide, PL, fy, Cm(0.18), Cm(1.7), fill_color=col)
        add_rect(slide, PL + Cm(0.4), fy + Cm(0.3),
                 Cm(3.5), Cm(1.1), fill_color=C_DARK_NAVY)
        add_textbox(slide, PL + Cm(0.5), fy + Cm(0.3), Cm(3.4), Cm(1.1),
                    tag, font_size=13, bold=True, color=col,
                    font_name="Courier New")
        add_textbox(slide, PL + Cm(4.3), fy + Cm(0.3),
                    col_w - Cm(4.6), Cm(1.1),
                    text, font_size=16, color=C_WHITE, word_wrap=True)
        fy += Cm(1.85)

    checks = [
        "Isocentre — centred in PTV; not in lung or density interface",
        "Beam weighting — re-weight to pull hot spot into target",
        "Wedges — tilt isodoses to compensate for sloping surface",
        "Field-in-field (FiF) — sub-segments block emerging hot spots; routine in breast",
        "Monitor units — sanity check; excess MU adds scatter to whole patient",
    ]
    cx = PL + col_w + Cm(2.0)
    cy = y
    for check in checks:
        add_rect(slide, cx, cy, col_w, Cm(1.7), fill_color=C_NAVY2)
        add_rect(slide, cx, cy, Cm(0.18), Cm(1.7), fill_color=C_GREEN)
        add_textbox(slide, cx + Cm(0.5), cy + Cm(0.35),
                    col_w - Cm(0.8), Cm(1.1),
                    "✓  " + check, font_size=17, color=C_WHITE)
        cy += Cm(1.85)

    add_slide_num(slide, n)
    return slide


def make_hotspots_diagram(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block B · Technical Quality")
    y = rule(slide, y)
    y = h2_text(slide, "Hot spots, dose gradient & integral dose", y, size=34)
    y -= Cm(0.3)

    col_w = (CW - Cm(2.0)) / 2
    points = [
        ("Hot spot — location, not just size.",
         "A 107% inside the PTV is fine; the same 107% in skin or an OAR is not. Location governs acceptability."),
        ("Dose gradient",
         "Dose should fall steeply between the PTV edge and any adjacent OAR. Check isodose slices for every beam entry."),
        ("Integral dose / MU",
         "Sanity-check the monitor units. Excess MU adds scatter and integral dose to the whole patient."),
    ]
    py = y
    for title, detail in points:
        add_textbox(slide, PL, py, col_w, Cm(0.9),
                    "▶  " + title, font_size=21, bold=True, color=C_ORANGE)
        add_textbox(slide, PL, py + Cm(0.9), col_w, Cm(1.2),
                    detail, font_size=18, color=C_DIM, word_wrap=True)
        py += Cm(2.4)

    # Right side: Inside vs Outside PTV boxes
    rx = PL + col_w + Cm(2.0)
    add_textbox(slide, rx, y, col_w, Cm(0.9),
                "Don't just ask HOW BIG — ask WHERE it sits",
                font_size=18, italic=True, color=C_DIM, align=PP_ALIGN.CENTER)
    y2 = y + Cm(1.1)
    bh = (H - y2 - Cm(2.8)) / 2

    # Inside PTV box
    add_rect(slide, rx, y2, col_w, bh, fill_color=C_NAVY2)
    add_rect(slide, rx, y2, col_w, Cm(0.18), fill_color=C_GREEN)
    add_textbox(slide, rx + Cm(0.4), y2 + Cm(0.4), col_w - Cm(0.8), Cm(1.1),
                "Hot spot INSIDE the PTV — acceptable",
                font_size=20, bold=True, color=C_GREEN)
    add_textbox(slide, rx + Cm(0.4), y2 + Cm(1.6), col_w - Cm(0.8),
                bh - Cm(2.0),
                "107% inside the target is expected in 3DCRT. The prescription point is normalised to ~100%; a modest hotspot in the tumour itself is clinically acceptable.\n\nMonitor: ensure D2% < 107%",
                font_size=18, color=C_DIM, word_wrap=True)

    # Outside PTV box
    y3 = y2 + bh + Cm(0.3)
    add_rect(slide, rx, y3, col_w, bh, fill_color=C_NAVY2)
    add_rect(slide, rx, y3, col_w, Cm(0.18), fill_color=C_RED)
    add_textbox(slide, rx + Cm(0.4), y3 + Cm(0.4), col_w - Cm(0.8), Cm(1.1),
                "Hot spot OUTSIDE the PTV — concern",
                font_size=20, bold=True, color=RGBColor(0xF0, 0x6A, 0x6A))
    add_textbox(slide, rx + Cm(0.4), y3 + Cm(1.6), col_w - Cm(0.8),
                bh - Cm(2.0),
                "In OAR or skin = PROBLEM regardless of magnitude.\n\nCauses: deep isocentre, unweighted beams, density heterogeneity.\n\nAction: move iso, adjust weighting, check BEV.",
                font_size=18, color=C_DIM, word_wrap=True)

    # Quote
    add_rect(slide, PL, H - Cm(2.0), CW, Cm(1.4), fill_color=C_DARK_NAVY)
    add_textbox(slide, PL + Cm(0.4), H - Cm(1.9), CW - Cm(0.8), Cm(1.2),
                '"The same hot spot is acceptable inside the PTV and a problem outside it"',
                font_size=18, italic=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

    add_slide_num(slide, n)
    return slide


def make_dvh_volume_hierarchy(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block C · Target Coverage — DVH & Volumes")
    y = rule(slide, y)

    col_w = (CW - Cm(2.0)) / 2
    add_textbox(slide, PL, y, col_w, Cm(1.5),
                "First: name your volumes", font_size=28, bold=True, color=C_WHITE)
    add_textbox(slide, PL + col_w + Cm(2.0), y, col_w, Cm(1.5),
                "The DVH — your primary instrument", font_size=28, bold=True, color=C_WHITE)
    y += Cm(1.7)

    volumes = [
        ("GTV", "Gross Tumour Volume — what you can see, palpate or image",          C_RED),
        ("CTV", "GTV + suspected microscopic spread — the true clinical target",      C_ORANGE),
        ("ITV", "CTV + internal motion margin (ICRU 62)",                             C_CYAN),
        ("PTV", "ITV/CTV + set-up uncertainty — the geometric safety margin",         C_GREEN),
    ]
    vh = Cm(1.8)
    vy = y
    for abbr, desc, col in volumes:
        add_rect(slide, PL, vy, col_w, vh, fill_color=C_NAVY2)
        add_rect(slide, PL, vy, Cm(0.18), vh, fill_color=col)
        add_textbox(slide, PL + Cm(0.5), vy + Cm(0.25), Cm(2.5), Cm(1.2),
                    abbr, font_size=24, bold=True, color=col, font_name="Courier New")
        add_textbox(slide, PL + Cm(3.2), vy + Cm(0.3), col_w - Cm(3.5), Cm(1.2),
                    desc, font_size=17, color=C_DIM, word_wrap=True)
        vy += vh + Cm(0.15)

    # Prescribe-to-PTV note
    hbox(slide, PL, vy + Cm(0.2), col_w, Cm(1.8), accent=C_ORANGE,
         text="We prescribe to and evaluate the PTV — but the clinical aim is to cover the CTV.",
         size=17)

    # Right col: DVH info
    cx = PL + col_w + Cm(2.0)
    dvh_points = [
        ("Collapses 3-D distribution into one readable curve — every point: what volume gets ≥ this dose?"),
        ("Read UP from a dose → V-metric (V95, V20 …)"),
        ("Read ACROSS from a volume → D-metric (D95, D2 … ICRU 83 preferred)"),
        ("No spatial information — never tells you WHERE the cold or hot spot sits"),
    ]
    dy = y
    for i, pt in enumerate(dvh_points):
        col = RGBColor(0xF0, 0x6A, 0x6A) if i == 3 else C_DIM
        add_rect(slide, cx, dy, col_w, Cm(1.8), fill_color=C_NAVY2)
        add_rect(slide, cx, dy, Cm(0.18), Cm(1.8),
                 fill_color=C_RED if i == 3 else C_CYAN)
        add_textbox(slide, cx + Cm(0.5), dy + Cm(0.3), col_w - Cm(0.8), Cm(1.3),
                    pt, font_size=17, color=col, word_wrap=True)
        dy += Cm(1.95)

    hbox(slide, cx, dy + Cm(0.1), col_w, Cm(1.8),
         accent=C_RED, bg=RGBColor(0x28, 0x0E, 0x0E),
         text="The DVH screens; the isodose slices diagnose. Always scroll every plane before you trust the DVH.",
         size=17)

    add_slide_num(slide, n)
    return slide


def make_coverage_metrics(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block C · Coverage Metrics & Homogeneity")
    y = rule(slide, y)

    col_w = (CW - Cm(2.0)) / 2
    add_textbox(slide, PL, y, col_w, Cm(1.5),
                "The numbers that matter", font_size=28, bold=True, color=C_WHITE)
    y += Cm(1.7)

    metrics = [
        ("D95%",  C_CYAN,   "Volume of PTV receiving ≥95% of prescription.\nStandard aim: D95 ≥ 95%."),
        ("D2%",   C_ORANGE, "Near-maximum — hotspot criterion.\nICRU 83: D2% < 107%. Hotspot should sit inside the target."),
        ("HI",    C_GREEN,  "(D2−D98) ÷ D50. Lower = flatter.\n< 0.1 good  ·  < 0.15 acceptable."),
        ("CI",    C_DIM,    "PIV ÷ TV. Ideal 1.0–1.2.\nHigher values expected and acceptable in 3DCRT."),
    ]
    mh = Cm(3.5)
    for i, (abbr, col, desc) in enumerate(metrics):
        mx = PL + (i % 2) * (col_w + Cm(2.0))
        my = y + (i // 2) * (mh + Cm(0.3))
        add_rect(slide, mx, my, col_w, mh, fill_color=C_NAVY2)
        add_rect(slide, mx, my, col_w, Cm(0.18), fill_color=col)
        add_textbox(slide, mx + Cm(0.4), my + Cm(0.4), Cm(4.0), Cm(1.5),
                    abbr, font_size=36, bold=True, color=col,
                    font_name="Courier New")
        add_textbox(slide, mx + Cm(0.4), my + Cm(2.0), col_w - Cm(0.8),
                    mh - Cm(2.3),
                    desc, font_size=18, color=C_DIM, word_wrap=True)

    gy = y + 2 * (mh + Cm(0.3)) + Cm(0.3)
    col_w2 = (CW - Cm(0.8)) / 2
    hbox(slide, PL, gy, col_w2, Cm(2.4), accent=C_CYAN,
         text="Why D98/D2 not Dmin/Dmax? Dmin and Dmax are single-voxel point doses — statistically noisy, shifting with grid size. ICRU 83 replaced them with D98% and D2%: clinically equivalent but stable plan-to-plan.",
         size=17)
    hbox(slide, PL + col_w2 + Cm(0.8), gy, col_w2, Cm(2.4), accent=C_ORANGE,
         text="Coverage is negotiable where PTV abuts a serial OAR — accept a small underdose at the overlap. Safety wins. Know where the cold sits.",
         size=17)

    add_slide_num(slide, n)
    return slide


def make_serial_parallel(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "Block O · OAR Constraints")
    y = rule(slide, y)
    y = h2_text(slide, "Architecture decides the metric — getting this wrong is a class of error",
                y, size=30)
    y -= Cm(0.2)

    col_w = (CW - Cm(2.0)) / 2
    ph = H - y - Cm(3.2)

    # Serial panel
    sx = PL
    add_rect(slide, sx, y, col_w, ph, fill_color=C_NAVY2)
    add_rect(slide, sx, y, col_w, Cm(0.18), fill_color=C_ORANGE)
    add_textbox(slide, sx + Cm(0.5), y + Cm(0.5), col_w - Cm(1.0), Cm(1.4),
                "Serial organs", font_size=32, bold=True, color=C_ORANGE)
    add_rect(slide, sx + Cm(0.5), y + Cm(2.0), col_w - Cm(1.0), Cm(1.0),
             fill_color=RGBColor(0x3A, 0x1E, 0x0E))
    add_textbox(slide, sx + Cm(0.6), y + Cm(2.05), col_w - Cm(1.2), Cm(0.9),
                "Metric: Dmax or D1cc",
                font_size=18, bold=True, color=C_ORANGE, font_name="Courier New")
    add_textbox(slide, sx + Cm(0.5), y + Cm(3.3), col_w - Cm(1.0), Cm(1.5),
                "One failed sub-unit = organ failure. Even a small dose spike above tolerance causes irreversible damage.",
                font_size=18, color=C_DIM, word_wrap=True)
    add_textbox(slide, sx + Cm(0.5), y + Cm(5.0), col_w - Cm(1.0), Cm(0.9),
                "Spinal cord  ·  Brainstem  ·  Optic chiasm  ·  Optic nerves",
                font_size=17, color=C_DIM)
    hbox(slide, sx + Cm(0.5), y + Cm(6.0), col_w - Cm(1.0), Cm(1.3),
         accent=C_ORANGE,
         text="Apply a PRV margin. Evaluate dose to the PRV, not the bare organ contour.",
         size=17)
    add_textbox(slide, sx + Cm(0.5), y + Cm(7.6), col_w - Cm(1.0), Cm(0.9),
                "Cord ≤45 Gy  ·  Brainstem ≤54 Gy  ·  Chiasm ≤54 Gy",
                font_size=18, bold=True, color=C_WHITE, font_name="Courier New")

    # Parallel panel
    px = PL + col_w + Cm(2.0)
    add_rect(slide, px, y, col_w, ph, fill_color=C_NAVY2)
    add_rect(slide, px, y, col_w, Cm(0.18), fill_color=C_CYAN)
    add_textbox(slide, px + Cm(0.5), y + Cm(0.5), col_w - Cm(1.0), Cm(1.4),
                "Parallel organs", font_size=32, bold=True, color=C_CYAN)
    add_rect(slide, px + Cm(0.5), y + Cm(2.0), col_w - Cm(1.0), Cm(1.0),
             fill_color=RGBColor(0x0E, 0x2E, 0x3A))
    add_textbox(slide, px + Cm(0.6), y + Cm(2.05), col_w - Cm(1.2), Cm(0.9),
                "Metric: Mean dose or Vx",
                font_size=18, bold=True, color=C_CYAN, font_name="Courier New")
    add_textbox(slide, px + Cm(0.5), y + Cm(3.3), col_w - Cm(1.0), Cm(1.5),
                "Redundant sub-units — organ tolerates partial irradiation. Widespread low dose depletes functional reserve.",
                font_size=18, color=C_DIM, word_wrap=True)
    add_textbox(slide, px + Cm(0.5), y + Cm(5.0), col_w - Cm(1.0), Cm(0.9),
                "Lung  ·  Liver  ·  Kidney  ·  Parotid gland  ·  Bowel",
                font_size=17, color=C_DIM)
    hbox(slide, px + Cm(0.5), y + Cm(6.0), col_w - Cm(1.0), Cm(1.3),
         accent=C_CYAN,
         text="Mean dose predicts function loss; Vx sets the toxicity threshold. Both matter.",
         size=17)
    add_textbox(slide, px + Cm(0.5), y + Cm(7.6), col_w - Cm(1.0), Cm(0.9),
                "Lung V20 <30%  ·  Parotid mean <26 Gy  ·  Rectum V70 <20%",
                font_size=18, bold=True, color=C_WHITE, font_name="Courier New")

    hbox(slide, PL, H - Cm(2.6), CW, Cm(1.8), accent=C_DIM,
         text="Constraints are not absolutes. Re-irradiation, concurrent chemotherapy, comorbidities, and treatment intent all move the line.",
         size=17)

    add_slide_num(slide, n)
    return slide


def make_tiers_scorecard(prs, n):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, "P · S — Bringing it together")
    y = rule(slide, y)

    col_w = (CW - Cm(2.0)) / 2
    add_textbox(slide, PL, y, col_w, Cm(1.5),
                "Sort every finding into a tier", font_size=28, bold=True, color=C_WHITE)
    add_textbox(slide, PL + col_w + Cm(2.0), y, col_w, Cm(1.5),
                "Your task — the evaluation routine", font_size=28, bold=True, color=C_WHITE)
    y += Cm(1.7)

    ty = y
    ty = tier_row(slide, PL, ty, col_w,
                  "CRITICAL",
                  "Re-plan required — geographic miss · serial OAR above hard limit · hot spot in critical organ",
                  C_RED, RGBColor(0x28, 0x0E, 0x0E))
    ty = tier_row(slide, PL, ty, col_w,
                  "MAJOR",
                  "Correct or justify — parallel OAR over constraint · V95 90–95% · avoidable beam through OAR",
                  C_ORANGE, RGBColor(0x28, 0x18, 0x0A))
    ty = tier_row(slide, PL, ty, col_w,
                  "MINOR",
                  "Comment only — slightly high MU · non-ideal isocentre · minor inhomogeneity",
                  C_CYAN, RGBColor(0x0A, 0x1E, 0x28))

    add_rect(slide, PL, ty + Cm(0.1), col_w, Cm(2.5), fill_color=C_NAVY2)
    add_rect(slide, PL, ty + Cm(0.1), col_w, Cm(0.18), fill_color=C_RED)
    add_textbox(slide, PL + Cm(0.4), ty + Cm(0.5), col_w - Cm(0.8), Cm(0.9),
                "Reject on sight if…", font_size=18, bold=True,
                color=RGBColor(0xF0, 0x6A, 0x6A))
    add_textbox(slide, PL + Cm(0.4), ty + Cm(1.4), col_w - Cm(0.8), Cm(1.1),
                "1. Cold dose reaching into the CTV  ·  2. Serial OAR exceeds hard tolerance  ·  3. Plan is unsafe",
                font_size=17, color=C_WHITE, word_wrap=True)

    # Right column
    rx = PL + col_w + Cm(2.0)
    routine = [
        "Prescription → Contours → Beams → Isodose visual → Coverage → Hotspots → DVH → OARs → Decision",
        "Visual review BEFORE the DVH — scroll every plane first; the DVH cannot tell you WHERE",
        "Name your top 3 concerns before the ideal answer is shown — this separates critical from minor thinking",
        "Fix problems in the right order: contours → geometry → normalisation → optimisation",
    ]
    ry = y
    for rt in routine:
        add_rect(slide, rx, ry, col_w, Cm(2.0), fill_color=C_NAVY2)
        add_rect(slide, rx, ry, Cm(0.18), Cm(2.0), fill_color=C_CYAN)
        add_textbox(slide, rx + Cm(0.5), ry + Cm(0.3), col_w - Cm(0.8), Cm(1.5),
                    rt, font_size=17, color=C_DIM, word_wrap=True)
        ry += Cm(2.15)

    hbox(slide, rx, ry, col_w, Cm(1.5), accent=C_CYAN,
         text='"Would I confidently treat my own family member with this plan?"',
         size=18)

    add_slide_num(slide, n)
    return slide


# ── Case helpers ─────────────────────────────────────────────────

def make_case_opener(prs, n, case_num, case_title, rx, technique, oars, timer=None):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_footer(slide)

    # Decorative background circles
    add_rect(slide, W - Cm(12), Cm(-2.5), Cm(13), Cm(13),
             fill_color=RGBColor(0x10, 0x25, 0x3C))
    slide.shapes[-1].line.fill.background()
    add_rect(slide, Cm(-1.5), H - Cm(6), Cm(7), Cm(7),
             fill_color=RGBColor(0x0E, 0x22, 0x38))
    slide.shapes[-1].line.fill.background()

    # Giant decorative case number — faint cyan, bottom-right
    num_str = f"0{case_num}"
    add_textbox(slide, W - Cm(14), H - Cm(12), Cm(14), Cm(12),
                num_str, font_size=165, bold=True,
                color=RGBColor(0x0C, 0x50, 0x70),
                font_name="Courier New", align=PP_ALIGN.RIGHT)

    # Label
    y = lbl(slide, f"Case {case_num}  ·  {['Full Scorecard Exercise','Gynaecology','Prostate','Pituitary','Head & Neck'][case_num-1]}")
    y = rule(slide, y)

    # Title
    add_textbox(slide, PL, y, Cm(26), Cm(3.0),
                case_title, font_size=52, bold=True, color=C_WHITE)
    y += Cm(3.3)

    # Info cards
    card_defs = [
        ("PRESCRIPTION", rx,        C_CYAN),
        ("TECHNIQUE",    technique,  C_CYAN),
        ("KEY OARS",     "  ·  ".join(oars), C_CYAN),
    ]
    cw3 = (CW - Cm(1.0)) / 3
    for i, (ctitle, ctext, col) in enumerate(card_defs):
        cx = PL + i * (cw3 + Cm(0.5))
        add_rect(slide, cx, y, cw3, Cm(3.5), fill_color=C_NAVY2)
        add_rect(slide, cx, y, cw3, Cm(0.18), fill_color=col)
        add_textbox(slide, cx + Cm(0.4), y + Cm(0.35), cw3 - Cm(0.8), Cm(0.7),
                    ctitle, font_size=16, bold=True, color=col,
                    font_name="Courier New")
        add_textbox(slide, cx + Cm(0.4), y + Cm(1.1), cw3 - Cm(0.8), Cm(2.2),
                    ctext, font_size=18, bold=(ctitle == "PRESCRIPTION"),
                    color=C_WHITE if ctitle == "PRESCRIPTION" else C_DIM,
                    word_wrap=True)

    # Timer badge
    if timer:
        ty = y + Cm(4.2)
        add_rect(slide, PL, ty, Cm(16), Cm(1.4),
                 fill_color=RGBColor(0x1A, 0x3A, 0x18),
                 line_color=C_GREEN, line_width=1)
        add_textbox(slide, PL + Cm(0.4), ty + Cm(0.2), Cm(15.5), Cm(1.0),
                    timer, font_size=18, bold=True, color=C_GREEN,
                    font_name="Courier New")

    add_slide_num(slide, n)
    return slide


def make_case_review(prs, n, case_label, guide_items):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    col_w = (CW - Cm(2.0)) / 2
    y = lbl(slide, f"{case_label}  ·  What do you see?")
    y = rule(slide, y)

    # Left: image placeholder
    add_rect(slide, PL, y, col_w, H - y - Cm(1.5),
             fill_color=RGBColor(0x0C, 0x20, 0x35))
    add_rect(slide, PL, y, col_w, H - y - Cm(1.5),
             fill_color=RGBColor(0x0C, 0x20, 0x35),
             line_color=RGBColor(0x1A, 0x50, 0x70), line_width=1)
    ph_y = (y + H - Cm(1.5)) / 2 - Cm(1.5)
    add_textbox(slide, PL + Cm(0.5), ph_y, col_w - Cm(1.0), Cm(1.0),
                "[ DVH  +  Isodose images go here ]",
                font_size=18, color=RGBColor(0x1A, 0x60, 0x80),
                font_name="Courier New", align=PP_ALIGN.CENTER)

    # Right: guide
    rx = PL + col_w + Cm(2.0)
    add_textbox(slide, rx, y, col_w, Cm(1.5),
                "What do you see?", font_size=28, bold=True, color=C_WHITE)
    add_textbox(slide, rx, y + Cm(1.6), col_w, Cm(0.9),
                "Score every FCB-CHOPS row before we discuss.",
                font_size=18, color=C_DIM, italic=True)

    add_rect(slide, rx, y + Cm(2.7), col_w, Cm(0.18), fill_color=C_CYAN)
    gy = y + Cm(3.1)
    add_textbox(slide, rx, gy, col_w, Cm(0.8),
                "Guide your eye to…",
                font_size=18, bold=True, color=C_CYAN, font_name="Courier New")
    gy += Cm(1.0)
    for item in guide_items:
        add_rect(slide, rx, gy, col_w, Cm(1.4), fill_color=C_NAVY2)
        add_rect(slide, rx, gy, Cm(0.18), Cm(1.4), fill_color=C_CYAN)
        add_textbox(slide, rx + Cm(0.5), gy + Cm(0.2), col_w - Cm(0.7), Cm(1.0),
                    item, font_size=17, color=C_DIM, word_wrap=True)
        gy += Cm(1.55)

    add_slide_num(slide, n)
    return slide


def make_case_findings(prs, n, case_label, findings):
    """findings = list of (tag, text, color)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    col_w = (CW - Cm(2.0)) / 2
    y = lbl(slide, f"{case_label}  ·  Findings")
    y = rule(slide, y)

    half = len(findings) // 2 + len(findings) % 2
    left_finds = findings[:half]
    right_finds = findings[half:]

    fy = y
    for tag, text, col in left_finds:
        fy = find_row(slide, PL, fy, col_w, tag, text, col)

    ry = y
    for tag, text, col in right_finds:
        ry = find_row(slide, PL + col_w + Cm(2.0), ry, col_w, tag, text, col)

    add_slide_num(slide, n)
    return slide


def make_case_findings_full(prs, n, case_label, left_findings, left_title,
                             right_content, right_title):
    """Findings slide with left column findings + right column custom content."""
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    col_w = (CW - Cm(2.0)) / 2
    y = lbl(slide, f"{case_label}  ·  Findings")
    y = rule(slide, y)

    add_textbox(slide, PL, y, col_w, Cm(1.3),
                left_title, font_size=24, bold=True, color=C_WHITE)
    add_textbox(slide, PL + col_w + Cm(2.0), y, col_w, Cm(1.3),
                right_title, font_size=24, bold=True, color=C_WHITE)
    y += Cm(1.5)

    fy = y
    for tag, text, col in left_findings:
        fy = find_row(slide, PL, fy, col_w, tag, text, col)

    add_slide_num(slide, n)
    return slide, y


def make_case_solutions(prs, n, case_label, solutions, footer=None):
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, f"{case_label}  ·  Solutions")
    y = rule(slide, y)

    for i, sol in enumerate(solutions):
        add_rect(slide, PL, y, CW, Cm(1.75), fill_color=C_NAVY2)
        add_rect(slide, PL, y, Cm(0.18), Cm(1.75), fill_color=C_GREEN)
        # Number circle
        add_rect(slide, PL + Cm(0.4), y + Cm(0.3), Cm(1.1), Cm(1.1),
                 fill_color=RGBColor(0x14, 0x3E, 0x28))
        add_textbox(slide, PL + Cm(0.4), y + Cm(0.3), Cm(1.1), Cm(1.1),
                    str(i+1), font_size=18, bold=True, color=C_GREEN,
                    align=PP_ALIGN.CENTER, font_name="Courier New")
        add_textbox(slide, PL + Cm(2.0), y + Cm(0.35), CW - Cm(2.3), Cm(1.1),
                    sol, font_size=18, color=C_WHITE, word_wrap=True)
        y += Cm(1.9)

    if footer:
        y += Cm(0.2)
        hbox(slide, PL, y, CW, Cm(1.8), accent=C_ORANGE,
             bg=RGBColor(0x20, 0x18, 0x08),
             text=f'"{footer}"', size=18)

    add_slide_num(slide, n)
    return slide


def make_teaching_moment(prs, n, case_label, title, cols3):
    """cols3 = list of (label, metric_text, detail, color)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    # Green-tinted teaching background (.tch class)
    set_bg(slide, RGBColor(0x0B, 0x2A, 0x1C))
    add_footer(slide)

    y = lbl(slide, f"Teaching Moment  ·  {case_label}", color=C_GREEN)
    y = rule(slide, y, color=C_GREEN)
    y = h2_text(slide, title, y, size=30)
    y -= Cm(0.3)

    col_w = (CW - Cm(1.0)) / 3
    ch = H - y - Cm(2.5)
    for i, (label, metric, detail, col) in enumerate(cols3):
        cx = PL + i * (col_w + Cm(0.5))
        add_rect(slide, cx, y, col_w, ch, fill_color=C_NAVY2)
        add_rect(slide, cx, y, col_w, Cm(0.18), fill_color=col)
        add_textbox(slide, cx + Cm(0.4), y + Cm(0.4), col_w - Cm(0.8), Cm(0.8),
                    label, font_size=18, bold=True, color=col,
                    font_name="Courier New")
        add_textbox(slide, cx + Cm(0.4), y + Cm(1.4), col_w - Cm(0.8), Cm(1.4),
                    metric, font_size=30, bold=True, color=col,
                    font_name="Courier New")
        add_textbox(slide, cx + Cm(0.4), y + Cm(3.1), col_w - Cm(0.8),
                    ch - Cm(3.4),
                    detail, font_size=18, color=C_DIM, word_wrap=True)

    add_slide_num(slide, n)
    return slide


def make_two_plan_compare(prs, n, slide_title, plan_data, footer=None):
    """plan_data = list of (title, bullets, color)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)

    y = lbl(slide, slide_title)
    y = rule(slide, y)

    col_w = (CW - Cm(2.0)) / 2
    ph = H - y - Cm(3.5)

    for i, (title, bullets, col) in enumerate(plan_data):
        cx = PL + i * (col_w + Cm(2.0))
        add_rect(slide, cx, y, col_w, ph, fill_color=C_NAVY2)
        add_rect(slide, cx, y, col_w, Cm(0.18), fill_color=col)
        add_textbox(slide, cx + Cm(0.4), y + Cm(0.4), col_w - Cm(0.8), Cm(1.2),
                    title, font_size=22, bold=True, color=col)
        by = y + Cm(1.8)
        for b in bullets:
            add_textbox(slide, cx + Cm(0.4), by, col_w - Cm(0.8), Cm(0.95),
                        b, font_size=17, color=C_WHITE,
                        font_name="Courier New")
            by += Cm(1.0)

    if footer:
        fy = y + ph + Cm(0.4)
        hbox(slide, PL, fy, CW, Cm(2.0), accent=C_ORANGE,
             bg=RGBColor(0x20, 0x18, 0x08),
             text=f'"{footer}"', size=18)

    add_slide_num(slide, n)
    return slide


# ═══════════════════════════════════════════════════════════════════
# BUILD DECK
# ═══════════════════════════════════════════════════════════════════

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

    # 7. Block B Transition
    make_block_transition(prs, 7,
        letter="B",
        title="Beams &\nTechnical Quality",
        block_num_str="01",
        block_label="BLOCK ONE",
        timer_text="12 MINUTES",
        bullets=[
            "Isocentre placement — position and tissue type",
            "Beam geometry, entry angles and arrangement",
            "MLC shaping, collimation and flash",
            "Wedges, field-in-field and dose shaping tools",
            "MU sanity check",
        ],
        bg_color=C_BLOCK_B,
        ghost_color=C_GHOST_B,
        label_color=C_LABEL_B)

    # 8. Field Arrangement Diagram
    make_field_arrangement(prs, 8)

    # 9. Technical Quality
    make_technical_quality(prs, 9)

    # 10. Hot Spots Diagram
    make_hotspots_diagram(prs, 10)

    # 11. Block C Transition
    make_block_transition(prs, 11,
        letter="C",
        title="Coverage",
        block_num_str="02",
        block_label="BLOCK TWO",
        timer_text="10 MINUTES",
        bullets=[
            "Volume hierarchy: GTV → CTV → ITV → PTV",
            "Reading the DVH — V-metrics and D-metrics",
            "D95, D2, homogeneity index (HI), conformity index (CI)",
            "Where cold spots live and when they matter",
        ],
        bg_color=C_BLOCK_C,
        ghost_color=C_GHOST_C,
        label_color=C_LABEL_C)

    # 12. DVH + Volume Hierarchy
    make_dvh_volume_hierarchy(prs, 12)

    # 13. Coverage Metrics
    make_coverage_metrics(prs, 13)

    # 14. Block O Transition
    make_block_transition(prs, 14,
        letter="O",
        title="Organs at Risk",
        block_num_str="03",
        block_label="BLOCK THREE",
        timer_text="10 MINUTES",
        bullets=[
            "Serial vs parallel organ architecture",
            "Matching the right metric: Dmax vs mean vs Vx",
            "QUANTEC constraints — the working reference",
            "PRV margins and tolerance negotiation",
        ],
        bg_color=C_BLOCK_O,
        ghost_color=C_GHOST_O,
        label_color=C_LABEL_O)

    # 15. Serial vs Parallel OARs
    make_serial_parallel(prs, 15)

    # 16. Block P·S Transition
    make_block_transition(prs, 16,
        letter="P·S",
        title="Prescription &\nSummation",
        block_num_str="04",
        block_label="BLOCK FOUR",
        timer_text="8 MINUTES",
        bullets=[
            "Verifying dose, fractionation and normalisation",
            "Laterality, intent and plan documentation",
            "Cumulative dose — or acknowledging none to sum",
            "Triage: critical, major, minor — and the gut-check",
        ],
        bg_color=C_BLOCK_PS,
        ghost_color=C_GHOST_PS,
        label_color=C_LABEL_PS)

    # 17. Tiers + Scorecard
    make_tiers_scorecard(prs, 17)

    # ── CASE 1: BREAST ──────────────────────────────────────────
    # 18. Case 1 Opener
    make_case_opener(prs, 18, 1,
        "Whole-Breast Tangents",
        "2.67 Gy × 15 = 40.05 Gy  ·  Whole breast, right side",
        "3DCRT tangential fields  ·  MLC shaped  ·  possible FiF or wedges",
        ["Heart", "Ipsilateral lung", "Contralateral breast"],
        timer="5 minutes — score Plan A in the workshop tool. Mark every FCB-CHOPS row before we discuss.")

    # 19. Breast Plan A Review
    make_case_review(prs, 19, "Case 1 · Breast · Plan A", [
        "Isocentre position — depth, tissue type",
        "Tangent field borders — crossing midline?",
        "Flash beyond breast tissue (≥2 cm?)",
        "Normalisation point location",
        "Hotspot magnitude and location (<115%?)",
        "Heart + lung on DVH",
    ])

    # 20. Breast Bad Plan Findings
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)
    col_w = (CW - Cm(2.0)) / 2
    y = lbl(slide, "Case 1 · Breast · Plan A — Findings")
    y = rule(slide, y)
    add_textbox(slide, PL, y, col_w, Cm(1.5),
                "Plan A is the bad plan", font_size=28, bold=True,
                color=RGBColor(0xF0, 0x6A, 0x6A))
    add_textbox(slide, PL + col_w + Cm(2.0), y, col_w, Cm(1.5),
                "What would have to change?", font_size=28, bold=True, color=C_WHITE)
    y += Cm(1.7)
    bad_finds = [
        ("CRITICAL", "Normalisation point in lung — every DVH number is unreliable",       C_RED),
        ("CRITICAL", "Hotspot >115% — unacceptable magnitude; likely outside PTV",          C_RED),
        ("CRITICAL", "Heart exceeding tolerance — cardiac dose above constraint",            C_RED),
        ("MAJOR",    "Isocentre too deep — in lung and heart",                               C_ORANGE),
        ("MAJOR",    "Fields crossing midline — unnecessary contralateral dose",             C_ORANGE),
        ("MAJOR",    "No flash applied — superficial breast under-coverage",                 C_ORANGE),
        ("MINOR",    "No wedges, no dose shaping — but secondary to the above",              C_YELLOW),
    ]
    fy = y
    for tag, text, col in bad_finds:
        fy = find_row(slide, PL, fy, col_w, tag, text, col)

    changes = [
        "Move isocentre to breast tissue, shallower depth",
        "Normalisation point into stable breast tissue",
        "Tangent borders must not cross midline",
        "Add ≥2 cm flash beyond breast tissue",
        "Add wedges or FiF to reduce hotspot <115%",
        "Re-check heart and lung on DVH after re-plan",
    ]
    rx = PL + col_w + Cm(2.0)
    ry = y
    for ch_text in changes:
        add_rect(slide, rx, ry, col_w, Cm(1.55), fill_color=C_NAVY2)
        add_rect(slide, rx, ry, Cm(0.18), Cm(1.55), fill_color=C_GREEN)
        add_textbox(slide, rx + Cm(0.5), ry + Cm(0.25), col_w - Cm(0.8), Cm(1.1),
                    "→  " + ch_text, font_size=17, color=C_WHITE, word_wrap=True)
        ry += Cm(1.7)

    hbox(slide, rx, ry, col_w, Cm(1.8),
         accent=C_RED, bg=RGBColor(0x28, 0x0E, 0x0E),
         text="This plan cannot be treated as drawn. The normalisation point failure alone makes every DVH number untrustworthy.",
         size=17)
    add_slide_num(slide, 20)

    # 21. Breast Acceptable vs Good
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_NAVY)
    add_footer(slide)
    col_w = (CW - Cm(2.0)) / 2
    y = lbl(slide, "Case 1 · Breast · Plans B & C — Acceptable → Good")
    y = rule(slide, y)

    plans_bc = [
        ("Plan B — Acceptable", "Wedges", [
            ("✓", "Good isocentre · norm point in breast · ≥2 cm flash · stays in midline", C_GREEN),
            ("✓", "Hotspot <115% · heart & lung within tolerance", C_GREEN),
            ("~", "HI < 0.15 with wedges — better than Plan A; FiF would improve further", C_YELLOW),
        ], C_ORANGE),
        ("Plan C — Good", "Field-in-Field", [
            ("✓", "Same coverage and OAR sparing as Plan B", C_GREEN),
            ("✓", "Hotspot <110% · HI <0.10 — homogeneous distribution", C_GREEN),
            ("✓", "FiF sub-segments selectively block emerging hot spots", C_GREEN),
        ], C_GREEN),
    ]
    ph = H - y - Cm(1.5)
    for i, (plan_title, plan_badge, finds, col) in enumerate(plans_bc):
        cx = PL + i * (col_w + Cm(2.0))
        add_rect(slide, cx, y, col_w, ph, fill_color=C_NAVY2)
        add_rect(slide, cx, y, col_w, Cm(0.18), fill_color=col)
        add_textbox(slide, cx + Cm(0.4), y + Cm(0.4), col_w - Cm(6.0), Cm(1.1),
                    plan_title, font_size=22, bold=True, color=col)
        add_rect(slide, cx + col_w - Cm(5.5), y + Cm(0.4), Cm(5.0), Cm(1.1),
                 fill_color=darken(col, 0.2))
        add_textbox(slide, cx + col_w - Cm(5.5), y + Cm(0.4), Cm(5.0), Cm(1.1),
                    plan_badge, font_size=18, bold=True, color=col,
                    font_name="Courier New", align=PP_ALIGN.CENTER)

        # Image placeholder
        add_rect(slide, cx + Cm(0.4), y + Cm(1.8), col_w - Cm(0.8), Cm(4.5),
                 fill_color=RGBColor(0x0C, 0x20, 0x35),
                 line_color=RGBColor(0x1A, 0x50, 0x70), line_width=1)
        add_textbox(slide, cx + Cm(0.4), y + Cm(3.3), col_w - Cm(0.8), Cm(1.0),
                    f"[ Plan {chr(66+i)} screenshot ]",
                    font_size=16, color=RGBColor(0x1A, 0x60, 0x80),
                    font_name="Courier New", align=PP_ALIGN.CENTER)

        fy = y + Cm(6.8)
        for tag, text, tcol in finds:
            fy = find_row(slide, cx + Cm(0.0), fy, col_w, tag, text, tcol)
    add_slide_num(slide, 21)

    # 22. Breast Teaching Moment
    make_teaching_moment(prs, 22, "Case 1",
        "Homogeneity — why it matters and how to achieve it",
        [
            ("BAD — No shaping",     "HI > 0.2",   "No wedges — dose piles up at beam entry. Hotspot is large and often exits the PTV. Normalisation in lung makes all numbers unreliable.", C_RED),
            ("ACCEPTABLE — Wedges",  "HI < 0.15",  "Wedge tilts the isodoses to compensate for the sloping breast surface. Hotspot <115%. Clinically deliverable and safe.", C_ORANGE),
            ("GOOD — Field-in-Field","HI < 0.10",  "FiF sub-segments are forward-planned to block the peak of each hotspot. Same coverage, same OAR sparing — just flatter dose distribution.", C_GREEN),
        ])

    # ── CASE 2: GYNAE ────────────────────────────────────────────
    # 23. Case 2 Opener
    make_case_opener(prs, 23, 2,
        "Whole-Pelvis Irradiation",
        "1.8 Gy × 25 = 45 Gy  ·  Whole pelvis · cervix + parametria + regional nodes",
        "3DCRT · 4-field box · AP/PA + 2 laterals · FiF applied",
        ["Rectum", "Small bowel", "Bladder", "Femoral heads"],
        timer="Score the plan in the workshop tool. What are your top three concerns?")

    # 24. Gynae Review
    make_case_review(prs, 24, "Case 2 · Gynae", [
        "Does the PTV receive adequate coverage?",
        "Which OARs appear on the DVH — are any lines elevated?",
        "Isodose lines — which normal structures are included?",
        "Rectum: where does the high-dose isodose sit?",
        "Small bowel: any loops visible in the high-dose region?",
        "Bladder + femoral heads within tolerance?",
    ])

    # 25. Gynae Findings
    make_case_findings(prs, 25, "Case 2 · Gynae", [
        ("✓ Coverage",    "PTV well covered — target dose maintained throughout",                                  C_GREEN),
        ("↑ Rectum",      "V45/V30 elevated — high-dose isodose wraps anterior rectal wall; above QUANTEC tolerance", C_RED),
        ("~ Small bowel", "AP/PA component brings anterior bowel into moderate-dose region. V45 elevated.",        C_ORANGE),
        ("✓ Bladder",     "Within acceptable limits",                                                              C_GREEN),
        ("✓ Fem heads",   "Within tolerance bilaterally",                                                          C_GREEN),
        ("NOTE",          "Brachytherapy boost to follow — cumulative rectal & bladder dose must be tracked",      C_CYAN),
        ("QUANTEC ref",   "Rectum: V70 <20%, V50 <50%  ·  45 Gy whole pelvis: lower-dose constraints apply",     C_DIM),
    ])

    # 26. Gynae Solutions
    make_case_solutions(prs, 26, "Case 2 · Gynae", [
        "Adjust AP/PA vs lateral field weighting to reduce posterior dose to the rectum",
        "Patient positioning & preparation — full bladder, belly board, small bowel displacement",
        "Reduce posterior PTV margin at the cervix-rectum interface where clearly documented",
        "Adjust prescription if clinically appropriate — requires senior discussion",
        "Accept with documentation — geometric constraint, not a planning error; record the trade-off explicitly",
    ], footer="EBRT is phase 1. Brachytherapy adds significant rectal and bladder dose. Track cumulative dose from the start.")

    # ── CASE 3: PROSTATE ─────────────────────────────────────────
    # 27. Case 3 Opener
    make_case_opener(prs, 27, 3,
        "Prostate — 3-Field Plan",
        "2.5 Gy × 27 = 67.5 Gy  ·  Prostate + proximal SVs",
        "AP + 2 posterior obliques  ·  full bladder  ·  empty rectum",
        ["Rectum", "Penile bulb", "Bladder", "Femoral heads", "Small bowel"],
        timer="Score the plan in the workshop tool. Which OARs concern you most?")

    # 28. Prostate Review
    make_case_review(prs, 28, "Case 3 · Prostate", [
        "PTV coverage — D95 and D98 on the DVH",
        "Rectum — where does the curve sit vs tolerance lines?",
        "Penile bulb — is it contoured? What is the mean dose?",
        "Bladder and femoral heads — within tolerance?",
        "Isodose: where does the posterior margin sit relative to the rectum?",
        "Is the isocentre in a stable tissue position?",
    ])

    # 29. Prostate Findings
    make_case_findings(prs, 29, "Case 3 · Prostate", [
        ("✓ Coverage",    "CTV D98 ≈ 99.2%  ·  PTV D98 ≈ 96.6%  ·  Coverage good",                C_GREEN),
        ("↑ Rectum",      "V55 ≈ 37.9%  ·  V59 ≈ 32.9% — over QUANTEC tolerance",                  C_RED),
        ("↑ Penile bulb", "Mean ≈ 62.8 Gy — above QUANTEC limit of 52 Gy (D90 <50 Gy)",            C_RED),
        ("~ Small bowel", "Minor focal involvement — within tolerance",                              C_YELLOW),
        ("✓ Bladder",     "Within tolerance",                                                        C_GREEN),
        ("✓ Fem heads",   "Within tolerance",                                                        C_GREEN),
        ("QUANTEC ref",   "Rectum: V70 <20%  V50 <50%  ·  Penile bulb: Dmean <52 Gy  D90 <50 Gy", C_DIM),
    ])

    # 30. Prostate Solutions
    make_case_solutions(prs, 30, "Case 3 · Prostate", [
        "Reduce inferior field margin — decrease penile bulb dose directly",
        "Reduce posterior PTV margin at the prostate–rectum interface",
        "Adjust posterior oblique field weighting to shift dose anteriorly",
        "Adjust prescription if clinically appropriate — senior discussion required",
        "Accept with documentation — geometric proximity, not a planning error; name it, don't ignore it",
    ], footer="Coverage was good — the problem was the OARs. Name it, don't ignore it.")

    # ── CASE 4: PITUITARY ────────────────────────────────────────
    # 31. Case 4 Opener
    make_case_opener(prs, 31, 4,
        "Pituitary Adenoma — Coplanar 3-Field Plan",
        "1.8 Gy × 30 = 54 Gy  ·  Pituitary fossa + margin",
        "3-field coplanar  ·  MLC shaped",
        ["Optic chiasm", "Optic nerves (bilateral)", "Brainstem", "Brain"],
        timer="Pay very close attention to the optic chiasm dose.")

    # 32. Pituitary Review
    make_case_review(prs, 32, "Case 4 · Pituitary", [
        "D95/D98/Dmin — is the PTV fully covered?",
        "Optic chiasm Dmax — QUANTEC hard limit is 54 Gy",
        "Optic nerves — bilateral Dmax",
        "Brainstem — Dmax and mean dose",
        "Inferior undercoverage — sphenoid sinus physics effect?",
        "Non-coplanar option — does it improve coverage or OAR sparing?",
    ])

    # 33. Pituitary Findings
    make_case_findings(prs, 33, "Case 4 · Pituitary", [
        ("✓ Coverage",     "PTV D95 adequate",                                                                     C_GREEN),
        ("↑ CHIASM",       "Dmax 54.9 Gy — ABOVE 54 Gy QUANTEC hard limit — mandatory senior sign-off",           C_RED),
        ("~ Inferior PTV", "PTV min 41.75 Gy — reduced inferior coverage",                                        C_YELLOW),
        ("✓ Brainstem",    "Dmax 52.5 Gy  ·  Mean ~11 Gy  ·  Within tolerance",                                  C_GREEN),
        ("NOTE",           "Sphenoid air sinus undercoverage 88.7% — physics-driven, expected",                    C_DIM),
        ("Architecture",   "Chiasm is serial: Dmax governs — even small breach = tolerance violation",             C_ORANGE),
    ])

    # 34. Pituitary Solutions
    make_two_plan_compare(prs, 34,
        "Case 4 · Pituitary — Solutions: Coplanar vs Non-Coplanar",
        [
            ("Plan A — Coplanar", [
                "Chiasm Dmax:      54.9 Gy  ↑  (above tolerance)",
                "PTV min:           41.75 Gy",
                "Brainstem mean:    ~11.0 Gy",
            ], C_ORANGE),
            ("Plan B — Non-Coplanar", [
                "Chiasm Dmax:      ~54.9 Gy  (SAME — cannot move chiasm)",
                "PTV min:           43.9 Gy  (+2.1 Gy improved)  ↑",
                "Brainstem mean:    9.7 Gy  (↓ 1.6 Gy improved)",
            ], C_GREEN),
        ],
        footer="Non-coplanar fields spread entrance angles — but cannot move the chiasm. Documentation required in both plans.")

    # ── CASE 5: H&N ──────────────────────────────────────────────
    # 35. Case 5 Opener
    make_case_opener(prs, 35, 5,
        "Head & Neck — Parotid Tumour",
        "2 Gy × 33 = 66 Gy  ·  Parotid tumour + regional nodes",
        "4-field 3DCRT  ·  photon + possible electron match",
        ["Contralateral parotid", "Spinal cord", "Brainstem", "Oral cavity"],
        timer="Focus especially on the contralateral parotid dose.")

    # 36. H&N Review
    make_case_review(prs, 36, "Case 5 · H&N", [
        "PTV coverage — D95, D98",
        "Cord Dmax — hard limit 45 Gy PRV / 50 Gy cord",
        "Brainstem Dmax",
        "Contralateral parotid mean — QUANTEC <26 Gy",
        "Ipsilateral parotid — likely sacrifice; document",
        "Beam entry/exit — does right lateral exit through contralateral parotid?",
    ])

    # 37. H&N Findings
    make_case_findings(prs, 37, "Case 5 · H&N", [
        ("✓ Coverage",        "PTV mean ≈ 66.0 Gy — well covered",                                                C_GREEN),
        ("↑ Contra parotid",  "Mean ≈ 3.30 Gy — higher than necessary (Plan A, 4-field)",                        C_ORANGE),
        ("✓ Cord",            "Dmax 19.9 Gy  ·  Within tolerance",                                                C_GREEN),
        ("✓ Brainstem",       "Dmax 18.2 Gy  ·  Within tolerance",                                                C_GREEN),
        ("~ Submandibular",   "R mean ≈ 40.4 Gy — elevated",                                                      C_YELLOW),
        ("NOTE",              "Electron/photon match junction — dose at junction must be summed",                  C_CYAN),
        ("Cause",             "Right lateral field exits through contralateral parotid — geometry problem",        C_DIM),
        ("Late effect",       "Xerostomia: most common serious late effect of H&N RT — always prioritise sparing", C_DIM),
    ])

    # 38. H&N Solutions
    make_two_plan_compare(prs, 38,
        "Case 5 · H&N — Solutions: 4-Field vs 3-Field",
        [
            ("Plan A — 4-Field", [
                "PTV mean:           ≈ 66.0 Gy",
                "Contra parotid:     3.30 Gy  ↑  (above optimal)",
                "Cord Dmax:          19.9 Gy",
                "Brainstem Dmax:     18.2 Gy",
            ], C_ORANGE),
            ("Plan B — 3-Field\n(right lateral removed)", [
                "PTV mean:           ≈ 66.0 Gy  (maintained)",
                "Contra parotid:     1.14 Gy  ↓  (↓ 2.2 Gy improvement)",
                "Cord Dmax:          17.8 Gy  (↓ 2.1 Gy)",
                "Brainstem Dmax:     16.1 Gy  (↓ 2.1 Gy)",
            ], C_GREEN),
        ],
        footer="Fewer fields ≠ worse plan. Remove the beam that exits through the contralateral parotid.")

    # ── WRAP-UP ──────────────────────────────────────────────────
    # 39. Wrap-up
    slide = prs.slides.add_slide(blank_layout(prs))
    set_bg(slide, C_DARK_NAVY)
    add_footer(slide)
    # Left accent bar
    add_rect(slide, 0, 0, Cm(0.4), H, fill_color=C_CYAN)

    y = lbl(slide, "Five Key Takeaways", top=Cm(1.0))
    y = rule(slide, y)

    takeaways = [
        ("1", "FCB-CHOPS every time",
         "Use the framework as a safety net — don't skip letters"),
        ("2", "DVH screens; slices diagnose",
         "Use both — DVH flags the problem, isodose slices locate it exactly"),
        ("3", "Match metric to organ architecture",
         "Serial → Dmax / D1cc   ·   Parallel → mean / Vx"),
        ("4", "Normalisation point matters",
         "In lung or air = unreliable; must be in stable representative tissue"),
        ("5", "Document every trade-off",
         "If you accept a compromise, write it down — name it, don't ignore it"),
    ]
    for num, headline, detail in takeaways:
        add_rect(slide, PL, y, CW, Cm(2.0), fill_color=C_NAVY)
        add_rect(slide, PL, y, Cm(0.18), Cm(2.0), fill_color=C_CYAN)
        add_textbox(slide, PL + Cm(0.5), y + Cm(0.2), Cm(1.5), Cm(1.6),
                    num, font_size=28, bold=True, color=RGBColor(0x0F, 0x9E, 0xD5),
                    align=PP_ALIGN.CENTER, font_name="Courier New")
        add_textbox(slide, PL + Cm(2.3), y + Cm(0.2), Cm(18), Cm(0.9),
                    headline, font_size=20, bold=True, color=C_WHITE)
        add_textbox(slide, PL + Cm(2.3), y + Cm(1.1), CW - Cm(2.6), Cm(0.8),
                    detail, font_size=17, color=C_DIM)
        y += Cm(2.15)

    add_textbox(slide, PL, H - Cm(2.2), CW, Cm(0.7),
                "References: FCB-CHOPS Weisman 2024 (PMID 40017913)  ·  ICRU 50/62/83  ·  QUANTEC 2010  ·  Emami 1991  ·  Khan textbook",
                font_size=12, color=C_DIM)
    add_rect(slide, PL, H - Cm(1.5), CW, Cm(1.1), fill_color=C_NAVY)
    add_textbox(slide, PL + Cm(0.3), H - Cm(1.45), CW - Cm(0.6), Cm(1.0),
                '"A good plan is a documented compromise"',
                font_size=20, italic=True, bold=True, color=C_ORANGE,
                align=PP_ALIGN.CENTER)

    add_slide_num(slide, 39)

    return prs


# ── Save ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    out = "/home/user/3DCRT-Evaluation-Workshop/3DCRT_Workshop_Deck_v2.pptx"
    prs = build_deck()
    prs.save(out)
    print(f"Saved → {out}")
    print(f"Slides: {len(prs.slides)}")
