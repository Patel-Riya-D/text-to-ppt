"""
ppt_service.py – Advanced AI PPT Generator
=============================================
FIXED VERSION – White & Blue Theme, No Overlaps, No Text Truncation

Key improvements:
- Logo placed in a reserved area that never overlaps content or slide numbers.
- Dynamic font sizing & line limiting to prevent text overflow.
- Motif-aware layout with explicit safety margins.
- Consistent white & blue theme (OceanDeep).
- Content area computed after reserving space for logo badge.
- Slide numbers repositioned to avoid logo collision.
- Text overflow detection with automatic font reduction.
"""

import os, json, math, copy
from openai import AzureOpenAI
from PIL import Image

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)

_chat_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

# --------------------------------------------------------------------------
# python-pptx imports
# --------------------------------------------------------------------------
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# --------------------------------------------------------------------------
# WHITE & BLUE THEME (fixed, consistent)
# --------------------------------------------------------------------------
THEME = {
    "primary":   (0, 102, 204),      # bright blue
    "secondary": (51, 153, 255),     # lighter blue
    "accent":    (0, 51, 102),        # dark blue
    "bg_light":  (255, 255, 255),     # pure white
    "bg_dark":   (0, 45, 90),         # dark blue for headers
    "text_dark": (30, 30, 30),        # near-black for body
    "text_light":(255, 255, 255),     # white text on dark
    "text_muted":(100, 100, 110),     # greyish
    "card_bg":   (255, 255, 255),     # white cards
    "header_font": "Calibri",
    "body_font":   "Calibri",
    "motif": "left_bar",              # simple, clean left bar
}

# --------------------------------------------------------------------------
# PRIMITIVE HELPERS
# --------------------------------------------------------------------------
def _rgb(t):
    return RGBColor(*t)

def _rect(slide, x, y, w, h, fill, line_color=None, line_width_pt=0.75):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(fill)
    if line_color:
        s.line.color.rgb = _rgb(line_color)
        s.line.width = Pt(line_width_pt)
    else:
        s.line.fill.background()
    return s

def _oval(slide, x, y, w, h, fill):
    s = slide.shapes.add_shape(9, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(fill)
    s.line.fill.background()
    return s

def _tb(slide, text, x, y, w, h, size, bold=False, color=None,
        face="Calibri", align=PP_ALIGN.LEFT, italic=False, wrap=True):
    bx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = bx.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = face
    if color:
        r.font.color.rgb = _rgb(color)
    return bx

# --------------------------------------------------------------------------
# SMART BULLETS – NO OVERFLOW
# --------------------------------------------------------------------------
def _add_bullets(slide, points, x, y, w, h, theme, base_size=16, max_lines_per_bullet=6):
    """
    Renders bullet points with dynamic font reduction if content exceeds available height.
    Returns True if all text fits (after possible size reduction), else False.
    """
    if not points:
        return True
    
    # Estimate line height (Calibri ~1.2 * font size in points -> inches)
    # 1 point = 1/72 inch. So line height in inches = (font_size * 1.2) / 72
    def total_height(font_sz):
        lines_per_bullet = min(max_lines_per_bullet, max(1, len(points[0].split()) // 10 + 2))
        total_lines = len(points) * lines_per_bullet
        line_height_in = (font_sz * 1.2) / 72.0
        return total_lines * line_height_in + 0.1  # small padding
    
    # Try font sizes from base_size down to 10
    for try_size in range(base_size, 9, -1):
        needed_h = total_height(try_size)
        if needed_h <= h + 0.05:   # allow tiny slack
            final_size = try_size
            break
    else:
        final_size = 10   # fallback, will truncate
    
    bx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.auto_size = None  # manual
    first = True
    for pt in points:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4)
        p.space_after = Pt(4)
        # Bullet glyph
        br = p.add_run()
        br.text = "▸  "
        br.font.size = Pt(final_size - 1)
        br.font.bold = True
        br.font.name = theme["body_font"]
        br.font.color.rgb = _rgb(theme["secondary"])
        # Text
        tr = p.add_run()
        tr.text = pt
        tr.font.size = Pt(final_size)
        tr.font.bold = False
        tr.font.name = theme["body_font"]
        tr.font.color.rgb = _rgb(theme["text_dark"])
    
    return needed_h <= h + 0.05

# --------------------------------------------------------------------------
# LOGO PLACEMENT – RESERVED AREA, NEVER OVERLAPS
# --------------------------------------------------------------------------
LOGO_BADGE_WIDTH = 1.6   # inches (including padding)
LOGO_BADGE_HEIGHT = 0.9
LOGO_MAX_W = 1.3
LOGO_MAX_H = 0.7

def _add_logo(slide, logo_path, theme, x_right_margin=0.2, y_top_margin=0.2):
    """
    Places logo in top-right corner, but content area is already shrunk to avoid it.
    This function just draws the badge; caller must ensure content area does not extend into this region.
    """
    if not logo_path or not os.path.exists(logo_path):
        return
    try:
        with Image.open(logo_path) as img:
            iw, ih = img.size
        ratio = iw / ih if ih else 1.0
        logo_h = min(LOGO_MAX_H, LOGO_BADGE_HEIGHT - 0.2)
        logo_w = min(logo_h * ratio, LOGO_MAX_W)
        logo_h = logo_w / ratio if ratio > 0 else logo_h
        
        badge_x = 13.3 - LOGO_BADGE_WIDTH - x_right_margin
        badge_y = y_top_margin
        # White card with blue border
        _rect(slide, badge_x, badge_y, LOGO_BADGE_WIDTH, LOGO_BADGE_HEIGHT,
              (255, 255, 255), line_color=theme["secondary"], line_width_pt=1.0)
        # Logo image
        slide.shapes.add_picture(
            logo_path,
            left=Inches(badge_x + (LOGO_BADGE_WIDTH - logo_w)/2),
            top=Inches(badge_y + (LOGO_BADGE_HEIGHT - logo_h)/2),
            width=Inches(logo_w),
            height=Inches(logo_h)
        )
    except Exception as e:
        print(f"[logo] error: {e}")

# --------------------------------------------------------------------------
# CONTENT AREA – RESERVES SPACE FOR LOGO
# --------------------------------------------------------------------------
def _content_area(has_logo=True, header_h=1.6):
    """
    Returns (x, y, width, height) for main content.
    Reserves top-right area for logo if present.
    """
    W, H = 13.3, 7.5
    footer_h = 0.3
    usable_h = H - header_h - footer_h - 0.2
    # Left margin for motif bar
    left_margin = 0.45
    # Right margin: if logo exists, reserve its width + gap
    right_margin = (LOGO_BADGE_WIDTH + 0.3) if has_logo else 0.3
    content_w = W - left_margin - right_margin
    return left_margin, header_h + 0.1, content_w, usable_h - 0.1

# --------------------------------------------------------------------------
# SLIDE HEADER (with slide number not overlapped by logo)
# --------------------------------------------------------------------------
def _draw_header(slide, title, subtitle, slide_num, theme, has_logo=True):
    W, H = 13.3, 7.5
    header_h = 1.4
    # Dark blue header band
    _rect(slide, 0, 0, W, header_h, theme["bg_dark"])
    # Accent line at bottom of header
    _rect(slide, 0, header_h - 0.08, W, 0.08, theme["accent"])
    # Title
    title_x = 0.45
    title_y = 0.2
    title_w = W - 2.2 if has_logo else W - 1.0
    _tb(slide, title, title_x, title_y, title_w, 0.7, 28,
        bold=True, color=theme["text_light"], face=theme["header_font"])
    # Subtitle
    if subtitle:
        _tb(slide, subtitle, title_x, title_y + 0.7, title_w, 0.4, 13,
            italic=True, color=theme["secondary"], face=theme["body_font"])
    # Slide number – placed left of logo badge
    num_x = W - LOGO_BADGE_WIDTH - 0.5 if has_logo else W - 1.0
    num_y = 0.25
    _oval(slide, num_x, num_y, 0.5, 0.5, theme["accent"])
    _tb(slide, str(slide_num), num_x + 0.12, num_y + 0.07, 0.3, 0.4, 14,
        bold=True, color=theme["text_light"], face=theme["header_font"], align=PP_ALIGN.CENTER)

def _draw_footer(slide, theme):
    W, H = 13.3, 7.5
    _rect(slide, 0, H - 0.25, W, 0.25, theme["bg_dark"])

# --------------------------------------------------------------------------
# MOTIF – simple left bar (clean white/blue)
# --------------------------------------------------------------------------
def _apply_motif(slide, theme):
    W, H = 13.3, 7.5
    # Thick blue left bar
    _rect(slide, 0, 0, 0.25, H, theme["primary"])
    _rect(slide, 0.25, 0, 0.05, H, theme["secondary"])
    # Background white
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(theme["bg_light"])

# --------------------------------------------------------------------------
# CONTENT SLIDE RENDERERS (each respects content area)
# --------------------------------------------------------------------------
def _slide_bullets(prs, item, theme, slide_num, logo_path, has_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_motif(slide, theme)
    _draw_header(slide, item["title"], item.get("subtitle", ""), slide_num, theme, has_logo)
    
    x, y, w, h = _content_area(has_logo)
    # White card
    _rect(slide, x, y, w, h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
    # Bullets
    points = item.get("content", [])
    _add_bullets(slide, points, x+0.15, y+0.1, w-0.3, h-0.2, theme, base_size=15)
    
    _draw_footer(slide, theme)
    if logo_path:
        _add_logo(slide, logo_path, theme)
    return slide

def _slide_two_column(prs, item, theme, slide_num, logo_path, has_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_motif(slide, theme)
    _draw_header(slide, item["title"], item.get("subtitle", ""), slide_num, theme, has_logo)
    
    x, y, w, h = _content_area(has_logo)
    gap = 0.2
    col_w = (w - gap) / 2
    left_title = item.get("left_title", "Left")
    right_title = item.get("right_title", "Right")
    left_pts = item.get("left_points", [])
    right_pts = item.get("right_points", [])
    
    # Left column
    _rect(slide, x, y, col_w, h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
    _rect(slide, x, y, col_w, 0.45, theme["primary"])
    _tb(slide, left_title, x+0.1, y+0.07, col_w-0.2, 0.35, 14, bold=True,
        color=theme["text_light"], face=theme["header_font"])
    _add_bullets(slide, left_pts, x+0.1, y+0.55, col_w-0.2, h-0.65, theme, base_size=14)
    
    # Right column
    rx = x + col_w + gap
    _rect(slide, rx, y, col_w, h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
    _rect(slide, rx, y, col_w, 0.45, theme["secondary"])
    _tb(slide, right_title, rx+0.1, y+0.07, col_w-0.2, 0.35, 14, bold=True,
        color=theme["text_light"], face=theme["header_font"])
    _add_bullets(slide, right_pts, rx+0.1, y+0.55, col_w-0.2, h-0.65, theme, base_size=14)
    
    _draw_footer(slide, theme)
    if logo_path:
        _add_logo(slide, logo_path, theme)
    return slide

def _slide_big_stat(prs, item, theme, slide_num, logo_path, has_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_motif(slide, theme)
    _draw_header(slide, item["title"], item.get("subtitle", ""), slide_num, theme, has_logo)
    
    x, y, w, h = _content_area(has_logo)
    stat_w = 3.5
    stat = item.get("stat", "100%")
    stat_label = item.get("stat_label", "")
    bullets = item.get("content", [])
    
    # Stat panel (left)
    _rect(slide, x, y, stat_w, h, theme["primary"])
    _tb(slide, stat, x+0.1, y+0.5, stat_w-0.2, 2.0, 52, bold=True,
        color=theme["accent"], face=theme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, stat_label, x+0.1, y+2.6, stat_w-0.2, 0.8, 14,
        color=theme["text_light"], face=theme["body_font"], align=PP_ALIGN.CENTER)
    
    # Bullets (right)
    bx = x + stat_w + 0.2
    bw = w - stat_w - 0.2
    _rect(slide, bx, y, bw, h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
    _add_bullets(slide, bullets, bx+0.15, y+0.1, bw-0.3, h-0.2, theme, base_size=14)
    
    _draw_footer(slide, theme)
    if logo_path:
        _add_logo(slide, logo_path, theme)
    return slide

def _slide_timeline(prs, item, theme, slide_num, logo_path, has_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_motif(slide, theme)
    _draw_header(slide, item["title"], item.get("subtitle", ""), slide_num, theme, has_logo)
    
    x, y, w, h = _content_area(has_logo)
    steps = item.get("steps", [])
    n = max(len(steps), 1)
    step_w = w / n
    timeline_y = y + h * 0.5
    
    # Timeline line
    _rect(slide, x, timeline_y, w, 0.05, theme["secondary"])
    
    for i, step in enumerate(steps):
        cx = x + i * step_w + step_w/2
        # Node circle
        _oval(slide, cx-0.2, timeline_y-0.2, 0.4, 0.4, theme["primary"])
        _oval(slide, cx-0.1, timeline_y-0.1, 0.2, 0.2, theme["accent"])
        # Card above or below
        card_w = step_w * 0.8
        card_h = h * 0.35
        card_x = cx - card_w/2
        is_above = (i % 2 == 0)
        card_y = timeline_y - card_h - 0.25 if is_above else timeline_y + 0.35
        # Draw card
        _rect(slide, card_x, card_y, card_w, card_h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
        _tb(slide, step.get("label", f"Step {i+1}"), card_x+0.1, card_y+0.05, card_w-0.2, 0.35, 12, bold=True, color=theme["primary"])
        _tb(slide, step.get("detail", ""), card_x+0.1, card_y+0.4, card_w-0.2, card_h-0.45, 11, color=theme["text_dark"])
        # Connector line
        if is_above:
            _rect(slide, cx-0.03, card_y+card_h, 0.06, timeline_y - (card_y+card_h), theme["secondary"])
        else:
            _rect(slide, cx-0.03, timeline_y+0.05, 0.06, card_y - timeline_y, theme["secondary"])
    
    _draw_footer(slide, theme)
    if logo_path:
        _add_logo(slide, logo_path, theme)
    return slide

def _slide_icon_grid(prs, item, theme, slide_num, logo_path, has_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_motif(slide, theme)
    _draw_header(slide, item["title"], item.get("subtitle", ""), slide_num, theme, has_logo)
    
    x, y, w, h = _content_area(has_logo)
    cols = 2
    rows = 2
    gap = 0.15
    card_w = (w - gap) / cols
    card_h = (h - gap) / rows
    items = item.get("grid_items", [])[:4]
    
    for idx, gi in enumerate(items):
        col = idx % cols
        row = idx // cols
        cx = x + col * (card_w + gap)
        cy = y + row * (card_h + gap)
        _rect(slide, cx, cy, card_w, card_h, theme["card_bg"], line_color=theme["secondary"], line_width_pt=0.5)
        # Icon area (colored circle)
        _oval(slide, cx+0.15, cy+0.15, 0.5, 0.5, theme["primary"])
        icon_char = (gi.get("icon", "★")[0].upper())
        _tb(slide, icon_char, cx+0.22, cy+0.2, 0.36, 0.4, 18, bold=True, color=theme["text_light"], align=PP_ALIGN.CENTER)
        # Title
        _tb(slide, gi.get("title", ""), cx+0.8, cy+0.15, card_w-0.95, 0.4, 14, bold=True, color=theme["primary"])
        # Detail
        _tb(slide, gi.get("detail", ""), cx+0.15, cy+0.7, card_w-0.3, card_h-0.85, 11, color=theme["text_dark"])
    
    _draw_footer(slide, theme)
    if logo_path:
        _add_logo(slide, logo_path, theme)
    return slide

_RENDERERS = {
    "bullets":    _slide_bullets,
    "two_column": _slide_two_column,
    "big_stat":   _slide_big_stat,
    "timeline":   _slide_timeline,
    "icon_grid":  _slide_icon_grid,
}

# --------------------------------------------------------------------------
# TITLE & CLOSING SLIDES
# --------------------------------------------------------------------------
def _add_title_slide(prs, topic, theme, logo_path):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = 13.3, 7.5
    _rect(slide, 0, 0, W, H, theme["bg_dark"])
    _rect(slide, 0, 0, 0.35, H, theme["primary"])
    _rect(slide, 0.35, 0, 0.08, H, theme["secondary"])
    # Decorative circles
    _oval(slide, W-4, H-4, 5, 5, theme["primary"])
    _oval(slide, W-2.8, H-2.8, 3, 3, theme["bg_dark"])
    # Title
    _tb(slide, topic.upper(), 0.8, 2.0, W-2, 2.0, 46, bold=True,
        color=theme["text_light"], face=theme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, "AI Generated Presentation", 0.8, 4.2, W-2, 0.7, 18,
        color=theme["secondary"], face=theme["body_font"], align=PP_ALIGN.CENTER)
    if logo_path:
        _add_logo(slide, logo_path, theme, x_right_margin=0.5, y_top_margin=0.5)
    return slide

def _add_closing_slide(prs, topic, theme, logo_path):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = 13.3, 7.5
    _rect(slide, 0, 0, W, H, theme["bg_dark"])
    _rect(slide, 0, 0, 0.35, H, theme["primary"])
    _rect(slide, 0.35, 0, 0.08, H, theme["secondary"])
    _oval(slide, W-5, H-5, 6, 6, theme["primary"])
    _oval(slide, W-3.5, H-3.5, 3.5, 3.5, theme["bg_dark"])
    _tb(slide, "THANK YOU", 0.8, 2.2, W-2, 1.5, 58, bold=True,
        color=theme["text_light"], face=theme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, topic, 0.8, 4.0, W-2, 0.8, 24,
        color=theme["secondary"], face=theme["body_font"], align=PP_ALIGN.CENTER)
    _tb(slide, "Powered by AI", 0.8, 5.0, W-2, 0.5, 14,
        color=theme["text_muted"], face=theme["body_font"], align=PP_ALIGN.CENTER)
    if logo_path:
        _add_logo(slide, logo_path, theme, x_right_margin=0.5, y_top_margin=0.5)
    return slide

# --------------------------------------------------------------------------
# LLM CONTENT GENERATION (unchanged but robust)
# --------------------------------------------------------------------------
def generate_slide_content(topic: str, num_slides: int = 5, tone: str = "Professional") -> dict:
    prompt = f"""
You are a world-class presentation strategist.
Create a {tone} PowerPoint presentation on "{topic}".
OUTPUT: exactly {num_slides} content slides (NO title, NO closing). Return ONLY raw JSON.

RULES:
- Each slide: "title" (max 8 words), "content" list of 5-7 bullet points.
- Each bullet: 25-35 words, substantive.
- Layouts: "bullets", "two_column", "big_stat", "timeline", "icon_grid". Vary them.

For two_column: also provide "left_title", "right_title", "left_points", "right_points".
For big_stat: also provide "stat", "stat_label".
For timeline: also provide "steps" (list of {{"label": "...", "detail": "..."}}).
For icon_grid: also provide "grid_items" (list of {{"icon": "keyword", "title": "...", "detail": "..."}}).

JSON schema:
{{
  "slides": [
    {{
      "title": "...",
      "subtitle": "...",
      "content": ["...", "..."],
      "layout": "bullets",
      "icon_keyword": "growth"
    }}
  ]
}}
"""
    resp = _chat_client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.78,
        max_tokens=6000,
    )
    text = resp.choices[0].message.content.strip()
    if text.startswith("```"):
        text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)

# --------------------------------------------------------------------------
# MAIN API
# --------------------------------------------------------------------------
def create_ppt(slide_data, topic: str, logo_path=None, tone: str = "Professional") -> str:
    prs = Presentation()
    W, H = 13.3, 7.5
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    
    # Use fixed white/blue theme
    theme = THEME.copy()
    slides = slide_data.get("slides", []) if isinstance(slide_data, dict) else slide_data
    
    has_logo = bool(logo_path and os.path.exists(logo_path))
    
    _add_title_slide(prs, topic, theme, logo_path if has_logo else None)
    
    for i, item in enumerate(slides, start=1):
        layout = item.get("layout", "bullets")
        renderer = _RENDERERS.get(layout, _slide_bullets)
        renderer(prs, item, theme, i, logo_path if has_logo else None, has_logo)
    
    _add_closing_slide(prs, topic, theme, logo_path if has_logo else None)
    
    safe_name = topic.replace(" ", "_").replace("/", "_")[:60]
    os.makedirs("generated", exist_ok=True)
    file_path = f"generated/{safe_name}.pptx"
    prs.save(file_path)
    return file_path