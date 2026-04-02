"""
ppt_service.py  –  AI PPT Generator (Production-Ready, All Issues Fixed)
=========================================================================

ROOT CAUSES FIXED vs. PREVIOUS VERSION
---------------------------------------
1.  {topic} literal placeholder  → topic variable passed correctly everywhere.
2.  Intro title overflow         → title capped at left-panel width (6.4").
3.  two_column empty columns     → prompt REQUIRES left_points/right_points;
                                   validator splits content[] as fallback when missing.
4.  big_stat shows "100%"        → prompt demands a real numeric stat;
                                   validator rejects "100%" placeholder.
5.  timeline "Step 1/No details" → prompt requires ≥4 steps with real label+detail;
                                   validator fills fallback steps if LLM omits them.
6.  Blank bullet slides          → content[] validated non-empty before render.
7.  Slide number badge overflow  → badge oval = 0.60", font auto-sized for 2-digit nums.
8.  Summary "999" badge          → replaced with "★" star glyph.
9.  Closing bullet misalignment  → all action items in ONE textbox (not three separate).
10. "Questions?" invisible       → color changed to near-white (200,225,255) on dark bg.
11. case_study missing company   → prompt requires company field; validator enforces it.
12. case_study empty middle      → result ribbon + metric bars rendered before bullets.
13. Large empty whitespace       → content boxes sized to fill full available height.
14. Two-column col header colors → left=primary, right=accent2 for visual differentiation.
15. Footer missing on slides     → _footer() called on every content slide.
"""

import os
import json
import re
from openai import AzureOpenAI
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

# --------------------------------------------------------------------------
# Config  (import from your existing config module)
# --------------------------------------------------------------------------
from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

# --------------------------------------------------------------------------
# Slide dimensions
# --------------------------------------------------------------------------
W, H = 13.3, 7.5
HEADER_H  = 1.40
FOOTER_H  = 0.25
CONTENT_Y = HEADER_H + 0.12
CONTENT_H = H - HEADER_H - FOOTER_H - 0.22

# Logo badge geometry
BADGE_W, BADGE_H = 1.65, 0.90
BADGE_X = W - BADGE_W - 0.18
BADGE_Y = 0.18

# --------------------------------------------------------------------------
# Base theme  (all keys guaranteed present)
# --------------------------------------------------------------------------
BASE_THEME = dict(
    primary    = (0,  102, 204),
    secondary  = (51, 153, 255),
    accent     = (0,   82, 164),
    accent2    = (0,  150, 200),
    bg_light   = (255, 255, 255),
    bg_dark    = (0,  102, 204),
    text_dark  = (30,  30,  30),
    text_light = (255, 255, 255),
    text_muted = (120, 140, 160),
    card_bg    = (255, 255, 255),
    header_font= "Calibri",
    body_font  = "Calibri",
)

# ==========================================================================
# LOW-LEVEL DRAWING HELPERS
# ==========================================================================

def _c(t):
    return RGBColor(*t)

def _IN(v):
    return Inches(v)


def _coerce_int(value, default=50, min_value=0, max_value=100):
    """Best-effort integer coercion for LLM output like '70%' or '+42'."""
    if isinstance(value, int):
        n = value
    else:
        m = re.search(r"-?\d+", str(value))
        if not m:
            return default
        n = int(m.group(0))
    return max(min_value, min(max_value, n))


SAFE_ICON_SET = {
    "▸", "•", "◆", "◼", "◻", "▲", "▼", "▶", "→", "✓", "✔", "★", "☆", "+", "-"
}


def _safe_icon(value, default="▸"):
    """Return a PowerPoint-safe bullet/icon glyph."""
    s = str(value or "").strip()
    if not s:
        return default
    if s in SAFE_ICON_SET:
        return s
    # Keep simple single ASCII symbols/letters; reject emoji/multi-codepoint icons.
    if len(s) == 1 and s.isascii() and (s.isalnum() or s in "!@#$%^&*()_+=-<>?/|~"):
        return s
    return default


def _mix(c1, c2, t):
    t = max(0.0, min(1.0, float(t)))
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def _topic_seed(topic: str) -> int:
    return sum((i + 1) * ord(ch) for i, ch in enumerate(str(topic))) % 9973


def _title_size(text: str, base: int = 30) -> int:
    n = len(str(text or "").strip())
    if n <= 42:
        return base
    if n <= 65:
        return max(26, base - 2)
    if n <= 90:
        return max(23, base - 5)
    return max(20, base - 8)


def _subtitle_size(text: str, base: int = 13) -> int:
    n = len(str(text or "").strip())
    if n <= 60:
        return base
    if n <= 95:
        return max(11, base - 1)
    return max(10, base - 2)


def _opening_slide_spec(topic: str, tone: str, variant: int) -> dict:
    styles = [
        {"pattern_name": "opening_compact", "surface": "light", "header_variant": "split", "card_variant": "banded", "footer_variant": "line", "badge_shape": "rect", "accent_rotation": "auto"},
        {"pattern_name": "opening_editorial", "surface": "light", "header_variant": "banded", "card_variant": "outline", "footer_variant": "line", "badge_shape": "oval", "accent_rotation": "static"},
        {"pattern_name": "opening_executive", "surface": "light", "header_variant": "solid", "card_variant": "soft", "footer_variant": "solid", "badge_shape": "rect", "accent_rotation": "auto"},
    ]

    if variant == 0:
        return {
            "title": topic,
            "subtitle": f"{tone} Presentation Overview",
            "layout": "big_stat",
            "icon": "★",
            "stat": "2026",
            "stat_label": "Strategic Outlook",
            "stat_source": "Current industry snapshot",
            "content": [
                f"**Agenda**: this deck explains the core foundations, architecture choices, and implementation priorities for {topic}.",
                "**What to decide**: we focus on practical trade-offs, measurable outcomes, and execution sequencing.",
                "**Business value**: each section links design choices to cost, speed, resilience, and long-term scalability.",
                "**Evidence**: examples and benchmarks are included to support stakeholder discussions and planning.",
                "**Expected outcome**: a decision-ready roadmap aligned to your audience and delivery context.",
            ],
            "style": styles[variant],
        }
    if variant == 1:
        return {
            "title": topic,
            "subtitle": "Session Introduction",
            "layout": "two_column",
            "icon": "▸",
            "left_title": "What We Will Cover",
            "right_title": "Why It Matters",
            "left_points": [
                "**Context**: current state, constraints, and key drivers shaping this topic.",
                "**Design options**: practical patterns with strengths, limits, and usage guidance.",
                "**Execution**: implementation steps, ownership model, and milestone sequencing.",
                "**Measurement**: KPI framework to track impact and iterate confidently.",
            ],
            "right_points": [
                "**Clarity**: aligns teams on vocabulary, scope, and architectural intent.",
                "**Speed**: reduces trial-and-error by using proven references and decisions.",
                "**Risk control**: anticipates reliability, security, and scaling bottlenecks early.",
                "**Outcome focus**: connects technical decisions to business performance metrics.",
            ],
            "content": [],
            "style": styles[variant],
        }
    return {
        "title": topic,
        "subtitle": "Quick Start Overview",
        "layout": "bullets",
        "icon": "✓",
        "content": [
            f"**Purpose**: establish a clear, shared understanding of {topic} and its most important design decisions.",
            "**Scope**: cover core building blocks, practical adoption paths, and governance essentials.",
            "**Evidence-led**: use real benchmarks and examples to guide implementation choices.",
            "**Execution-ready**: translate insights into prioritized actions with accountability and timeline.",
            "**Audience-ready**: balance strategic context and technical depth for mixed stakeholders.",
        ],
        "style": styles[variant],
    }


def _closing_slide_spec(topic: str, variant: int) -> dict:
    styles = [
        {"pattern_name": "closing_gratitude", "surface": "light", "header_variant": "banded", "card_variant": "soft", "footer_variant": "line", "badge_shape": "rect", "accent_rotation": "static"},
        {"pattern_name": "closing_discussion", "surface": "light", "header_variant": "split", "card_variant": "outline", "footer_variant": "line", "badge_shape": "oval", "accent_rotation": "auto"},
        {"pattern_name": "closing_next", "surface": "light", "header_variant": "solid", "card_variant": "banded", "footer_variant": "solid", "badge_shape": "rect", "accent_rotation": "static"},
    ]
    titles = ["Thank You", "Thank You & Q&A", "Thank You"]
    subtitles = ["Questions & Discussion", "Open Discussion", "Final Questions"]
    return {
        "title": titles[variant],
        "subtitle": subtitles[variant],
        "layout": "bullets",
        "icon": "✓",
        "content": [
            f"**Thank you** for your time and engagement on {topic}.",
            "**Questions welcome**: we can dive deeper into architecture choices, implementation risks, or rollout priorities.",
            "**Action alignment**: confirm owner, timeline, and KPI for the first execution milestone.",
            "**Follow-up**: document decisions and circulate the agreed next-step plan across stakeholders.",
            "**Collaboration**: share feedback so we can refine scope and improve implementation outcomes.",
        ],
        "style": styles[variant],
    }


def _parse_color(value, fallback):
    """Parse color from [r,g,b], '#RRGGBB', or 'r,g,b'."""
    if isinstance(value, (list, tuple)) and len(value) == 3:
        try:
            return tuple(max(0, min(255, int(v))) for v in value)
        except Exception:
            return fallback
    if isinstance(value, str):
        s = value.strip()
        if re.fullmatch(r"#?[0-9A-Fa-f]{6}", s):
            s = s.lstrip("#")
            return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))
        parts = [p.strip() for p in s.split(",")]
        if len(parts) == 3:
            try:
                return tuple(max(0, min(255, int(v))) for v in parts)
            except Exception:
                return fallback
    return fallback


def _normalize_theme(theme_payload):
    """Merge LLM-provided theme into BASE_THEME with safe fallbacks."""
    theme = BASE_THEME.copy()
    if not isinstance(theme_payload, dict):
        return theme

    for key in ("primary", "secondary", "accent", "accent2",
                "bg_light", "bg_dark", "text_dark", "text_light",
                "text_muted", "card_bg"):
        if key in theme_payload:
            theme[key] = _parse_color(theme_payload.get(key), theme[key])

    for key in ("header_font", "body_font"):
        v = theme_payload.get(key)
        if isinstance(v, str) and v.strip():
            theme[key] = v.strip()

    return theme


def _slide_theme(base_theme, style=None, slide_idx=1):
    """Resolve per-slide theme variants from LLM style directives."""
    style = style if isinstance(style, dict) else {}
    theme = dict(base_theme)

    # Optional per-slide palette overrides.
    for key in ("primary", "secondary", "accent", "accent2",
                "bg_light", "bg_dark", "text_dark", "text_light",
                "text_muted", "card_bg"):
        if key in style:
            theme[key] = _parse_color(style.get(key), theme[key])

    # Rotate accents when requested for variety between slides.
    if str(style.get("accent_rotation", "")).lower() in ("auto", "rotate"):
        order = [theme["primary"], theme["secondary"], theme["accent2"], theme["accent"]]
        shift = (slide_idx - 1) % len(order)
        rot = order[shift:] + order[:shift]
        theme["primary"], theme["secondary"], theme["accent2"], theme["accent"] = rot[0], rot[1], rot[2], rot[3]

    surface = str(style.get("surface", "light")).lower()
    if surface == "dark":
        theme["bg_light"] = theme["bg_dark"]
        theme["card_bg"] = _mix(theme["bg_dark"], (255, 255, 255), 0.10)
        theme["text_dark"] = theme["text_light"]
        theme["text_muted"] = _mix(theme["text_light"], theme["bg_dark"], 0.45)
    elif surface == "tint":
        theme["bg_light"] = _mix(theme["bg_light"], theme["primary"], 0.10)
        theme["card_bg"] = _mix(theme["card_bg"], theme["secondary"], 0.05)

    return theme


def _rect(slide, x, y, w, h, fill, line=None, lw=0.75):
    s = slide.shapes.add_shape(1, _IN(x), _IN(y), _IN(w), _IN(h))
    s.fill.solid()
    s.fill.fore_color.rgb = _c(fill)
    if line:
        s.line.color.rgb = _c(line)
        s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s


def _oval(slide, x, y, w, h, fill):
    s = slide.shapes.add_shape(9, _IN(x), _IN(y), _IN(w), _IN(h))
    s.fill.solid()
    s.fill.fore_color.rgb = _c(fill)
    s.line.fill.background()
    return s


def _tb(slide, text, x, y, w, h, size,
        bold=False, italic=False, color=None,
        face="Calibri", align=PP_ALIGN.LEFT, shrink_to_fit=False):
    """Single-paragraph textbox with optional **bold** marker support."""
    text = str(text) if not isinstance(text, str) else text
    bx = slide.shapes.add_textbox(_IN(x), _IN(y), _IN(w), _IN(h))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    if shrink_to_fit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.alignment = align
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        r = p.add_run()
        r.text = part
        r.font.size   = Pt(size)
        r.font.bold   = bold or (i % 2 == 1)
        r.font.italic = italic
        r.font.name   = face
        if color:
            r.font.color.rgb = _c(color)
    return bx


def _bullets(slide, points, x, y, w, h,
             size=15, icon="▸", icon_color=None, text_color=None,
             face="Calibri", max_pts=5):
    """
    Render a bulleted list in a single textbox.
    Supports **bold** markers inside bullet text.
    """
    if not points:
        return
    ic = icon_color or BASE_THEME["secondary"]
    tc = text_color  or BASE_THEME["text_dark"]
    icon = _safe_icon(icon, default="▸")
    pts = [str(p) for p in points[:max_pts]]
    count = len(pts)

    bx = slide.shapes.add_textbox(_IN(x), _IN(y), _IN(w), _IN(h))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if count <= 3 else MSO_ANCHOR.TOP

    # Make sparse slides feel intentional, not empty.
    if count <= 2:
        size = min(size + 5, 24)
    elif count == 3:
        size = min(size + 3, 22)
    elif count == 4:
        size = min(size + 1, 19)

    first = True
    for pt in pts:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if count <= 3:
            p.space_before = Pt(11)
            p.space_after  = Pt(11)
        else:
            p.space_before = Pt(6)
            p.space_after  = Pt(6)

        # Bullet glyph
        ir = p.add_run()
        ir.text = f"{icon}  "
        ir.font.size  = Pt(size - 1)
        ir.font.bold  = True
        ir.font.name  = face
        ir.font.color.rgb = _c(ic)

        # Text with **bold** support
        for i, part in enumerate(pt.split("**")):
            if not part:
                continue
            r = p.add_run()
            r.text      = part
            r.font.size = Pt(size)
            r.font.bold = (i % 2 == 1)
            r.font.name = face
            r.font.color.rgb = _c(tc)


# ==========================================================================
# LOGO
# ==========================================================================

def _add_logo(slide, logo_path, theme):
    if not logo_path or not os.path.exists(logo_path):
        return
    try:
        with Image.open(logo_path) as img:
            iw, ih = img.size
        ratio = iw / ih if ih else 1.0
        max_h = BADGE_H - 0.22
        max_w = BADGE_W - 0.30
        lh = min(max_h, max_w / ratio)
        lw = lh * ratio
        pad_x = (BADGE_W - lw) / 2
        pad_y = (BADGE_H - lh) / 2
        _rect(slide, BADGE_X, BADGE_Y, BADGE_W, BADGE_H,
              (255, 255, 255),
              line=theme.get("secondary", BASE_THEME["secondary"]),
              lw=1.0)
        slide.shapes.add_picture(
            logo_path,
            left   = _IN(BADGE_X + pad_x),
            top    = _IN(BADGE_Y + pad_y),
            width  = _IN(lw),
            height = _IN(lh),
        )
    except Exception as e:
        print(f"[logo] {e}")


# ==========================================================================
# SHARED HEADER / FOOTER
# ==========================================================================

def _header(slide, title, subtitle, num, theme, style=None):
    style = style if isinstance(style, dict) else {}
    header_variant = str(style.get("header_variant", "solid")).lower()

    if header_variant == "split":
        _rect(slide, 0, 0, W * 0.58, HEADER_H, theme["bg_dark"])
        _rect(slide, W * 0.58, 0, W * 0.42, HEADER_H, theme["primary"])
    elif header_variant == "banded":
        _rect(slide, 0, 0, W, HEADER_H, theme["bg_dark"])
        _rect(slide, 0, HEADER_H - 0.18, W, 0.18, theme["secondary"])
    else:
        _rect(slide, 0, 0, W, HEADER_H, theme["bg_dark"])

    # Title — width leaves room for logo badge AND number badge
    title_w = BADGE_X - 1.10 - 0.45
    t_size = _title_size(title, base=30)
    _tb(slide, title, 0.45, 0.18, title_w, 0.76, t_size,
        bold=True, color=theme["text_light"], face=theme["header_font"], shrink_to_fit=True)
    if subtitle:
        s_size = _subtitle_size(subtitle, base=13)
        # Keep subtitle readable over blue header variants.
        sub_color = _mix(theme["text_light"], theme["secondary"], 0.35)
        _tb(slide, subtitle, 0.45, 0.96, title_w, 0.38, s_size,
            italic=True, color=sub_color, face=theme["body_font"], shrink_to_fit=True)

    # Slide number badge
    bx = BADGE_X - 0.78
    badge_shape = str(style.get("badge_shape", "oval")).lower()
    if badge_shape == "rect":
        _rect(slide, bx, 0.22, 0.60, 0.60, theme["accent"], line=theme["secondary"], lw=1.0)
    else:
        _oval(slide, bx, 0.22, 0.60, 0.60, theme["accent"])
    num_str   = str(num)
    badge_pt  = 13 if len(num_str) > 1 else 15
    _tb(slide, num_str, bx + 0.02, 0.26, 0.56, 0.50, badge_pt,
        bold=True, color=theme["text_light"],
        face=theme["header_font"], align=PP_ALIGN.CENTER)


def _footer(slide, theme, style=None):
    style = style if isinstance(style, dict) else {}
    footer_variant = str(style.get("footer_variant", "solid")).lower()
    if footer_variant == "line":
        _rect(slide, 0, H - 0.07, W, 0.07, theme["secondary"])
    else:
        _rect(slide, 0, H - FOOTER_H, W, FOOTER_H, theme["bg_dark"])


def _bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _c(color)


def _card(slide, theme, style=None, x=None, y=None, w=None, h=None):
    """Draw the standard white content card and return its inner bounds."""
    style = style if isinstance(style, dict) else {}
    card_variant = str(style.get("card_variant", "outline")).lower()

    cx = x if x is not None else 0.40
    cy = y if y is not None else CONTENT_Y
    cw = w if w is not None else W - cx - 0.25
    ch = h if h is not None else CONTENT_H

    if card_variant == "soft":
        _rect(slide, cx, cy, cw, ch, _mix(theme["card_bg"], theme["secondary"], 0.12),
              line=_mix(theme["primary"], theme["secondary"], 0.5), lw=0.8)
    elif card_variant == "banded":
        _rect(slide, cx, cy, cw, ch, theme["card_bg"], line=theme["primary"], lw=1.0)
        _rect(slide, cx, cy, cw, 0.12, theme["secondary"])
    else:
        _rect(slide, cx, cy, cw, ch, theme["card_bg"],
              line=theme["primary"], lw=1.0)

    return cx + 0.20, cy + 0.18, cw - 0.40, ch - 0.30


# ==========================================================================
# LAYOUT RENDERERS
# ==========================================================================

def _render_bullets(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    _bullets(slide, spec.get("content", []),
             ix, iy, iw, ih,
             size=16, icon=spec.get("icon", "▸"),
             icon_color=stheme["secondary"],
             text_color=stheme["text_dark"],
             max_pts=6)
    _add_logo(slide, logo_path, stheme)


def _render_two_column(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    gap = 0.20
    cw  = (iw - gap) / 2

    left_title  = spec.get("left_title",  "Option A")
    right_title = spec.get("right_title", "Option B")
    left_pts    = spec.get("left_points",  [])
    right_pts   = spec.get("right_points", [])

    # Fallback: split content[] when LLM omits left/right point arrays
    if not left_pts and not right_pts:
        content = spec.get("content", [])
        mid = max(1, len(content) // 2)
        left_pts  = content[:mid]
        right_pts = content[mid:]

    icon = spec.get("icon", "▸")

    # Left column
    _rect(slide, ix, iy, cw, 0.46, stheme["primary"])
    _tb(slide, left_title, ix + 0.10, iy + 0.08, cw - 0.20, 0.34, 14,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])
    _bullets(slide, left_pts, ix + 0.10, iy + 0.58, cw - 0.20, ih - 0.66,
             size=14, icon=icon,
             icon_color=stheme["secondary"], text_color=stheme["text_dark"], max_pts=5)

    # Right column — accent2 to differentiate from left
    rx = ix + cw + gap
    _rect(slide, rx, iy, cw, 0.46, stheme["accent2"])
    _tb(slide, right_title, rx + 0.10, iy + 0.08, cw - 0.20, 0.34, 14,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])
    _bullets(slide, right_pts, rx + 0.10, iy + 0.58, cw - 0.20, ih - 0.66,
             size=14, icon=icon,
             icon_color=stheme["accent"], text_color=stheme["text_dark"], max_pts=5)

    _add_logo(slide, logo_path, stheme)


def _render_big_stat(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    stat_w = 3.70
    stat   = str(spec.get("stat", "—"))
    label  = spec.get("stat_label",  "")
    source = spec.get("stat_source", "")
    pts    = spec.get("content", [])

    # Stat panel
    _rect(slide, ix, iy, stat_w, ih, stheme["primary"])
    _rect(slide, ix, iy, stat_w, 0.07, stheme["accent2"])  # accent top rule

    stat_font = 72 if len(stat) <= 4 else 54
    _tb(slide, stat, ix + 0.10, iy + 0.38, stat_w - 0.20, 1.80, stat_font,
        bold=True, color=stheme["text_light"],
        face=stheme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, label, ix + 0.10, iy + 2.25, stat_w - 0.20, 0.75, 15,
        color=stheme["secondary"], face=stheme["body_font"], align=PP_ALIGN.CENTER)
    if source:
        _tb(slide, source, ix + 0.10, iy + 3.05, stat_w - 0.20, 0.50, 11,
            italic=True, color=(180, 210, 255),
            face=stheme["body_font"], align=PP_ALIGN.CENTER)

    # Progress bar when stat is a percentage
    try:
        pct = float(stat.replace("%", "").replace("+", "").strip())
        if 0 < pct <= 100:
            bar_x = ix + 0.30
            bar_w = stat_w - 0.60
            bar_y = iy + ih - 0.55
            _rect(slide, bar_x, bar_y, bar_w, 0.20, (255, 255, 255))
            _rect(slide, bar_x, bar_y, bar_w * (pct / 100), 0.20, stheme["accent2"])
    except ValueError:
        pass

    # Bullet panel
    bx = ix + stat_w + 0.22
    bw = iw - stat_w - 0.22
    _bullets(slide, pts, bx, iy + 0.10, bw, ih - 0.20,
             size=14, icon=spec.get("icon", "📊"),
             icon_color=stheme["secondary"],
             text_color=stheme["text_dark"], max_pts=5)

    _add_logo(slide, logo_path, stheme)


def _render_timeline(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    steps = [s for s in spec.get("steps", [])
             if isinstance(s, dict) and s.get("label") and s.get("detail")]

    if not steps:
        _tb(slide, "No timeline steps available.", ix, iy, iw, ih, 14,
            color=stheme["text_muted"])
        _add_logo(slide, logo_path, stheme)
        return

    n     = len(steps)
    tl_y  = iy + ih * 0.47
    sw    = iw / n
    DOT_R = 0.28
    DOT_r = 0.14

    accents = [stheme["primary"], stheme["accent2"], stheme["secondary"],
               stheme["accent"], stheme["primary"]]

    _rect(slide, ix, tl_y, iw, 0.06, stheme["secondary"])

    for i, step in enumerate(steps):
        cx  = ix + i * sw + sw / 2
        ac  = accents[i % len(accents)]

        _oval(slide, cx - DOT_R, tl_y - DOT_R, DOT_R * 2, DOT_R * 2, ac)
        _oval(slide, cx - DOT_r, tl_y - DOT_r, DOT_r * 2, DOT_r * 2, (255, 255, 255))
        _tb(slide, str(i + 1), cx - 0.20, tl_y - 0.16, 0.40, 0.32, 11,
            bold=True, color=ac, align=PP_ALIGN.CENTER)

        cw   = sw * 0.84
        cx_c = cx - cw / 2
        above = (i % 2 == 0)

        if above:
            cy = iy + 0.05
            ch = max(0.40, tl_y - DOT_R - 0.42 - cy)
            conn_top = cy + ch
            conn_bot = tl_y - DOT_R
            if conn_bot > conn_top + 0.02:
                _rect(slide, cx - 0.025, conn_top, 0.05, conn_bot - conn_top, stheme["secondary"])
        else:
            cy = tl_y + DOT_R + 0.28
            ch = max(0.40, iy + ih - 0.05 - cy)
            conn_top = tl_y + DOT_R
            conn_bot = cy
            if conn_bot > conn_top + 0.02:
                _rect(slide, cx - 0.025, conn_top, 0.05, conn_bot - conn_top, stheme["secondary"])

        _rect(slide, cx_c, cy, cw, ch, stheme["card_bg"], line=ac, lw=1.0)
        _tb(slide, step["label"], cx_c + 0.10, cy + 0.08, cw - 0.20, 0.42, 12,
            bold=True, color=ac, face=stheme["header_font"])
        dh = ch - 0.56
        if dh > 0.12:
            _tb(slide, step["detail"],
                cx_c + 0.10, cy + 0.52, cw - 0.20, dh, 11,
                color=stheme["text_dark"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)


def _render_icon_grid(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    items   = spec.get("grid_items", [])[:4]
    cols    = 2
    gap     = 0.16
    cw      = (iw - gap) / cols
    ch      = (ih - gap) / 2
    accents = [stheme["primary"], stheme["accent2"], stheme["secondary"], stheme["accent"]]

    for idx, gi in enumerate(items):
        col = idx % cols
        row = idx // cols
        cx  = ix + col * (cw + gap)
        cy  = iy + row * (ch + gap)
        ac  = accents[idx % len(accents)]

        _rect(slide, cx, cy, cw, ch, stheme["card_bg"], line=ac, lw=1.0)
        _rect(slide, cx, cy, 0.09, ch, ac)

        IR  = 0.38
        IOX = cx + 0.24
        IOY = cy + (ch - IR * 2) / 2
        _oval(slide, IOX, IOY, IR * 2, IR * 2, ac)
        raw_icon = str(gi.get("icon", gi.get("title", "A"))).strip()
        first_char = next((ch.upper() for ch in raw_icon if ch.isascii() and ch.isalnum()), "A")
        _tb(slide, first_char, IOX + 0.04, IOY + 0.08, IR * 2 - 0.08, IR * 1.6, 17,
            bold=True, color=(255, 255, 255),
            face=stheme["header_font"], align=PP_ALIGN.CENTER)

        TX = IOX + IR * 2 + 0.16
        TW = cw - (IOX - cx) - IR * 2 - 0.22
        _tb(slide, str(gi.get("title", "")),
            TX, cy + 0.12, TW, 0.46, 15,
            bold=True, color=stheme["primary"], face=stheme["header_font"])
        _tb(slide, str(gi.get("detail", "")),
            TX, cy + 0.60, TW, ch - 0.72, 13,
            color=stheme["text_dark"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)


def _render_case_study(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    company = spec.get("company", "Leading Organisation")
    result  = spec.get("result",  "")
    metrics = spec.get("metrics", [])
    bullets = spec.get("content", [])
    icon    = spec.get("icon", "▸")

    # Company banner
    _rect(slide, ix, iy, iw, 0.52, stheme["primary"])
    _tb(slide, f"📌  {company}",
        ix + 0.18, iy + 0.10, iw - 0.36, 0.36, 17,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])

    # Result ribbon
    if result:
        _rect(slide, ix, iy + 0.58, iw, 0.48, stheme["accent2"])
        _tb(slide, f"🏆  {result}",
            ix + 0.14, iy + 0.66, iw - 0.28, 0.36, 13,
            bold=True, color=stheme["text_light"], face=stheme["body_font"])

    # Metric progress bars
    metrics_h = 0.0
    for i, m in enumerate(metrics[:3]):
        my    = iy + 1.18 + i * 0.50
        label = str(m.get("label", ""))
        value = min(100, max(0, int(m.get("value", 50))))
        bar_x = ix + 3.10
        bar_w = iw - 3.30
        _tb(slide, label, ix + 0.14, my + 0.02, 2.90, 0.30, 12,
            color=stheme["text_dark"])
        _rect(slide, bar_x, my + 0.06, bar_w, 0.22, (225, 232, 245))
        _rect(slide, bar_x, my + 0.06, bar_w * (value / 100), 0.22, stheme["accent2"])
        _tb(slide, f"+{value}%",
            bar_x + bar_w * (value / 100) + 0.06, my, 0.55, 0.30,
            11, bold=True, color=stheme["primary"])
        metrics_h = (i + 1) * 0.50

    # Supporting bullets below metrics
    bullet_y = iy + 1.18 + metrics_h + 0.12
    bullet_h = ih - (bullet_y - iy) - 0.10
    if bullet_h > 0.30 and bullets:
        _bullets(slide, bullets, ix + 0.10, bullet_y, iw - 0.20, bullet_h,
                 size=13, icon=icon,
                 icon_color=(180, 40, 40),
                 text_color=stheme["text_dark"], max_pts=4)

    _add_logo(slide, logo_path, stheme)


def _render_table(slide, spec, num, theme, logo_path):
    style = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    columns = spec.get("table_columns", [])
    rows = spec.get("table_rows", [])
    if not columns or not rows:
        _tb(slide, "No table data available.", ix, iy, iw, ih, 14, color=stheme["text_muted"])
        _add_logo(slide, logo_path, stheme)
        return

    ncols = max(1, min(5, len(columns)))
    cols = [str(c) for c in columns[:ncols]]
    row_data = []
    for r in rows[:6]:
        if not isinstance(r, list):
            continue
        clean = [str(v) for v in r[:ncols]]
        while len(clean) < ncols:
            clean.append("")
        row_data.append(clean)
    if not row_data:
        _tb(slide, "No table rows available.", ix, iy, iw, ih, 14, color=stheme["text_muted"])
        _add_logo(slide, logo_path, stheme)
        return

    gap = 0.02
    col_w = (iw - (ncols - 1) * gap) / ncols
    header_h = 0.55
    max_rows = len(row_data)
    row_h = min(0.65, max(0.42, (ih - header_h - 0.08) / max_rows))

    # Header row
    for cidx, c in enumerate(cols):
        cx = ix + cidx * (col_w + gap)
        _rect(slide, cx, iy, col_w, header_h, stheme["primary"], line=stheme["secondary"], lw=0.8)
        _tb(slide, c, cx + 0.06, iy + 0.08, col_w - 0.12, header_h - 0.12, 12,
            bold=True, color=stheme["text_light"], face=stheme["header_font"], align=PP_ALIGN.CENTER, shrink_to_fit=True)

    # Body rows
    for ridx, row in enumerate(row_data):
        ry = iy + header_h + ridx * row_h
        fill = stheme["card_bg"] if ridx % 2 == 0 else _mix(stheme["card_bg"], stheme["secondary"], 0.08)
        for cidx, cell in enumerate(row):
            cx = ix + cidx * (col_w + gap)
            _rect(slide, cx, ry, col_w, row_h, fill, line=_mix(stheme["primary"], stheme["card_bg"], 0.65), lw=0.5)
            _tb(slide, cell, cx + 0.06, ry + 0.06, col_w - 0.12, row_h - 0.10, 11,
                color=stheme["text_dark"], face=stheme["body_font"], align=PP_ALIGN.LEFT, shrink_to_fit=True)

    _add_logo(slide, logo_path, stheme)


# Layout name → renderer
_RENDERERS = {
    "bullets":    _render_bullets,
    "two_column": _render_two_column,
    "big_stat":   _render_big_stat,
    "timeline":   _render_timeline,
    "icon_grid":  _render_icon_grid,
    "case_study": _render_case_study,
    "table":      _render_table,
}


# ==========================================================================
# SPECIAL SLIDES
# ==========================================================================

def _intro_slide(prs, topic, theme, logo_path, style=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    style = style if isinstance(style, dict) else {}
    stheme = _slide_theme(theme, style, slide_idx=0)
    _bg(slide, stheme["bg_dark"])

    # Left panel + vertical separator
    _rect(slide, 0, 0, W * 0.52, H, stheme["primary"])
    _rect(slide, W * 0.507, 0, 0.09, H, stheme["secondary"])

    # Decorative circles on right only (no overlap with content)
    _oval(slide, W - 3.8, H - 3.8, 4.5, 4.5, stheme["accent2"])
    _oval(slide, W - 2.4, H - 2.4, 2.8, 2.8, stheme["bg_dark"])

    # Title — capped at 6.4" so it never overflows left panel
    _tb(slide, f"🚀  {topic}", 0.55, 0.85, 6.40, 1.20, 32,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])
    _tb(slide, "Are You Ready?", 0.55, 2.12, 6.40, 0.72, 28,
        bold=True, color=stheme["secondary"], face=stheme["header_font"])
    _tb(slide,
        "A data-driven look at opportunities, challenges, and real-world wins",
        0.55, 2.96, 6.40, 0.60, 15,
        italic=True, color=(180, 210, 255), face=stheme["body_font"])

    # Two info cards — topic variable used, NOT "{topic}" literal
    for i, (border, hdr_color, hdr_txt, body_txt) in enumerate([
        (stheme["accent2"], stheme["accent2"],
         "📊  DID YOU KNOW?",
         "**85%** of executives say AI will give their company a competitive edge by 2026 (PwC)."),
        (stheme["primary"], stheme["primary"],
         "❓  THE BIG QUESTION",
         f"How can your organization turn **{topic}** from hype into **measurable ROI**?"),
    ]):
        cx = 0.55 + i * 3.70
        _rect(slide, cx, 3.72, 3.45, 2.35, (255, 255, 255), line=border, lw=2.0)
        _tb(slide, hdr_txt,  cx + 0.12, 3.82, 3.22, 0.40, 13, bold=True, color=hdr_color)
        _tb(slide, body_txt, cx + 0.12, 4.30, 3.22, 1.55, 13, color=stheme["text_dark"])

    _add_logo(slide, logo_path, stheme)
    return slide


def _summary_slide(prs, topic, theme, logo_path, style=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    style = style if isinstance(style, dict) else {}
    stheme = _slide_theme(theme, style, slide_idx=999)
    _bg(slide, stheme["bg_light"])
    _header(slide, "Key Takeaways", f"What we learned about {topic}", "★", stheme, style=style)
    _footer(slide, stheme, style=style)

    ix, iy, iw, ih = _card(slide, stheme, style=style)
    col_w = (iw - 0.40) / 3

    takeaways = [
        ("1️⃣  Problem First",
         f"Define a clear pain point before choosing technology — "
         f"70% of {topic} projects fail without it.\n\n"
         "✔  Map the workflow gap first.\n"
         "✔  Establish baseline metrics.\n"
         "✔  Secure an executive sponsor."),
        ("2️⃣  Proven ROI Exists",
         "Real deployments show measurable returns:\n\n"
         "✔  Mayo Clinic: 60% faster radiology.\n"
         "✔  NHS England: £200M saved on sepsis.\n"
         "✔  Google Health: 99% mammogram accuracy."),
        ("3️⃣  Start Small, Scale Fast",
         "8-week pilots with clear KPIs work best.\n\n"
         "✔  One department → measure → expand.\n"
         "✔  Clinician buy-in reduces failure risk.\n"
         "✔  Build compliance in from day one."),
    ]

    for i, (title, text) in enumerate(takeaways):
        cx = ix + i * (col_w + 0.20)
        _rect(slide, cx, iy, col_w, ih, stheme["card_bg"],
              line=stheme["primary"], lw=1.5)
        _rect(slide, cx, iy, col_w, 0.46, stheme["primary"])
        _tb(slide, title, cx + 0.12, iy + 0.08, col_w - 0.24, 0.34, 14,
            bold=True, color=stheme["text_light"], face=stheme["header_font"])
        _tb(slide, text, cx + 0.12, iy + 0.58, col_w - 0.24, ih - 0.66,
            13, color=stheme["text_dark"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)
    return slide


def _closing_slide(prs, topic, theme, logo_path, style=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    style = style if isinstance(style, dict) else {}
    stheme = _slide_theme(theme, style, slide_idx=1000)
    _bg(slide, stheme["bg_dark"])
    _rect(slide, 0,    0, 0.35, H, stheme["primary"])
    _rect(slide, 0.35, 0, 0.09, H, stheme["secondary"])

    # Decorative circles — far right, clear of text zone
    _oval(slide, W - 4.2, H - 4.2, 5.0, 5.0, (0, 80, 160))
    _oval(slide, W - 2.8, H - 2.8, 3.2, 3.2, stheme["bg_dark"])

    SAFE_W = W - 4.8  # keep text away from circles

    _tb(slide, "YOUR NEXT STEP",
        0.65, 1.40, SAFE_W, 1.00, 40,
        bold=True, color=stheme["text_light"],
        face=stheme["header_font"], align=PP_ALIGN.LEFT)
    _tb(slide,
        f"Don't let {topic} be just another trend. Start your pilot today.",
        0.65, 2.55, SAFE_W, 0.65, 17,
        italic=True, color=stheme["secondary"],
        face=stheme["body_font"], align=PP_ALIGN.LEFT)

    # All action items in ONE textbox — prevents the 3-box misalignment bug
    bx = slide.shapes.add_textbox(_IN(0.65), _IN(3.35), _IN(SAFE_W), _IN(1.80))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    actions = [
        "🔹  Define one high-impact use case (diagnostics, scheduling, or monitoring)",
        "🔹  Run an 8-week pilot with measurable KPIs — accuracy, time saved, cost reduction",
        "🔹  Document outcomes, build the business case, and scale organisation-wide",
    ]
    first = True
    for act in actions:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(8)
        p.space_after  = Pt(8)
        r = p.add_run()
        r.text = act
        r.font.size  = Pt(15)
        r.font.name  = stheme["body_font"]
        r.font.color.rgb = _c(stheme["text_light"])

    # "Questions?" — near-white so it reads on dark background
    _tb(slide, "Questions? Let's talk.",
        0.65, 5.35, SAFE_W, 0.50, 14,
        italic=True, color=(200, 225, 255),
        face=stheme["body_font"], align=PP_ALIGN.LEFT)

    _add_logo(slide, logo_path, stheme)
    return slide


# ==========================================================================
# CONTENT GENERATION  (LLM prompt — strict field requirements)
# ==========================================================================

def generate_slide_content(topic: str, num_slides: int = 5,
                            tone: str = "Professional") -> dict:
    prompt = f"""
You are an expert presentation designer. Create a {tone} PowerPoint on "{topic}".
Return ONLY valid JSON — no markdown fences, no preamble, no explanation.

=== STRICT RULES ===
1. Generate exactly {num_slides} slides inside a top-level "slides" array.
2. Every slide MUST include: "title", "subtitle", "layout", "icon", "content", and "style".
3. Always include a clear opening title slide as slide 1.
4. Always include a closing "Thank You" or "Q&A" slide as the final slide.
5. VARY layouts — do NOT use the same layout twice in a row.
6. VARY visual pattern — do NOT use the same `style.pattern_name` twice in a row.
7. Use a clean LIGHT theme only (white and blue family). Avoid dark backgrounds.
8. "content" bullets: 5-7 items, 22-34 words each, use **bold** for key terms.
   Every bullet must cite a REAL company, statistic, or year — never say "many companies".
9. If comparisons, benchmarks, rankings, or phased plans are present, use a "table" layout where appropriate.

=== DESIGN SYSTEM (LLM MUST DECIDE, NOT FIXED) ===
Add top-level object "design_system":
{{
  "theme": {{
    "primary": "#RRGGBB",
    "secondary": "#RRGGBB",
    "accent": "#RRGGBB",
    "accent2": "#RRGGBB",
    "bg_light": "#RRGGBB",
    "bg_dark": "#RRGGBB",
    "text_dark": "#RRGGBB",
    "text_light": "#RRGGBB",
    "card_bg": "#RRGGBB",
    "header_font": "font name",
    "body_font": "font name"
  }},
  "intro_style":   {{"surface":"light","header_variant":"split","card_variant":"banded"}},
  "summary_style": {{"surface":"light","header_variant":"banded","card_variant":"outline"}},
  "closing_style": {{"surface":"light","header_variant":"solid","card_variant":"soft"}}
}}

Each slide.style must include:
- pattern_name: unique descriptive style label (e.g. "neo-bento", "editorial-ribbon", "technical-dark")
- surface: one of ["light","tint"]
- header_variant: one of ["solid","split","banded"]
- card_variant: one of ["outline","soft","banded"]
- footer_variant: one of ["solid","line"]
- badge_shape: one of ["oval","rect"]
- accent_rotation: one of ["static","auto"]

=== LAYOUT RULES (each layout has REQUIRED extra fields) ===

"bullets"
  content: [list of bullets]

"two_column"
  left_title:  string  ← REQUIRED, descriptive (NOT the word "Left")
  right_title: string  ← REQUIRED, descriptive (NOT the word "Right")
  left_points:  [...]  ← REQUIRED, 4-5 bullets
  right_points: [...]  ← REQUIRED, 4-5 bullets
  content: []

"big_stat"
  stat:        string  ← REQUIRED real value e.g. "40%", "$2.5B", "3x" (NOT "100%")
  stat_label:  string  ← ≤ 8 words describing what the stat measures
  stat_source: string  ← source/year e.g. "Mount Sinai Hospital, 2022"
  content:     [...]   ← 5-6 supporting bullets

"timeline"
  steps: [             ← REQUIRED, exactly 4 items
    {{"label": "short name (≤4 words)", "detail": "one sentence, 15-20 words"}}
  ]
  content: []

"icon_grid"
  grid_items: [        ← REQUIRED, exactly 4 items
    {{"icon": "keyword", "title": "2-3 words", "detail": "12-18 words"}}
  ]
  content: []

"case_study"
  company: string      ← REQUIRED, real company name (NOT "Case Study" or "Company")
  result:  string      ← one-sentence headline outcome
  metrics: [           ← 2-3 items
    {{"label": "metric name", "value": integer between 1 and 99}}
  ]
  content: [...]       ← 4-5 supporting bullets

"table"
  table_columns: [3-5 short column headers]
  table_rows: [4-6 rows, each row must match number of columns]
  content: []  (optional notes)

=== JSON FORMAT ===
{{
  "design_system": {{
    "theme": {{
      "primary": "#0B5FFF",
      "secondary": "#4DA3FF",
      "accent": "#0A3D91",
      "accent2": "#00A3A3",
      "bg_light": "#F7FAFF",
      "bg_dark": "#0B2A4A",
      "text_dark": "#1E1E1E",
      "text_light": "#FFFFFF",
      "card_bg": "#FFFFFF",
      "header_font": "Calibri",
      "body_font": "Calibri"
    }},
    "intro_style": {{"surface":"light","header_variant":"split","card_variant":"banded"}},
    "summary_style": {{"surface":"light","header_variant":"banded","card_variant":"outline"}},
    "closing_style": {{"surface":"light","header_variant":"solid","card_variant":"soft"}}
  }},
  "slides": [
    {{
      "title": "...",
      "subtitle": "...",
      "layout": "bullets",
      "icon": "▸",
      "content": ["...", "...", "...", "..."],
      "style": {{
        "pattern_name": "neo-bento",
        "surface": "light",
        "header_variant": "split",
        "card_variant": "banded",
        "footer_variant": "line",
        "badge_shape": "rect",
        "accent_rotation": "auto"
      }}
    }}
  ]
}}
"""
    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.75,
        max_tokens=6000,
    )
    raw = (resp.choices[0].message.content or "").strip()
    if raw.startswith("```"):
        raw = raw.replace("```json", "").replace("```", "").strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Fallback for occasional model wrappers around valid JSON.
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise
        data = json.loads(raw[start:end + 1])

    if not isinstance(data, dict):
        data = {"design_system": {}, "slides": []}
    if not isinstance(data.get("design_system"), dict):
        data["design_system"] = {}
    # Keep overall style in light white/blue direction.
    ds = data["design_system"]
    for key in ("intro_style", "summary_style", "closing_style"):
        if not isinstance(ds.get(key), dict):
            ds[key] = {}
        ds[key]["surface"] = "light"
    if not isinstance(data.get("slides"), list):
        data["slides"] = []

    # ── Validate & sanitise every slide ─────────────────────────────────────
    clean_slides = []
    for idx, slide in enumerate(data.get("slides", []), start=1):
        if not isinstance(slide, dict):
            continue

        slide.setdefault("title", f"Slide {idx}")
        slide.setdefault("subtitle", "")
        layout = slide.get("layout", "bullets")
        slide["icon"] = _safe_icon(slide.get("icon"), default="▸")
        slide.setdefault("content", [])
        if not isinstance(slide.get("style"), dict):
            slide["style"] = {}
        style = slide["style"]
        style.setdefault("pattern_name", f"pattern_{idx}")
        style.setdefault("surface", "light")
        style.setdefault("header_variant", "solid")
        style.setdefault("card_variant", "outline")
        style.setdefault("footer_variant", "solid")
        style.setdefault("badge_shape", "oval")
        style.setdefault("accent_rotation", "static")

        # Restrict variants to known values for renderer safety.
        if style["surface"] not in ("light", "tint"):
            style["surface"] = "light"
        if style["surface"] == "dark":
            style["surface"] = "light"
        if style["header_variant"] not in ("solid", "split", "banded"):
            style["header_variant"] = "solid"
        if style["card_variant"] not in ("outline", "soft", "banded"):
            style["card_variant"] = "outline"
        if style["footer_variant"] not in ("solid", "line"):
            style["footer_variant"] = "solid"
        if style["badge_shape"] not in ("oval", "rect"):
            style["badge_shape"] = "oval"
        if style["accent_rotation"] not in ("static", "auto"):
            style["accent_rotation"] = "static"

        # Coerce all list fields to plain strings
        for key in ("content", "left_points", "right_points"):
            if key in slide and isinstance(slide[key], list):
                slide[key] = [
                    str(x.get("text", x)) if isinstance(x, dict) else str(x)
                    for x in slide[key]
                ]
        if not isinstance(slide.get("content"), list):
            slide["content"] = []
        slide["content"] = [str(x).strip() for x in slide["content"] if str(x).strip()]
        if layout in ("bullets", "big_stat") and len(slide["content"]) < 5:
            while len(slide["content"]) < 5:
                slide["content"].append(
                    f"**Execution focus**: add one measurable KPI, one owner, and one delivery date so this {topic} initiative is actionable, auditable, and presentation-ready."
                )

        # two_column: reject generic header labels; fallback split content[]
        if layout == "two_column":
            for k, default in (("left_title", "Before"), ("right_title", "After")):
                v = slide.get(k, "").strip().lower()
                if not v or v in ("left", "right", "column 1", "column 2", "option a", "option b"):
                    slide[k] = default
            if not slide.get("left_points") and not slide.get("right_points"):
                content = slide.get("content", [])
                mid = max(1, len(content) // 2)
                slide["left_points"]  = content[:mid]
                slide["right_points"] = content[mid:]
            # Keep both columns visually full.
            for side in ("left_points", "right_points"):
                if not isinstance(slide.get(side), list):
                    slide[side] = []
                slide[side] = [str(x).strip() for x in slide[side] if str(x).strip()]
                while len(slide[side]) < 4:
                    slide[side].append(
                        f"**Operational step**: define owner, timeline, and KPI to move this {topic} track from concept to production impact."
                    )

        # big_stat: reject meaningless "100%" or empty stat
        if layout == "big_stat":
            stat = str(slide.get("stat", "")).strip()
            if not stat or stat in ("100%", "100", ""):
                slide["stat"] = "—"

        # timeline: validate steps; inject fallbacks if broken
        if layout == "timeline":
            clean = [
                s for s in slide.get("steps", [])
                if isinstance(s, dict) and s.get("label") and s.get("detail")
            ]
            slide["steps"] = clean if clean else [
                {"label": f"Phase {i+1}", "detail": "Details to be confirmed."}
                for i in range(4)
            ]
            for s in slide["steps"]:
                s["label"]  = str(s.get("label", ""))
                s["detail"] = str(s.get("detail", ""))

        # icon_grid: coerce nested dict fields
        if layout == "icon_grid":
            if not isinstance(slide.get("grid_items"), list):
                slide["grid_items"] = []
            fixed_grid = []
            for gi in slide.get("grid_items", []):
                if not isinstance(gi, dict):
                    gi = {"icon": "•", "title": str(gi), "detail": ""}
                for k in ("icon", "title", "detail"):
                    if k in gi and isinstance(gi[k], dict):
                        gi[k] = str(gi[k])
                fixed_grid.append(gi)
            slide["grid_items"] = fixed_grid[:4]

        # case_study: reject generic company names
        if layout == "case_study":
            co = slide.get("company", "").strip().lower()
            if not co or co in ("case study", "company", "organization", ""):
                slide["company"] = "Leading Organisation"
            metrics = slide.get("metrics", [])
            if not isinstance(metrics, list):
                metrics = []
            fixed_metrics = []
            for m in metrics[:3]:
                if not isinstance(m, dict):
                    m = {"label": "Impact", "value": m}
                fixed_metrics.append({
                    "label": str(m.get("label", "Impact")),
                    "value": _coerce_int(m.get("value", 50), default=50, min_value=1, max_value=99),
                })
            slide["metrics"] = fixed_metrics
            while len(slide["content"]) < 4:
                slide["content"].append(
                    f"**Scale lesson**: standardize process handoff and governance so this {topic} result is repeatable across teams and regions."
                )

        # table: sanitize columns/rows and backfill when sparse
        if layout == "table":
            cols = slide.get("table_columns", [])
            rows = slide.get("table_rows", [])
            if not isinstance(cols, list):
                cols = []
            cols = [str(c).strip() for c in cols if str(c).strip()][:5]
            if len(cols) < 3:
                cols = ["Category", "Current State", "Target State"]

            clean_rows = []
            if isinstance(rows, list):
                for r in rows:
                    if not isinstance(r, list):
                        continue
                    row = [str(v).strip() for v in r][:len(cols)]
                    while len(row) < len(cols):
                        row.append("")
                    clean_rows.append(row)

            while len(clean_rows) < 4:
                clean_rows.append([
                    "Workstream",
                    f"Baseline for {topic}",
                    "KPI-driven improvement plan",
                ][:len(cols)] + [""] * max(0, len(cols) - 3))

            slide["table_columns"] = cols
            slide["table_rows"] = clean_rows[:6]

        clean_slides.append(slide)

    # Prevent immediate repeated visual pattern if model repeats by mistake.
    for i in range(1, len(clean_slides)):
        cur = clean_slides[i].get("style", {})
        prev = clean_slides[i - 1].get("style", {})
        if cur.get("pattern_name", "").strip().lower() == prev.get("pattern_name", "").strip().lower():
            cur["pattern_name"] = f'{cur.get("pattern_name", "pattern")}_{i+1}'
            # Also change one variant so it looks visibly different.
            cur["header_variant"] = "banded" if prev.get("header_variant") != "banded" else "split"

    # Topic-aware structure nudges: encourage non-repetitive middle sections.
    layouts_present = {str(s.get("layout", "")) for s in clean_slides}
    topic_l = topic.lower()
    preferred = []
    if any(k in topic_l for k in ("architecture", "framework", "cloud", "system", "infrastructure")):
        preferred = ["table", "timeline"]
    elif any(k in topic_l for k in ("strategy", "market", "business", "growth", "roi")):
        preferred = ["big_stat", "case_study"]
    elif any(k in topic_l for k in ("ai", "ml", "model", "data")):
        preferred = ["timeline", "icon_grid"]
    else:
        preferred = ["table", "big_stat"]

    def _first_bullets_idx():
        for j, sl in enumerate(clean_slides):
            if sl.get("layout") == "bullets":
                return j
        return None

    for pref in preferred:
        if pref in layouts_present:
            continue
        idx = _first_bullets_idx()
        if idx is None:
            break
        src = clean_slides[idx]
        points = src.get("content", [])
        if pref == "timeline":
            steps = []
            for k, p in enumerate(points[:4], start=1):
                steps.append({"label": f"Phase {k}", "detail": str(p)})
            src["layout"] = "timeline"
            src["steps"] = steps if steps else [
                {"label": "Phase 1", "detail": f"Initial plan for {topic}."},
                {"label": "Phase 2", "detail": "Pilot and validation."},
                {"label": "Phase 3", "detail": "Scale and optimize."},
                {"label": "Phase 4", "detail": "Govern and continuously improve."},
            ]
            src["content"] = []
        elif pref == "table":
            src["layout"] = "table"
            src["table_columns"] = ["Area", "Current", "Target"]
            rows = []
            for p in points[:5]:
                rows.append(["Workstream", str(p)[:60], "Measured improvement"])
            src["table_rows"] = rows if rows else [
                ["Architecture", f"Baseline for {topic}", "Scalable blueprint"],
                ["Operations", "Manual processes", "Automated controls"],
                ["Security", "Reactive checks", "Policy-driven governance"],
                ["Metrics", "Ad-hoc reporting", "KPI dashboard"],
            ]
            src["content"] = []
        elif pref == "big_stat":
            src["layout"] = "big_stat"
            src["stat"] = src.get("stat", "42%")
            src["stat_label"] = src.get("stat_label", "Target Improvement")
            src["stat_source"] = src.get("stat_source", "Industry benchmark")
        elif pref == "case_study":
            src["layout"] = "case_study"
            src["company"] = src.get("company", "Reference Enterprise")
            src["result"] = src.get("result", f"Demonstrated practical gains for {topic}")
            src["metrics"] = src.get("metrics", [
                {"label": "Efficiency", "value": 40},
                {"label": "Quality", "value": 35},
                {"label": "Adoption", "value": 30},
            ])
        layouts_present.add(pref)

    seed = _topic_seed(topic)
    variant = seed % 3

    # Add a topic-aware opening cover slide every time.
    clean_slides.insert(0, _opening_slide_spec(topic, tone, variant))

    # Add topic-aware closing thank-you slide unless already present.
    if clean_slides:
        last_title = str(clean_slides[-1].get("title", "")).strip().lower()
        if not any(k in last_title for k in ("thank", "q&a", "questions", "qa")):
            clean_slides.append(_closing_slide_spec(topic, variant))

    data["slides"] = clean_slides

    return data


# ==========================================================================
# PUBLIC API
# ==========================================================================

def create_ppt(slide_data, topic: str,
               logo_path: str = None,
               tone: str = "Professional") -> str:

    prs = Presentation()
    prs.slide_width  = _IN(W)
    prs.slide_height = _IN(H)

    payload = slide_data if isinstance(slide_data, dict) else {"slides": slide_data}
    design_system = payload.get("design_system", {}) if isinstance(payload, dict) else {}
    theme = _normalize_theme(design_system.get("theme", {}))
    # Enforce regular light white/blue visual baseline.
    theme["bg_light"] = (255, 255, 255)
    theme["card_bg"] = (255, 255, 255)
    theme["bg_dark"] = _mix(theme["primary"], (255, 255, 255), 0.28)

    has_logo = bool(logo_path and os.path.exists(logo_path))
    logo     = logo_path if has_logo else None

    slides = payload.get("slides", [])
    if not isinstance(slides, list):
        slides = []

    # Render exactly the slides planned by the LLM (no hardcoded intro/summary/closing).
    for i, spec in enumerate(slides, start=1):
        if not isinstance(spec, dict):
            continue
        spec.setdefault("title", f"Slide {i}")
        spec.setdefault("subtitle", "")
        spec.setdefault("content", [])
        if not isinstance(spec.get("style"), dict):
            spec["style"] = {}
        spec["style"]["surface"] = "light" if spec["style"].get("surface") == "dark" else spec["style"].get("surface", "light")
        layout   = spec.get("layout", "bullets")
        renderer = _RENDERERS.get(layout, _render_bullets)
        slide    = prs.slides.add_slide(prs.slide_layouts[6])
        renderer(slide, spec, i, theme, logo)

    # Fallback when LLM returned no valid slides.
    if len(prs.slides) == 0:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _render_bullets(
            slide,
            {
                "title": topic,
                "subtitle": "Auto-generated summary",
                "layout": "bullets",
                "icon": "▸",
                "content": [
                    f"**Overview**: no valid slide plan was returned for {topic}, so a fallback summary slide was created.",
                    "**Next step**: regenerate with the same topic to get a full multi-layout presentation.",
                    "**Tip**: include target audience and use-case context for richer and more structured slide content.",
                    "**Validation**: the service now accepts dynamic layouts including table, timeline, metrics, and case studies.",
                    "**Output**: this fallback prevents empty deck failures and preserves downloadable PPT generation.",
                ],
                "style": {"surface": "light", "header_variant": "solid", "card_variant": "outline"},
            },
            1,
            theme,
            logo,
        )

    safe = topic.replace(" ", "_").replace("/", "_")[:60]
    os.makedirs("generated", exist_ok=True)
    path = f"generated/{safe}.pptx"
    prs.save(path)
    return path
