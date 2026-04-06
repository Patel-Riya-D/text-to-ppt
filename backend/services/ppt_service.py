"""
ppt_service.py  –  Fully Dynamic AI PPT Generator
===================================================

IMPROVEMENTS over previous version:
1. _pick_deck_profile() is now truly RANDOM (was deterministic hash — same topic always same profile)
2. Each profile has a DISTINCT color family (was all blue nudges — magazine/executive/tech all looked same)
3. bg_dark uses real dark values per profile (was 28% white mix — headers appeared light)
4. Font pairs are profile-specific (was always Calibri)
5. Added 2 new profiles: "bold" and "minimal" — 6 total for more variety
6. Added dark cover mode that uses full dark background for cover slide variety
7. Accent color is derived per-profile, not from a single formula
"""

import os
import json
import re
import random
from openai import AzureOpenAI
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

# ==========================================================================
# DECK CONFIG  — single source of truth for every numeric / structural constant
# ==========================================================================

DECK_CFG = dict(
    slide_w          = 13.3,
    slide_h          = 7.5,

    header_h         = 1.40,
    footer_h         = 0.25,
    header_content_gap = 0.12,

    badge_w          = 1.65,
    badge_h          = 0.90,
    badge_margin_r   = 0.18,
    badge_margin_t   = 0.18,
    badge_inner_pad  = 0.15,

    num_badge_w      = 0.60,
    num_badge_h      = 0.60,
    num_badge_gap    = 0.78,
    num_badge_top    = 0.22,
    num_font_1digit  = 15,
    num_font_2digit  = 13,

    card_margin_l    = 0.40,
    card_margin_r    = 0.25,
    card_inner_pad   = 0.20,
    card_inner_pad_v = 0.18,
    card_inner_pad_b = 0.30,

    title_base_pt    = 30,
    title_size_steps = [(42, 0), (65, -4), (90, -7), (999, -10)],
    cover_title_base = 46,

    subtitle_base_pt  = 13,
    subtitle_size_steps = [(60, 0), (95, -1), (999, -2)],
    cover_subtitle_base = 18,

    bullet_size_base     = 16,
    bullet_max_default   = 6,
    bullet_sparse_2_size = 24,
    bullet_sparse_3_size = 22,
    bullet_sparse_4_size = 19,
    bullet_space_sparse  = 11,
    bullet_space_normal  = 6,

    two_col_gap          = 0.20,
    two_col_header_h     = 0.46,
    two_col_bullet_size  = 14,
    two_col_min_bullets  = 4,

    stat_panel_w         = 3.70,
    stat_font_short      = 72,
    stat_font_long       = 54,
    stat_label_size      = 15,
    stat_source_size     = 11,
    stat_bar_h           = 0.20,
    stat_bar_margin_x    = 0.30,
    stat_bullet_size     = 14,

    timeline_dot_outer   = 0.28,
    timeline_dot_inner   = 0.14,
    timeline_rail_h      = 0.06,
    timeline_label_size  = 12,
    timeline_detail_size = 11,
    timeline_conn_w      = 0.05,

    icon_grid_cols       = 2,
    icon_grid_gap        = 0.16,
    icon_grid_radius     = 0.38,
    icon_grid_title_size = 15,
    icon_grid_detail_size= 13,
    icon_grid_bar_w      = 0.09,

    case_banner_h        = 0.52,
    case_banner_size     = 17,
    case_ribbon_h        = 0.48,
    case_ribbon_size     = 13,
    case_metric_row_h    = 0.50,
    case_metric_label_w  = 2.90,
    case_metric_label_size = 12,
    case_metric_pct_size = 11,
    case_bullet_size     = 13,

    table_header_h       = 0.55,
    table_header_size    = 12,
    table_cell_size      = 11,
    table_max_cols       = 5,
    table_max_rows       = 6,
    table_col_gap        = 0.02,
    table_cell_pad       = 0.06,
    chart_max_items      = 5,
    chart_label_size     = 12,
    chart_value_size     = 11,
    chart_source_size    = 10,
    chart_bar_h          = 0.24,
    chart_bar_gap        = 0.18,
    split_image_ratio    = 0.44,
    split_caption_size   = 12,
    split_bullet_size    = 14,

    cover_bar_top_h      = 1.25,
    cover_bar_top_gap    = 0.08,
    cover_footer_h       = 0.18,
    cover_title_x        = 0.70,
    cover_title_y        = 2.05,
    cover_title_pad      = 2.00,
    cover_title_h        = 1.95,
    cover_sub_y          = 4.00,
    cover_sub_h          = 0.75,
    cover_bullet_y       = 4.95,
    cover_bullet_h       = 1.80,
    cover_bullet_max     = 3,
    cover_bullet_size    = 14,

    header_title_x       = 0.45,
    header_title_y       = 0.18,
    header_title_h       = 0.76,
    header_sub_y         = 0.96,
    header_sub_h         = 0.38,

    accent_rotation_keys = ["primary", "secondary", "accent2", "accent"],

    tint_bg_mix          = 0.10,
    tint_card_mix        = 0.05,
    sub_color_mix        = 0.35,
    stripe_mix           = 0.65,

    fallback_bullet_marker = "**Execution focus**",

    allowed_surfaces        = ("light", "tint"),
    allowed_header_variants = ("solid", "split", "banded"),
    allowed_card_variants   = ("outline", "soft", "banded"),
    allowed_footer_variants = ("solid", "line"),
    allowed_badge_shapes    = ("oval", "rect"),
    allowed_accent_rotations= ("static", "auto"),

    min_bullets_by_layout = {"bullets": 5, "big_stat": 5, "case_study": 4, "image_text_split": 4},

    two_col_generic_headers = {"left", "right", "column 1", "column 2",
                               "option a", "option b"},
    case_study_generic_companies = {"case study", "company", "organization", ""},
    stat_placeholder_values = {"100%", "100", ""},
)

# ==========================================================================
# PROFILE SYSTEM — 6 visually distinct profiles, randomly selected
# ==========================================================================

DECK_PROFILE_CFG = {
    "classic":   {
        "cover_mode": "bands",    "header_mode": "full",    "footer_mode": "std",
        "card_shift_x": 0.00, "card_shift_y": 0.00, "card_shrink_w": 0.00,
    },
    "magazine":  {
        "cover_mode": "sidebar",  "header_mode": "sidebar", "footer_mode": "line",
        "card_shift_x": 0.18, "card_shift_y": 0.05, "card_shrink_w": 0.10,
    },
    "executive": {
        "cover_mode": "centered", "header_mode": "band",    "footer_mode": "std",
        "card_shift_x": 0.00, "card_shift_y": 0.02, "card_shrink_w": 0.00,
    },
    "tech":      {
        "cover_mode": "grid",     "header_mode": "split",   "footer_mode": "line",
        "card_shift_x": 0.08, "card_shift_y": 0.04, "card_shrink_w": 0.06,
    },
    "bold":      {
        "cover_mode": "dark_full","header_mode": "band",    "footer_mode": "std",
        "card_shift_x": 0.00, "card_shift_y": 0.00, "card_shrink_w": 0.00,
    },
    "minimal":   {
        "cover_mode": "centered", "header_mode": "split",   "footer_mode": "line",
        "card_shift_x": 0.12, "card_shift_y": 0.06, "card_shrink_w": 0.08,
    },
}

# ==========================================================================
# PROFILE PALETTES — each profile has its own distinct color family + fonts
# classic=blue, magazine=terracotta, executive=forest, tech=berry,
# bold=navy+coral, minimal=charcoal
# ==========================================================================

PROFILE_PALETTES = {
    "classic": {
        "primary":     (11,  95, 255),
        "secondary":   (77, 163, 255),
        "accent":      (0,   58, 160),
        "accent2":     (0,  163, 163),
        "bg_dark":     (11,  42,  74),
        "text_muted":  (100, 130, 160),
        "header_font": "Calibri",
        "body_font":   "Calibri",
    },
    "magazine": {
        "primary":     (184,  80,  66),
        "secondary":   (231, 194,  89),
        "accent":      (120,  45,  35),
        "accent2":     (167, 190, 174),
        "bg_dark":     ( 74,  28,  42),
        "text_muted":  (140, 100,  90),
        "header_font": "Georgia",
        "body_font":   "Calibri",
    },
    "executive": {
        "primary":     ( 44,  95,  45),
        "secondary":   (151, 188,  98),
        "accent":      ( 20,  55,  20),
        "accent2":     ( 90, 175, 150),
        "bg_dark":     ( 22,  55,  22),
        "text_muted":  ( 90, 120,  90),
        "header_font": "Trebuchet MS",
        "body_font":   "Calibri",
    },
    "tech": {
        "primary":     (109,  46,  70),
        "secondary":   (162, 103, 105),
        "accent":      ( 70,  22,  40),
        "accent2":     (236, 226, 208),
        "bg_dark":     ( 44,  18,  28),
        "text_muted":  (130,  90, 100),
        "header_font": "Consolas",
        "body_font":   "Calibri",
    },
    "bold": {
        "primary":     ( 20,  20,  90),
        "secondary":   (240,  80,  60),
        "accent":      ( 10,  10,  55),
        "accent2":     (255, 200,   0),
        "bg_dark":     ( 10,  10,  45),
        "text_muted":  (110, 110, 150),
        "header_font": "Arial Black",
        "body_font":   "Arial",
    },
    "minimal": {
        "primary":     ( 54,  69,  79),
        "secondary":   (120, 145, 160),
        "accent":      ( 28,  42,  50),
        "accent2":     (  0, 164, 180),
        "bg_dark":     ( 22,  35,  45),
        "text_muted":  (110, 130, 140),
        "header_font": "Calibri Light",
        "body_font":   "Calibri",
    },
}


# ==========================================================================
# LOW-LEVEL HELPERS
# ==========================================================================

def _c(t):
    return RGBColor(*t)

def _IN(v):
    return Inches(v)

def _cfg(key):
    return DECK_CFG[key]

def _mix(c1, c2, t):
    t = max(0.0, min(1.0, float(t)))
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))

def _pick_deck_profile() -> str:
    """Truly random profile selection — different visual style every generation."""
    return random.choice(list(DECK_PROFILE_CFG.keys()))

def _profile(style: dict):
    style = style if isinstance(style, dict) else {}
    name = str(style.get("_deck_profile", "classic"))
    return DECK_PROFILE_CFG.get(name, DECK_PROFILE_CFG["classic"])

def _coerce_int(value, default=50, min_value=0, max_value=100):
    if isinstance(value, int):
        n = value
    else:
        m = re.search(r"-?\d+", str(value))
        if not m:
            return default
        n = int(m.group(0))
    return max(min_value, min(max_value, n))

_SAFE_ICON_SET = {"▸", "•", "◆", "◼", "◻", "▲", "▼", "▶", "→", "✓", "✔", "★", "☆", "+", "-"}

def _safe_icon(value, default="▸"):
    s = str(value or "").strip()
    if not s:
        return default
    if s in _SAFE_ICON_SET:
        return s
    if len(s) == 1 and s.isascii() and (s.isalnum() or s in "!@#$%^&*()_+=-<>?/|~"):
        return s
    return default

def _title_size(text: str, base: int = None) -> int:
    if base is None:
        base = _cfg("title_base_pt")
    n = len(str(text or "").strip())
    for threshold, delta in _cfg("title_size_steps"):
        if n <= threshold:
            return max(20, base + delta)
    return max(20, base - 10)

def _subtitle_size(text: str, base: int = None) -> int:
    if base is None:
        base = _cfg("subtitle_base_pt")
    n = len(str(text or "").strip())
    for threshold, delta in _cfg("subtitle_size_steps"):
        if n <= threshold:
            return max(10, base + delta)
    return max(10, base - 2)

def _parse_color(value, fallback):
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


# ==========================================================================
# THEME SYSTEM
# ==========================================================================

BASE_THEME = dict(
    primary    = (11,  95, 255),
    secondary  = (77, 163, 255),
    accent     = (0,   58, 160),
    accent2    = (0,  163, 163),
    bg_light   = (255, 255, 255),
    bg_dark    = (11,  42,  74),
    text_dark  = (30,  30,  30),
    text_light = (255, 255, 255),
    text_muted = (120, 140, 160),
    card_bg    = (255, 255, 255),
    header_font= "Calibri",
    body_font  = "Calibri",
)

def _normalize_theme(theme_payload):
    theme = BASE_THEME.copy()
    if not isinstance(theme_payload, dict):
        return theme
    for key in ("primary", "secondary", "accent", "accent2",
                "bg_light", "bg_dark", "text_dark", "text_light",
                "text_muted", "card_bg"):
        if key in theme_payload:
            theme[key] = _parse_color(theme_payload[key], theme[key])
    for key in ("header_font", "body_font"):
        v = theme_payload.get(key)
        if isinstance(v, str) and v.strip():
            theme[key] = v.strip()
    return theme

def _slide_theme(base_theme, style=None, slide_idx=1):
    style = style if isinstance(style, dict) else {}
    theme = dict(base_theme)

    for key in ("primary", "secondary", "accent", "accent2",
                "bg_light", "bg_dark", "text_dark", "text_light",
                "text_muted", "card_bg"):
        if key in style:
            theme[key] = _parse_color(style[key], theme[key])

    if str(style.get("accent_rotation", "")).lower() in ("auto", "rotate"):
        keys = _cfg("accent_rotation_keys")
        order = [theme[k] for k in keys]
        shift = (slide_idx - 1) % len(order)
        rot = order[shift:] + order[:shift]
        for k, v in zip(keys, rot):
            theme[k] = v

    surface = str(style.get("surface", "light")).lower()
    if surface == "dark":
        surface = "light"
    if surface == "tint":
        theme["bg_light"] = _mix(theme["bg_light"], theme["primary"], _cfg("tint_bg_mix"))
        theme["card_bg"]  = _mix(theme["card_bg"],  theme["secondary"], _cfg("tint_card_mix"))

    return theme


# ==========================================================================
# PRIMITIVE DRAWING
# ==========================================================================

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
             size=None, icon="▸", icon_color=None, text_color=None,
             face="Calibri", max_pts=None):
    if not points:
        return
    if size is None:
        size = _cfg("bullet_size_base")
    if max_pts is None:
        max_pts = _cfg("bullet_max_default")

    ic = icon_color or BASE_THEME["secondary"]
    tc = text_color  or BASE_THEME["text_dark"]
    icon = _safe_icon(icon, default="▸")
    pts = [str(p) for p in points[:max_pts]]
    count = len(pts)

    bx = slide.shapes.add_textbox(_IN(x), _IN(y), _IN(w), _IN(h))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if count <= 3 else MSO_ANCHOR.TOP

    if count <= 2:
        size = min(_cfg("bullet_sparse_2_size"), size + 5)
    elif count == 3:
        size = min(_cfg("bullet_sparse_3_size"), size + 3)
    elif count == 4:
        size = min(_cfg("bullet_sparse_4_size"), size + 1)

    first = True
    for pt in pts:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        sp = _cfg("bullet_space_sparse") if count <= 3 else _cfg("bullet_space_normal")
        p.space_before = Pt(sp)
        p.space_after  = Pt(sp)

        ir = p.add_run()
        ir.text = f"{icon}  "
        ir.font.size  = Pt(size - 1)
        ir.font.bold  = True
        ir.font.name  = face
        ir.font.color.rgb = _c(ic)

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
        W       = _cfg("slide_w")
        bw      = _cfg("badge_w")
        bh      = _cfg("badge_h")
        bx      = W - bw - _cfg("badge_margin_r")
        by      = _cfg("badge_margin_t")
        pad     = _cfg("badge_inner_pad")
        max_h   = bh - pad * 2
        max_w   = bw - pad * 2

        with Image.open(logo_path) as img:
            iw, ih = img.size
        ratio = iw / ih if ih else 1.0
        lh    = min(max_h, max_w / ratio)
        lw    = lh * ratio
        ox    = (bw - lw) / 2
        oy    = (bh - lh) / 2

        _rect(slide, bx, by, bw, bh,
              (255, 255, 255),
              line=theme.get("secondary", BASE_THEME["secondary"]),
              lw=1.0)
        slide.shapes.add_picture(
            logo_path,
            left   = _IN(bx + ox),
            top    = _IN(by + oy),
            width  = _IN(lw),
            height = _IN(lh),
        )
    except Exception as e:
        print(f"[logo] {e}")


# ==========================================================================
# SHARED HEADER / FOOTER / CARD
# ==========================================================================

def _header(slide, title, subtitle, num, theme, style=None):
    style  = style if isinstance(style, dict) else {}
    W      = _cfg("slide_w")
    H_h    = _cfg("header_h")
    bw     = _cfg("badge_w")
    bmar_r = _cfg("badge_margin_r")
    bx_logo= W - bw - bmar_r
    nbw    = _cfg("num_badge_w")
    nbh    = _cfg("num_badge_h")
    nbt    = _cfg("num_badge_top")
    nbx    = bx_logo - _cfg("num_badge_gap")
    tx     = _cfg("header_title_x")
    ty     = _cfg("header_title_y")
    th     = _cfg("header_title_h")
    title_w = nbx - tx - 0.08

    prof    = _profile(style)
    variant = str(style.get("header_variant", "solid")).lower()
    mode    = prof["header_mode"]

    if mode == "sidebar":
        _rect(slide, 0, 0, 0.42, H_h, theme["primary"])
        _rect(slide, 0.42, 0, W - 0.42, H_h, _mix(theme["bg_dark"], theme["primary"], 0.18))
    elif mode == "band":
        _rect(slide, 0, 0, W, H_h, theme["bg_dark"])
        _rect(slide, 0, 0, W, 0.16, theme["secondary"])
    elif mode == "split" or variant == "split":
        _rect(slide, 0, 0, W * 0.58, H_h, theme["bg_dark"])
        _rect(slide, W * 0.58, 0, W * 0.42, H_h, theme["primary"])
    elif variant == "banded":
        _rect(slide, 0, 0, W, H_h, theme["bg_dark"])
        _rect(slide, 0, H_h - 0.18, W, 0.18, theme["secondary"])
    else:
        _rect(slide, 0, 0, W, H_h, theme["bg_dark"])

    t_size = _title_size(title, base=_cfg("title_base_pt"))
    _tb(slide, title, tx, ty, title_w, th, t_size,
        bold=True, color=theme["text_light"], face=theme["header_font"], shrink_to_fit=True)

    if subtitle:
        s_size    = _subtitle_size(subtitle, base=_cfg("subtitle_base_pt"))
        sub_color = _mix(theme["text_light"], theme["secondary"], _cfg("sub_color_mix"))
        sy        = _cfg("header_sub_y")
        sh        = _cfg("header_sub_h")
        _tb(slide, subtitle, tx, sy, title_w, sh, s_size,
            italic=True, color=sub_color, face=theme["body_font"], shrink_to_fit=True)

    badge_shape = str(style.get("badge_shape", "oval")).lower()
    if badge_shape == "rect":
        _rect(slide, nbx, nbt, nbw, nbh, theme["accent"], line=theme["secondary"], lw=1.0)
    else:
        _oval(slide, nbx, nbt, nbw, nbh, theme["accent"])

    num_str  = str(num)
    badge_pt = _cfg("num_font_2digit") if len(num_str) > 1 else _cfg("num_font_1digit")
    _tb(slide, num_str, nbx + 0.02, nbt + 0.04, nbw - 0.04, nbh - 0.08, badge_pt,
        bold=True, color=theme["text_light"],
        face=theme["header_font"], align=PP_ALIGN.CENTER)


def _footer(slide, theme, style=None):
    style   = style if isinstance(style, dict) else {}
    W       = _cfg("slide_w")
    H       = _cfg("slide_h")
    fh      = _cfg("footer_h")
    prof    = _profile(style)
    variant = str(style.get("footer_variant", "solid")).lower()
    if prof["footer_mode"] == "line" or variant == "line":
        _rect(slide, 0, H - 0.07, W, 0.07, theme["secondary"])
    else:
        _rect(slide, 0, H - fh, W, fh, theme["bg_dark"])


def _bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _c(color)


def _card(slide, theme, style=None, x=None, y=None, w=None, h=None):
    style   = style if isinstance(style, dict) else {}
    W       = _cfg("slide_w")
    H_h     = _cfg("header_h")
    gap     = _cfg("header_content_gap")
    fh      = _cfg("footer_h")
    ml      = _cfg("card_margin_l")
    mr      = _cfg("card_margin_r")
    ip      = _cfg("card_inner_pad")
    ipv     = _cfg("card_inner_pad_v")
    ipb     = _cfg("card_inner_pad_b")

    prof = _profile(style)
    cx = x if x is not None else ml + prof["card_shift_x"]
    cy = y if y is not None else H_h + gap + prof["card_shift_y"]
    cw = w if w is not None else W - cx - mr - prof["card_shrink_w"]

    if h is None:
        slide_h = _cfg("slide_h")
        ch = slide_h - H_h - fh - 0.22
    else:
        ch = h

    variant = str(style.get("card_variant", "outline")).lower()
    if prof["header_mode"] == "sidebar":
        _rect(slide, cx - 0.12, cy, 0.08, ch, _mix(theme["primary"], theme["secondary"], 0.5))

    if variant == "soft":
        _rect(slide, cx, cy, cw, ch,
              _mix(theme["card_bg"], theme["secondary"], 0.12),
              line=_mix(theme["primary"], theme["secondary"], 0.5), lw=0.8)
    elif variant == "banded":
        _rect(slide, cx, cy, cw, ch, theme["card_bg"], line=theme["primary"], lw=1.0)
        _rect(slide, cx, cy, cw, 0.12, theme["secondary"])
    else:
        _rect(slide, cx, cy, cw, ch, theme["card_bg"], line=theme["primary"], lw=1.0)

    return cx + ip, cy + ipv, cw - ip * 2, ch - ipv - ipb


# ==========================================================================
# LAYOUT RENDERERS
# ==========================================================================

def _render_bullets(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)
    _bullets(slide, spec.get("content", []),
             ix, iy, iw, ih,
             size=_cfg("bullet_size_base"),
             icon=spec.get("icon", "▸"),
             icon_color=stheme["secondary"],
             text_color=stheme["text_dark"],
             max_pts=_cfg("bullet_max_default"))
    _add_logo(slide, logo_path, stheme)


def _render_two_column(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    gap  = _cfg("two_col_gap")
    cw   = (iw - gap) / 2
    bsz  = _cfg("two_col_bullet_size")
    hh   = _cfg("two_col_header_h")

    left_title  = spec.get("left_title",  "Option A")
    right_title = spec.get("right_title", "Option B")
    left_pts    = spec.get("left_points",  [])
    right_pts   = spec.get("right_points", [])

    if not left_pts and not right_pts:
        content = spec.get("content", [])
        mid      = max(1, len(content) // 2)
        left_pts  = content[:mid]
        right_pts = content[mid:]

    icon = spec.get("icon", "▸")

    _rect(slide, ix, iy, cw, hh, stheme["primary"])
    _tb(slide, left_title, ix + 0.10, iy + 0.08, cw - 0.20, hh - 0.12, bsz,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])
    _bullets(slide, left_pts, ix + 0.10, iy + hh + 0.12, cw - 0.20, ih - hh - 0.18,
             size=bsz, icon=icon,
             icon_color=stheme["secondary"], text_color=stheme["text_dark"],
             max_pts=_cfg("two_col_min_bullets") + 1)

    rx = ix + cw + gap
    _rect(slide, rx, iy, cw, hh, stheme["accent2"])
    _tb(slide, right_title, rx + 0.10, iy + 0.08, cw - 0.20, hh - 0.12, bsz,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])
    _bullets(slide, right_pts, rx + 0.10, iy + hh + 0.12, cw - 0.20, ih - hh - 0.18,
             size=bsz, icon=icon,
             icon_color=stheme["accent"], text_color=stheme["text_dark"],
             max_pts=_cfg("two_col_min_bullets") + 1)

    _add_logo(slide, logo_path, stheme)


def _render_big_stat(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    pw     = _cfg("stat_panel_w")
    stat   = str(spec.get("stat", "—"))
    label  = spec.get("stat_label", "")
    source = spec.get("stat_source", "")
    pts    = spec.get("content", [])

    stat_font = _cfg("stat_font_short") if len(stat) <= 4 else _cfg("stat_font_long")

    _rect(slide, ix, iy, pw, ih, stheme["primary"])
    _rect(slide, ix, iy, pw, 0.07, stheme["accent2"])

    _tb(slide, stat, ix + 0.10, iy + 0.38, pw - 0.20, 1.80, stat_font,
        bold=True, color=stheme["text_light"],
        face=stheme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, label, ix + 0.10, iy + 2.25, pw - 0.20, 0.75, _cfg("stat_label_size"),
        color=stheme["secondary"], face=stheme["body_font"], align=PP_ALIGN.CENTER)
    if source:
        _tb(slide, source, ix + 0.10, iy + 3.05, pw - 0.20, 0.50, _cfg("stat_source_size"),
            italic=True, color=(180, 210, 255),
            face=stheme["body_font"], align=PP_ALIGN.CENTER)

    try:
        pct = float(stat.replace("%", "").replace("+", "").strip())
        if 0 < pct < 100:
            bmar  = _cfg("stat_bar_margin_x")
            bar_w = pw - bmar * 2
            bar_y = iy + ih - 0.55
            bh    = _cfg("stat_bar_h")
            _rect(slide, ix + bmar, bar_y, bar_w, bh, (255, 255, 255))
            _rect(slide, ix + bmar, bar_y, bar_w * (pct / 100), bh, stheme["accent2"])
    except ValueError:
        pass

    bx = ix + pw + 0.22
    bw = iw - pw - 0.22
    _bullets(slide, pts, bx, iy + 0.10, bw, ih - 0.20,
             size=_cfg("stat_bullet_size"),
             icon=spec.get("icon", "▸"),
             icon_color=stheme["secondary"],
             text_color=stheme["text_dark"],
             max_pts=_cfg("bullet_max_default"))

    _add_logo(slide, logo_path, stheme)


def _render_timeline(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
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
    DR    = _cfg("timeline_dot_outer")
    dr    = _cfg("timeline_dot_inner")
    rh    = _cfg("timeline_rail_h")
    lsz   = _cfg("timeline_label_size")
    dsz   = _cfg("timeline_detail_size")
    cw_   = _cfg("timeline_conn_w")

    accents = [stheme["primary"], stheme["accent2"], stheme["secondary"],
               stheme["accent"], stheme["primary"]]

    _rect(slide, ix, tl_y, iw, rh, stheme["secondary"])

    for i, step in enumerate(steps):
        cx  = ix + i * sw + sw / 2
        ac  = accents[i % len(accents)]

        _oval(slide, cx - DR, tl_y - DR, DR * 2, DR * 2, ac)
        _oval(slide, cx - dr, tl_y - dr, dr * 2, dr * 2, (255, 255, 255))
        _tb(slide, str(i + 1), cx - 0.20, tl_y - 0.16, 0.40, 0.32, lsz - 1,
            bold=True, color=ac, align=PP_ALIGN.CENTER)

        card_w = sw * 0.84
        cx_c   = cx - card_w / 2
        above  = (i % 2 == 0)

        if above:
            cy  = iy + 0.05
            ch  = max(0.40, tl_y - DR - 0.42 - cy)
            top = cy + ch
            bot = tl_y - DR
            if bot > top + 0.02:
                _rect(slide, cx - cw_ / 2, top, cw_, bot - top, stheme["secondary"])
        else:
            cy  = tl_y + DR + 0.28
            ch  = max(0.40, iy + ih - 0.05 - cy)
            top = tl_y + DR
            bot = cy
            if bot > top + 0.02:
                _rect(slide, cx - cw_ / 2, top, cw_, bot - top, stheme["secondary"])

        _rect(slide, cx_c, cy, card_w, ch, stheme["card_bg"], line=ac, lw=1.0)
        _tb(slide, step["label"], cx_c + 0.10, cy + 0.08, card_w - 0.20, 0.42, lsz,
            bold=True, color=ac, face=stheme["header_font"])
        dh = ch - 0.56
        if dh > 0.12:
            _tb(slide, step["detail"],
                cx_c + 0.10, cy + 0.52, card_w - 0.20, dh, dsz,
                color=stheme["text_dark"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)


def _render_icon_grid(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    items   = spec.get("grid_items", [])[:4]
    cols    = _cfg("icon_grid_cols")
    gap     = _cfg("icon_grid_gap")
    IR      = _cfg("icon_grid_radius")
    t_size  = _cfg("icon_grid_title_size")
    d_size  = _cfg("icon_grid_detail_size")
    bar_w   = _cfg("icon_grid_bar_w")
    cw_     = (iw - gap) / cols
    ch_     = (ih - gap) / 2
    accents = [stheme["primary"], stheme["accent2"], stheme["secondary"], stheme["accent"]]

    for idx, gi in enumerate(items):
        col = idx % cols
        row = idx // cols
        cx  = ix + col * (cw_ + gap)
        cy  = iy + row * (ch_ + gap)
        ac  = accents[idx % len(accents)]

        _rect(slide, cx, cy, cw_, ch_, stheme["card_bg"], line=ac, lw=1.0)
        _rect(slide, cx, cy, bar_w, ch_, ac)

        IOX = cx + 0.24
        IOY = cy + (ch_ - IR * 2) / 2
        _oval(slide, IOX, IOY, IR * 2, IR * 2, ac)
        raw_icon  = str(gi.get("icon", gi.get("title", "A"))).strip()
        first_char = next((c.upper() for c in raw_icon if c.isascii() and c.isalnum()), "A")
        _tb(slide, first_char, IOX + 0.04, IOY + 0.08, IR * 2 - 0.08, IR * 1.6, 17,
            bold=True, color=(255, 255, 255),
            face=stheme["header_font"], align=PP_ALIGN.CENTER)

        TX = IOX + IR * 2 + 0.16
        TW = cw_ - (IOX - cx) - IR * 2 - 0.22
        _tb(slide, str(gi.get("title", "")),
            TX, cy + 0.12, TW, 0.46, t_size,
            bold=True, color=stheme["primary"], face=stheme["header_font"])
        _tb(slide, str(gi.get("detail", "")),
            TX, cy + 0.60, TW, ch_ - 0.72, d_size,
            color=stheme["text_dark"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)


def _render_case_study(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    company  = spec.get("company", "Leading Organisation")
    result   = spec.get("result", "")
    metrics  = spec.get("metrics", [])
    bullets  = spec.get("content", [])
    icon     = spec.get("icon", "▸")

    bh   = _cfg("case_banner_h")
    bsz  = _cfg("case_banner_size")
    rh   = _cfg("case_ribbon_h")
    rsz  = _cfg("case_ribbon_size")
    mh   = _cfg("case_metric_row_h")
    mlw  = _cfg("case_metric_label_w")
    mlsz = _cfg("case_metric_label_size")
    mpsz = _cfg("case_metric_pct_size")
    cbsz = _cfg("case_bullet_size")

    _rect(slide, ix, iy, iw, bh, stheme["primary"])
    _tb(slide, f"  {company}",
        ix + 0.18, iy + 0.10, iw - 0.36, bh - 0.16, bsz,
        bold=True, color=stheme["text_light"], face=stheme["header_font"])

    if result:
        _rect(slide, ix, iy + bh + 0.06, iw, rh, stheme["accent2"])
        _tb(slide, f"  {result}",
            ix + 0.14, iy + bh + 0.14, iw - 0.28, rh - 0.14, rsz,
            bold=True, color=stheme["text_light"], face=stheme["body_font"])

    metrics_h = 0.0
    metric_start_y = iy + bh + (rh + 0.14 if result else 0.08)
    for i, m in enumerate(metrics[:3]):
        my    = metric_start_y + i * mh
        label = str(m.get("label", ""))
        value = min(100, max(0, int(m.get("value", 50))))
        bar_x = ix + mlw + 0.20
        bar_w = iw - mlw - 0.40
        _tb(slide, label, ix + 0.14, my + 0.04, mlw, 0.30, mlsz,
            color=stheme["text_dark"])
        _rect(slide, bar_x, my + 0.06, bar_w, 0.22, (225, 232, 245))
        _rect(slide, bar_x, my + 0.06, bar_w * (value / 100), 0.22, stheme["accent2"])
        _tb(slide, f"+{value}%",
            bar_x + bar_w * (value / 100) + 0.06, my, 0.55, 0.30,
            mpsz, bold=True, color=stheme["primary"])
        metrics_h = (i + 1) * mh

    bullet_y = metric_start_y + metrics_h + 0.12
    bullet_h = ih - (bullet_y - iy) - 0.10
    if bullet_h > 0.30 and bullets:
        _bullets(slide, bullets, ix + 0.10, bullet_y, iw - 0.20, bullet_h,
                 size=cbsz, icon=icon,
                 icon_color=stheme["secondary"],
                 text_color=stheme["text_dark"],
                 max_pts=_cfg("min_bullets_by_layout").get("case_study", 4))

    _add_logo(slide, logo_path, stheme)


def _render_table(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    columns  = spec.get("table_columns", [])
    rows     = spec.get("table_rows", [])
    max_cols = _cfg("table_max_cols")
    max_rows = _cfg("table_max_rows")
    hdr_h    = _cfg("table_header_h")
    hdr_sz   = _cfg("table_header_size")
    cell_sz  = _cfg("table_cell_size")
    col_gap  = _cfg("table_col_gap")
    pad      = _cfg("table_cell_pad")
    stripe   = _cfg("stripe_mix")

    if not columns or not rows:
        _tb(slide, "No table data available.", ix, iy, iw, ih, 14, color=stheme["text_muted"])
        _add_logo(slide, logo_path, stheme)
        return

    ncols = max(1, min(max_cols, len(columns)))
    cols  = [str(c) for c in columns[:ncols]]
    row_data = []
    for r in rows[:max_rows]:
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

    col_w = (iw - (ncols - 1) * col_gap) / ncols
    row_h = min(0.65, max(0.42, (ih - hdr_h - 0.08) / len(row_data)))

    for cidx, c in enumerate(cols):
        cx = ix + cidx * (col_w + col_gap)
        _rect(slide, cx, iy, col_w, hdr_h, stheme["primary"],
              line=stheme["secondary"], lw=0.8)
        _tb(slide, c, cx + pad, iy + pad, col_w - pad * 2, hdr_h - pad * 2, hdr_sz,
            bold=True, color=stheme["text_light"],
            face=stheme["header_font"], align=PP_ALIGN.CENTER, shrink_to_fit=True)

    for ridx, row in enumerate(row_data):
        ry   = iy + hdr_h + ridx * row_h
        fill = stheme["card_bg"] if ridx % 2 == 0 \
               else _mix(stheme["card_bg"], stheme["secondary"], 1 - stripe)
        for cidx, cell in enumerate(row):
            cx = ix + cidx * (col_w + col_gap)
            _rect(slide, cx, ry, col_w, row_h, fill,
                  line=_mix(stheme["primary"], stheme["card_bg"], stripe), lw=0.5)
            _tb(slide, cell, cx + pad, ry + pad, col_w - pad * 2, row_h - pad, cell_sz,
                color=stheme["text_dark"], face=stheme["body_font"],
                align=PP_ALIGN.LEFT, shrink_to_fit=True)

    _add_logo(slide, logo_path, stheme)


def _render_chart(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    chart_data = spec.get("chart_data", [])
    if not isinstance(chart_data, list):
        chart_data = []
    chart_data = [x for x in chart_data if isinstance(x, dict) and str(x.get("label", "")).strip()]
    if not chart_data:
        _tb(slide, "No chart data available.", ix, iy, iw, ih, 14, color=stheme["text_muted"])
        _add_logo(slide, logo_path, stheme)
        return

    items = chart_data[:_cfg("chart_max_items")]
    values = [max(1, _coerce_int(i.get("value", 0), default=1, min_value=1, max_value=100)) for i in items]
    vmax = max(values) if values else 100

    title_h = 0.48
    row_h = _cfg("chart_bar_h")
    row_gap = _cfg("chart_bar_gap")
    lsz = _cfg("chart_label_size")
    vsz = _cfg("chart_value_size")
    source_size = _cfg("chart_source_size")

    chart_top = iy + title_h + 0.06
    available_h = ih - title_h - 0.38
    needed_h = len(items) * row_h + (len(items) - 1) * row_gap
    if needed_h > available_h and len(items) > 1:
        scale = available_h / needed_h
        row_h *= scale
        row_gap *= scale

    _tb(slide, str(spec.get("chart_title", "Performance Snapshot")),
        ix + 0.02, iy + 0.02, iw - 0.04, title_h - 0.08, 15,
        bold=True, color=stheme["primary"], face=stheme["header_font"])

    label_w = max(1.95, iw * 0.28)
    bar_x = ix + label_w + 0.16
    bar_w = iw - label_w - 0.66
    for i, item in enumerate(items):
        y = chart_top + i * (row_h + row_gap)
        label = str(item.get("label", "Metric")).strip()[:40]
        val = max(1, _coerce_int(item.get("value", 0), default=1, min_value=1, max_value=100))
        pct = val / vmax if vmax else 0.0
        _tb(slide, label, ix + 0.02, y + 0.01, label_w - 0.08, row_h - 0.02, lsz,
            color=stheme["text_dark"], face=stheme["body_font"], shrink_to_fit=True)
        _rect(slide, bar_x, y, bar_w, row_h, _mix(stheme["card_bg"], stheme["secondary"], 0.90))
        _rect(slide, bar_x, y, bar_w * pct, row_h, stheme["accent2"])
        _tb(slide, f"{val}", bar_x + bar_w + 0.06, y + 0.01, 0.35, row_h - 0.02, vsz,
            bold=True, color=stheme["primary"], align=PP_ALIGN.RIGHT)

    source = str(spec.get("chart_source", "")).strip()
    if source:
        _tb(slide, source, ix + 0.02, iy + ih - 0.24, iw - 0.04, 0.18, source_size,
            italic=True, color=stheme["text_muted"], face=stheme["body_font"])

    _add_logo(slide, logo_path, stheme)


def _render_image_text_split(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    image_ratio = _cfg("split_image_ratio")
    gap = 0.20
    image_w = iw * image_ratio
    text_w = iw - image_w - gap
    image_side = str(spec.get("image_side", "")).strip().lower()
    if image_side not in {"left", "right"}:
        image_side = random.choice(["left", "right"])

    if image_side == "left":
        imx, tx = ix, ix + image_w + gap
    else:
        tx, imx = ix, ix + text_w + gap

    # Visual placeholder block that varies with theme and still looks intentional.
    _rect(slide, imx, iy, image_w, ih, _mix(stheme["secondary"], stheme["card_bg"], 0.75),
          line=stheme["secondary"], lw=1.0)
    _rect(slide, imx + 0.10, iy + 0.10, image_w - 0.20, ih * 0.20, stheme["primary"])
    _rect(slide, imx + 0.10, iy + ih * 0.34, image_w - 0.20, ih * 0.52,
          _mix(stheme["accent2"], stheme["card_bg"], 0.55))
    _oval(slide, imx + image_w * 0.10, iy + ih * 0.72, 0.45, 0.45, stheme["accent"])
    _oval(slide, imx + image_w * 0.78, iy + ih * 0.16, 0.36, 0.36, stheme["accent2"])

    caption = str(spec.get("image_caption", "")).strip()
    if caption:
        _tb(slide, caption, imx + 0.14, iy + ih - 0.48, image_w - 0.28, 0.38, _cfg("split_caption_size"),
            bold=True, color=stheme["text_dark"], face=stheme["body_font"], align=PP_ALIGN.CENTER)

    _bullets(slide, spec.get("content", []), tx, iy + 0.06, text_w, ih - 0.12,
             size=_cfg("split_bullet_size"),
             icon=spec.get("icon", "▸"),
             icon_color=stheme["secondary"],
             text_color=stheme["text_dark"],
             max_pts=_cfg("bullet_max_default"))

    _add_logo(slide, logo_path, stheme)


def _render_hybrid_insight(slide, spec, num, theme, logo_path):
    # Hybrid composition: KPI + mini bars + explanatory bullets.
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)
    _bg(slide, stheme["bg_light"])
    _header(slide, spec["title"], spec.get("subtitle", ""), num, stheme, style=style)
    _footer(slide, stheme, style=style)
    ix, iy, iw, ih = _card(slide, stheme, style=style)

    lw = iw * 0.43
    gap = 0.20
    rw = iw - lw - gap
    rx = ix + lw + gap

    stat = str(spec.get("stat", "")).strip() or "—"
    label = str(spec.get("stat_label", "")).strip() or "Key indicator"
    _rect(slide, ix, iy, lw, ih, _mix(stheme["card_bg"], stheme["secondary"], 0.86),
          line=stheme["secondary"], lw=1.0)
    _tb(slide, stat, ix + 0.12, iy + 0.30, lw - 0.24, 1.15, 52 if len(stat) <= 4 else 42,
        bold=True, color=stheme["primary"], face=stheme["header_font"], align=PP_ALIGN.CENTER)
    _tb(slide, label, ix + 0.12, iy + 1.45, lw - 0.24, 0.42, 13,
        color=stheme["text_dark"], face=stheme["body_font"], align=PP_ALIGN.CENTER)

    chart = spec.get("chart_data", [])
    if not isinstance(chart, list):
        chart = []
    rows = [r for r in chart if isinstance(r, dict)][:3]
    if not rows:
        rows = [{"label": "Current", "value": 52},
                {"label": "Target", "value": 71},
                {"label": "Potential", "value": 88}]
    vmax = max(max(1, _coerce_int(r.get("value", 1), default=1, min_value=1, max_value=100)) for r in rows)
    by = iy + 2.05
    row_h = 0.26
    row_gap = 0.16
    for i, row in enumerate(rows):
        y = by + i * (row_h + row_gap)
        v = max(1, _coerce_int(row.get("value", 1), default=1, min_value=1, max_value=100))
        _tb(slide, str(row.get("label", "Metric"))[:18], ix + 0.12, y - 0.01, 1.40, row_h, 11,
            color=stheme["text_dark"], face=stheme["body_font"], shrink_to_fit=True)
        bar_x = ix + 1.55
        bar_w = lw - 2.08
        _rect(slide, bar_x, y, bar_w, row_h, _mix(stheme["card_bg"], stheme["secondary"], 0.93))
        _rect(slide, bar_x, y, bar_w * (v / vmax), row_h, stheme["accent2"])
        _tb(slide, f"{v}", bar_x + bar_w + 0.03, y - 0.01, 0.26, row_h, 10,
            bold=True, color=stheme["primary"], align=PP_ALIGN.RIGHT)

    _rect(slide, rx, iy, rw, ih, stheme["card_bg"], line=stheme["primary"], lw=1.0)
    _bullets(slide, spec.get("content", []), rx + 0.10, iy + 0.10, rw - 0.20, ih - 0.20,
             size=14, icon=spec.get("icon", "▸"),
             icon_color=stheme["secondary"], text_color=stheme["text_dark"],
             max_pts=_cfg("bullet_max_default"))

    _add_logo(slide, logo_path, stheme)


def _render_title_cover(slide, spec, num, theme, logo_path):
    style  = spec.get("style", {})
    stheme = _slide_theme(theme, style, slide_idx=num)

    W  = _cfg("slide_w")
    H  = _cfg("slide_h")
    th = _cfg("cover_bar_top_h")
    tg = _cfg("cover_bar_top_gap")
    fh = _cfg("cover_footer_h")
    tx = _cfg("cover_title_x")
    ty = _cfg("cover_title_y")
    tw = W - _cfg("cover_title_pad")
    tH = _cfg("cover_title_h")

    prof = _profile(style)
    mode = prof["cover_mode"]

    if mode == "dark_full":
        # NEW: full dark background cover — premium look
        _bg(slide, stheme["bg_dark"])
        _rect(slide, 0, 0, 0.35, H, stheme["secondary"])
        _rect(slide, 0, H - 0.12, W, 0.12, stheme["accent2"])
        title_color  = stheme["text_light"]
        sub_color    = stheme["secondary"]
        bullet_color = stheme["text_light"]
        icon_color   = stheme["accent2"]
    elif mode == "sidebar":
        _bg(slide, stheme["bg_light"])
        _rect(slide, 0, 0, 2.05, H, stheme["primary"])
        _rect(slide, 2.05, 0, 0.10, H, stheme["secondary"])
        title_color  = stheme["primary"]
        sub_color    = stheme["accent"]
        bullet_color = stheme["text_dark"]
        icon_color   = stheme["secondary"]
    elif mode == "centered":
        _bg(slide, stheme["bg_light"])
        _rect(slide, 0, 0, W, th * 0.85, stheme["primary"])
        _oval(slide, W - 3.2, 0.8, 3.7, 3.7, _mix(stheme["secondary"], (255, 255, 255), 0.15))
        _rect(slide, 0, H - fh, W, fh, _mix(stheme["primary"], stheme["secondary"], 0.55))
        title_color  = stheme["primary"]
        sub_color    = stheme["accent"]
        bullet_color = stheme["text_dark"]
        icon_color   = stheme["secondary"]
    elif mode == "grid":
        _bg(slide, stheme["bg_light"])
        _rect(slide, 0, 0, W, th, stheme["primary"])
        _rect(slide, 0, th, W, tg, stheme["secondary"])
        for gx in (0.8, 2.6, 4.4, 6.2, 8.0, 9.8, 11.6):
            _rect(slide, gx, 0, 0.04, H, _mix(stheme["secondary"], stheme["bg_light"], 0.65))
        _rect(slide, 0, H - fh, W, fh, _mix(stheme["primary"], stheme["secondary"], 0.55))
        title_color  = stheme["primary"]
        sub_color    = stheme["accent"]
        bullet_color = stheme["text_dark"]
        icon_color   = stheme["secondary"]
    else:  # bands (classic)
        _bg(slide, stheme["bg_light"])
        _rect(slide, 0, 0, W, th, stheme["primary"])
        _rect(slide, 0, th, W, tg, stheme["secondary"])
        _rect(slide, 0, H - fh, W, fh, _mix(stheme["primary"], stheme["secondary"], 0.55))
        title_color  = stheme["primary"]
        sub_color    = stheme["accent"]
        bullet_color = stheme["text_dark"]
        icon_color   = stheme["secondary"]

    title    = spec.get("title", "Presentation")
    subtitle = spec.get("subtitle", "")
    t_size   = _title_size(title, base=_cfg("cover_title_base"))
    title_x  = tx if mode != "sidebar" else 2.40
    title_w  = tw if mode != "sidebar" else W - 3.00

    _tb(slide, title, title_x, ty, title_w, tH, t_size,
        bold=True, color=title_color,
        face=stheme["header_font"], align=PP_ALIGN.LEFT, shrink_to_fit=True)

    if subtitle:
        s_size = _subtitle_size(subtitle, base=_cfg("cover_subtitle_base"))
        sy     = _cfg("cover_sub_y")
        sh     = _cfg("cover_sub_h")
        _tb(slide, subtitle, title_x + 0.02, sy, title_w, sh, s_size,
            italic=True, color=sub_color,
            face=stheme["body_font"], shrink_to_fit=True)

    points = spec.get("content", [])
    if points:
        by   = _cfg("cover_bullet_y")
        bh_  = _cfg("cover_bullet_h")
        bsz  = _cfg("cover_bullet_size")
        bmax = _cfg("cover_bullet_max")
        _bullets(slide, points[:bmax], title_x + 0.04, by, title_w, bh_,
                 size=bsz,
                 icon=_safe_icon(spec.get("icon"), default="▸"),
                 icon_color=icon_color,
                 text_color=bullet_color,
                 max_pts=bmax)

    _add_logo(slide, logo_path, stheme)


# ==========================================================================
# LAYOUT REGISTRY
# ==========================================================================

_RENDERERS = {
    "title_cover": _render_title_cover,
    "bullets":     _render_bullets,
    "two_column":  _render_two_column,
    "big_stat":    _render_big_stat,
    "timeline":    _render_timeline,
    "icon_grid":   _render_icon_grid,
    "case_study":  _render_case_study,
    "table":       _render_table,
    "chart":       _render_chart,
    "image_text_split": _render_image_text_split,
    "hybrid_insight": _render_hybrid_insight,
}

_CONTENT_LAYOUTS = (
    "bullets", "two_column", "big_stat", "timeline", "icon_grid",
    "case_study", "table", "chart", "image_text_split", "hybrid_insight"
)
_LAYOUT_ALIASES = {
    "comparison": "two_column",
    "grid": "icon_grid",
    "infographic": "icon_grid",
    "kpi": "big_stat",
    "metric": "big_stat",
    "chart_slide": "chart",
    "graph": "chart",
    "bar_chart": "chart",
    "image_text": "image_text_split",
    "split": "image_text_split",
    "hybrid": "hybrid_insight",
    "storytelling": "hybrid_insight",
}
_BLUEPRINT_INTENTS = (
    "hook", "problem", "solution", "comparison", "evidence", "architecture",
    "roadmap", "risk", "use_case", "future", "summary"
)
_INTENT_LAYOUT_MAP = {
    "hook":         ("big_stat", "hybrid_insight", "image_text_split", "bullets"),
    "problem":      ("bullets", "two_column", "image_text_split"),
    "solution":     ("icon_grid", "image_text_split", "hybrid_insight", "two_column"),
    "comparison":   ("table", "two_column", "chart"),
    "evidence":     ("chart", "big_stat", "hybrid_insight", "case_study"),
    "architecture": ("image_text_split", "icon_grid", "hybrid_insight"),
    "roadmap":      ("timeline", "image_text_split", "bullets"),
    "risk":         ("two_column", "table", "bullets"),
    "use_case":     ("case_study", "image_text_split", "bullets"),
    "future":       ("icon_grid", "timeline", "bullets"),
    "summary":      ("bullets", "big_stat", "hybrid_insight"),
}
_YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
_NUMBER_RE = re.compile(r"\b\d+([.,]\d+)?\s*(%|x|k|m|b|bn|million|billion)?\b", re.IGNORECASE)


def _safe_lower_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def _canonical_layout(value, default="bullets") -> str:
    raw = _safe_lower_text(value)
    raw = _LAYOUT_ALIASES.get(raw, raw)
    return raw if raw in _CONTENT_LAYOUTS or raw == "title_cover" else default


def _parse_json_payload(raw: str):
    if not raw:
        return {}
    txt = raw.strip()
    if txt.startswith("```"):
        txt = txt.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        start = txt.find("{")
        end = txt.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return {}
        try:
            return json.loads(txt[start:end + 1])
        except json.JSONDecodeError:
            return {}


def _slide_text_blob(slide: dict, topic: str) -> str:
    parts = [topic, slide.get("title", ""), slide.get("subtitle", "")]
    for key in ("content", "left_points", "right_points"):
        val = slide.get(key)
        if isinstance(val, list):
            parts.extend([str(x) for x in val])
    steps = slide.get("steps", [])
    if isinstance(steps, list):
        for s in steps:
            if isinstance(s, dict):
                parts.append(str(s.get("label", "")))
                parts.append(str(s.get("detail", "")))
    for key in ("left_title", "right_title", "stat_label", "stat_source", "company"):
        parts.append(str(slide.get(key, "")))
    return " ".join(parts).lower()


def _extract_chart_data_from_content(content, max_items=5):
    if not isinstance(content, list):
        return []
    out = []
    for point in content:
        text = str(point).strip()
        if not text:
            continue
        num_match = re.search(r"(-?\d+(?:\.\d+)?)\s*(%|x)?", text)
        if not num_match:
            continue
        raw = float(num_match.group(1))
        value = int(max(1, min(100, round(raw if raw <= 100 else raw / 10))))
        # Create a short, readable label from the leading phrase.
        label = re.sub(r"\*\*", "", text)
        label = re.split(r"[:;,.\-]", label)[0].strip()
        words = label.split()
        label = " ".join(words[:4]) if words else "Metric"
        out.append({"label": label or "Metric", "value": value})
        if len(out) >= max_items:
            break
    return out


def _slide_signals(slide: dict, topic: str) -> dict:
    blob = _slide_text_blob(slide, topic)
    return {
        "timeline": bool(_YEAR_RE.search(blob)) or any(k in blob for k in
            ("timeline", "history", "evolution", "roadmap", "milestone", "phase", "journey")),
        "comparison": any(k in blob for k in
            ("compare", "comparison", "versus", " vs ", "before", "after", "pros", "cons", "trade-off")),
        "numbers": bool(_NUMBER_RE.search(blob)),
        "architecture": any(k in blob for k in
            ("architecture", "stack", "pipeline", "system design", "workflow", "layer", "component")),
        "risk": any(k in blob for k in
            ("risk", "challenge", "mitigation", "constraint", "limitation")),
        "use_case": any(k in blob for k in
            ("case study", "customer", "deployment", "implementation", "real-world", "example")),
        "future": any(k in blob for k in
            ("future", "next", "outlook", "trend", "road ahead")),
    }


def _count_years(text: str) -> int:
    return len({m.group(0) for m in _YEAR_RE.finditer(text or "")})


def _count_numeric_mentions(text: str) -> int:
    return len(_NUMBER_RE.findall(text))


def _layout_supported_by_content(slide: dict, layout: str, topic: str) -> bool:
    layout = _canonical_layout(layout, default="bullets")
    if layout in {"bullets", "icon_grid", "image_text_split"}:
        return True

    blob = _slide_text_blob(slide, topic)
    sig = _slide_signals(slide, topic)

    steps = slide.get("steps", [])
    if not isinstance(steps, list):
        steps = []
    has_valid_steps = len([s for s in steps if isinstance(s, dict) and s.get("label") and s.get("detail")]) >= 3

    metrics = slide.get("metrics", [])
    if not isinstance(metrics, list):
        metrics = []
    has_metrics = len([m for m in metrics if isinstance(m, dict)]) >= 2

    chart_data = slide.get("chart_data", [])
    if not isinstance(chart_data, list):
        chart_data = []
    has_chart_data = len([c for c in chart_data if isinstance(c, dict)]) >= 3

    table_cols = slide.get("table_columns", [])
    table_rows = slide.get("table_rows", [])
    if not isinstance(table_cols, list):
        table_cols = []
    if not isinstance(table_rows, list):
        table_rows = []

    left_pts = slide.get("left_points", [])
    right_pts = slide.get("right_points", [])
    if not isinstance(left_pts, list):
        left_pts = []
    if not isinstance(right_pts, list):
        right_pts = []

    stat = str(slide.get("stat", "")).strip()
    company = str(slide.get("company", "")).strip().lower()
    generic_case = company in _cfg("case_study_generic_companies")
    years = _count_years(blob)
    nums = _count_numeric_mentions(blob)

    if layout == "timeline":
        return has_valid_steps or years >= 2 or sig["timeline"]
    if layout == "case_study":
        return has_metrics or ("case study" in blob) or (company and not generic_case)
    if layout == "table":
        return (len(table_cols) >= 3 and len(table_rows) >= 3) or sig["comparison"]
    if layout == "two_column":
        return len(left_pts) >= 3 or len(right_pts) >= 3 or sig["comparison"] or sig["risk"]
    if layout == "chart":
        return has_chart_data or nums >= 3
    if layout == "big_stat":
        return bool(stat) and stat not in _cfg("stat_placeholder_values") or nums >= 1
    if layout == "hybrid_insight":
        return (nums >= 2 and (sig["architecture"] or sig["numbers"])) or has_chart_data
    return True


def _coerce_layout_by_evidence(slide: dict, layout: str, topic: str) -> str:
    layout = _canonical_layout(layout, default="bullets")
    if _layout_supported_by_content(slide, layout, topic):
        return layout
    fallback_map = {
        "timeline": ("image_text_split", "bullets"),
        "case_study": ("hybrid_insight", "bullets"),
        "table": ("two_column", "bullets"),
        "two_column": ("bullets", "image_text_split"),
        "chart": ("hybrid_insight", "bullets"),
        "big_stat": ("chart", "bullets"),
        "hybrid_insight": ("image_text_split", "bullets"),
    }
    for cand in fallback_map.get(layout, ("bullets",)):
        if _layout_supported_by_content(slide, cand, topic):
            return cand
    return "bullets"


def _infer_blueprint_intent(slide: dict, topic: str, idx: int, total: int) -> str:
    s = _slide_signals(slide, topic)
    if idx == total:
        return "summary"
    if s["architecture"]:
        return "architecture"
    if s["comparison"]:
        return "comparison"
    if s["timeline"]:
        return "roadmap"
    if s["use_case"]:
        return "use_case"
    if s["risk"]:
        return "risk"
    if s["numbers"]:
        return "evidence"
    if idx == 2:
        return random.choice(["hook", "problem"])
    if s["future"]:
        return "future"
    return random.choice(["solution", "evidence", "problem"])


def _pick_layout_for_intent(intent: str, previous_layout: str, used_counts: dict, caps: dict) -> str:
    candidates = list(_INTENT_LAYOUT_MAP.get(intent, ("bullets",)))
    random.shuffle(candidates)
    for c in candidates:
        if c == previous_layout:
            continue
        if used_counts.get(c, 0) >= caps.get(c, 1):
            continue
        return c
    for c in candidates:
        if c != previous_layout:
            return c
    return "bullets"


def _fallback_blueprint(topic: str, slides: list) -> list:
    # Build an intent-driven blueprint with optional sections (no forced timeline/case-study).
    if not isinstance(slides, list) or len(slides) <= 1:
        return []
    total = len(slides)
    caps = _layout_caps(total)
    used = {k: 0 for k in _CONTENT_LAYOUTS}
    prev = ""
    blueprint = []
    indices = list(range(2, total + 1))
    random.shuffle(indices)
    ordered = sorted(indices, key=lambda i: 0 if i in (2, total) else 1)
    for idx in ordered:
        slide = slides[idx - 1] if idx - 1 < len(slides) else {}
        intent = _infer_blueprint_intent(slide, topic, idx, total)
        layout = _pick_layout_for_intent(intent, prev, used, caps)
        used[layout] = used.get(layout, 0) + 1
        prev = layout
        blueprint.append({"source_index": idx, "intent": intent, "layout": layout})
    # Ensure actual slide order follows blueprint sequence.
    return blueprint


def _plan_blueprint_with_llm(topic: str, tone: str, slides: list) -> list:
    if not isinstance(slides, list) or len(slides) <= 1:
        return []
    compact = []
    for idx, s in enumerate(slides, start=1):
        if not isinstance(s, dict):
            s = {}
        sig = _slide_signals(s, topic)
        compact.append({
            "index": idx,
            "title": str(s.get("title", ""))[:110],
            "subtitle": str(s.get("subtitle", ""))[:130],
            "layout_hint": _canonical_layout(s.get("layout", "bullets")),
            "signals": sig,
            "sample_points": [str(x)[:140] for x in (s.get("content", []) if isinstance(s.get("content"), list) else [])[:2]],
        })
    prompt = f"""
You are a presentation blueprint architect.
Topic: "{topic}"
Tone: "{tone}"

Create a structurally diverse blueprint for slides 2..{len(slides)}.
Goal: avoid fixed patterns like always timeline/case-study/challenges.
Only include special sections when content signals justify them.

Allowed intents: {list(_BLUEPRINT_INTENTS)}
Allowed layouts: {list(_CONTENT_LAYOUTS)}

Return JSON only:
{{
  "blueprint": [
    {{"source_index": 2, "intent": "hook", "layout": "big_stat"}},
    ...
  ]
}}

Rules:
- Use each source_index from 2..{len(slides)} exactly once.
- You may reorder source_index to create a different narrative.
- Choose layout from intent and content signals.
- Keep adjacent layouts different when possible.
- Do NOT force timeline, case_study, or risk slides unless strongly supported.

Slide summaries:
{json.dumps(compact, ensure_ascii=True)}
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=1.0,
            max_tokens=1200,
        )
        parsed = _parse_json_payload((resp.choices[0].message.content or "").strip())
        bp = parsed.get("blueprint", []) if isinstance(parsed, dict) else []
        if not isinstance(bp, list):
            return []
        clean = []
        needed = set(range(2, len(slides) + 1))
        seen = set()
        for node in bp:
            if not isinstance(node, dict):
                continue
            src = _coerce_int(node.get("source_index", -1), default=-1, min_value=-1, max_value=1000)
            if src not in needed or src in seen:
                continue
            seen.add(src)
            intent = _safe_lower_text(node.get("intent", "solution"))
            if intent not in _BLUEPRINT_INTENTS:
                intent = _infer_blueprint_intent(slides[src - 1], topic, src, len(slides))
            layout = _canonical_layout(node.get("layout", "bullets"), default="bullets")
            clean.append({"source_index": src, "intent": intent, "layout": layout})
        if seen != needed:
            return []
        return clean
    except Exception:
        return []


def _reorder_and_stamp_by_blueprint(slides: list, topic: str, tone: str) -> list:
    if not isinstance(slides, list) or not slides:
        return []
    first = slides[0]
    rest = slides[1:]
    if not rest:
        return [first]
    llm_bp = _plan_blueprint_with_llm(topic=topic, tone=tone, slides=slides)
    bp = llm_bp if llm_bp else _fallback_blueprint(topic=topic, slides=slides)
    index_map = {i: s for i, s in enumerate(slides, start=1)}
    ordered = [first]
    prev_layout = ""
    for node in bp:
        s = index_map.get(node.get("source_index"), {})
        if not isinstance(s, dict):
            s = {}
        s = dict(s)
        layout = _canonical_layout(node.get("layout", s.get("layout", "bullets")), default="bullets")
        if layout == prev_layout:
            # force local diversity when blueprint repeats
            intent = _safe_lower_text(node.get("intent", "solution"))
            layout = _pick_layout_for_intent(intent, prev_layout, used_counts={}, caps=_layout_caps(len(slides)))
        s["layout"] = layout
        s["_blueprint_intent"] = _safe_lower_text(node.get("intent", "solution"))
        ordered.append(s)
        prev_layout = layout
    return ordered


def _layout_caps(num_slides: int) -> dict:
    # Keep heavy/specialized layouts rare and content-driven.
    core = max(1, num_slides - 1)
    return {
        "bullets": max(1, core),
        "two_column": 1 if core <= 5 else 2,
        "big_stat": 1,
        "timeline": 1,
        "icon_grid": 1 if core <= 6 else 2,
        "case_study": 1,
        "table": 1,
        "chart": 1 if core <= 6 else 2,
        "image_text_split": 1 if core <= 6 else 2,
        "hybrid_insight": 1 if core <= 6 else 2,
    }


def _choose_layout_for_slide(slide: dict, topic: str, idx: int, total: int,
                             prev_layout: str, used_counts: dict, caps: dict) -> str:
    blob = _slide_text_blob(slide, topic)
    requested = _canonical_layout(slide.get("layout", "bullets"), default="bullets")

    has_steps = isinstance(slide.get("steps"), list) and len(slide.get("steps")) >= 3
    has_table = isinstance(slide.get("table_columns"), list) and isinstance(slide.get("table_rows"), list) \
        and len(slide.get("table_columns")) >= 3 and len(slide.get("table_rows")) >= 2
    has_two_col = (isinstance(slide.get("left_points"), list) and len(slide.get("left_points")) >= 3) \
        or (isinstance(slide.get("right_points"), list) and len(slide.get("right_points")) >= 3)
    has_grid = isinstance(slide.get("grid_items"), list) and len(slide.get("grid_items")) >= 3
    has_metrics = isinstance(slide.get("metrics"), list) and len(slide.get("metrics")) >= 2
    has_stat = bool(str(slide.get("stat", "")).strip()) and str(slide.get("stat", "")).strip() != "—"
    has_chart = isinstance(slide.get("chart_data"), list) and len(slide.get("chart_data")) >= 3

    is_timeline = has_steps or bool(_YEAR_RE.search(blob)) or any(
        k in blob for k in ("timeline", "history", "evolution", "roadmap", "milestone", "journey", "phase")
    )
    is_table = has_table or any(
        k in blob for k in ("comparison", "compare", "benchmark", "matrix", "pricing", "options", "feature")
    ) or " vs " in blob
    is_two_col = has_two_col or any(
        k in blob for k in ("before after", "pros and cons", "pros", "cons", "challenges", "solutions",
                            "current state", "target state", "risks", "mitigation")
    )
    is_big_stat = has_stat or bool(_NUMBER_RE.search(blob)) and any(
        k in blob for k in ("growth", "reduction", "roi", "increase", "decrease", "adoption", "efficiency")
    )
    is_chart = has_chart or bool(_NUMBER_RE.search(blob)) and any(
        k in blob for k in ("distribution", "breakdown", "share", "segment", "performance", "trend", "quarter")
    )
    is_case = has_metrics or bool(str(slide.get("company", "")).strip()) or any(
        k in blob for k in ("case study", "customer story", "success story", "deployment", "implementation")
    )
    is_grid = has_grid or any(
        k in blob for k in ("pillars", "components", "capabilities", "framework", "building blocks", "modules")
    )
    is_split = any(
        k in blob for k in ("architecture", "workflow", "stack", "pipeline", "system design", "application")
    )
    is_hybrid = (is_chart and is_split) or (is_big_stat and is_split)

    if idx == total and any(k in blob for k in ("q&a", "questions", "thank you", "takeaway", "next steps")):
        candidates = ["bullets"]
    else:
        semantic_candidates = []
        if is_timeline:
            semantic_candidates.append("timeline")
        if is_table:
            semantic_candidates.append("table")
        if is_two_col:
            semantic_candidates.append("two_column")
        if is_case:
            semantic_candidates.append("case_study")
        if is_big_stat:
            semantic_candidates.append("big_stat")
        if is_chart:
            semantic_candidates.append("chart")
        if is_grid:
            semantic_candidates.append("icon_grid")
        if is_split:
            semantic_candidates.append("image_text_split")
        if is_hybrid:
            semantic_candidates.append("hybrid_insight")
        random.shuffle(semantic_candidates)
        candidates = semantic_candidates + [requested, "bullets"]

    seen = set()
    ordered = []
    for c in candidates:
        if c in _CONTENT_LAYOUTS and c not in seen:
            seen.add(c)
            ordered.append(c)

    for c in ordered:
        if used_counts.get(c, 0) >= caps.get(c, 1):
            continue
        if c == prev_layout and len(ordered) > 1:
            continue
        return c

    for c in ("two_column", "icon_grid", "bullets"):
        if used_counts.get(c, 0) < caps.get(c, 1) and c != prev_layout:
            return c
    return "bullets"


def _plan_layouts_with_llm(topic: str, tone: str, slides: list) -> list:
    if not isinstance(slides, list) or len(slides) <= 1:
        return []

    compact = []
    for idx, s in enumerate(slides, start=1):
        if not isinstance(s, dict):
            s = {}
        content = s.get("content", [])
        if not isinstance(content, list):
            content = []
        compact.append({
            "index": idx,
            "title": str(s.get("title", ""))[:120],
            "subtitle": str(s.get("subtitle", ""))[:140],
            "layout_hint": _canonical_layout(s.get("layout", "bullets")),
            "has_steps_data": isinstance(s.get("steps"), list) and len(s.get("steps")) >= 3,
            "has_table_data": isinstance(s.get("table_columns"), list) and isinstance(s.get("table_rows"), list),
            "has_metrics_data": isinstance(s.get("metrics"), list) and len(s.get("metrics")) >= 2,
            "has_stat_data": bool(str(s.get("stat", "")).strip()),
            "has_chart_data": isinstance(s.get("chart_data"), list) and len(s.get("chart_data")) >= 3,
            "sample_points": [str(x)[:160] for x in content[:2]],
        })

    prompt = f"""
You are a presentation layout planner.
Topic: "{topic}"
Tone: "{tone}"

Task:
Choose the best layout for each slide from 2..{len(slides)} based on semantic intent.
Do NOT follow any fixed order pattern.
Do NOT default slide 2 to timeline unless the content clearly indicates chronology/history/roadmap.
Across different topics, structure must differ naturally.

Allowed layouts:
{list(_CONTENT_LAYOUTS)}

Selection rules:
- timeline: only for chronology, milestones, phases, evolution, roadmap
- table: only for true comparison matrix/data rows
- two_column: for contrasts (before/after, pros/cons, challenges/solutions)
- big_stat: only when one key metric deserves emphasis
- case_study: for company/example outcomes
- icon_grid: for grouped capability pillars/components
- chart: for multiple numeric datapoints or performance breakdown
- image_text_split: for architecture/workflow/application explanations
- hybrid_insight: for combined KPI + mini-data + explanation on one slide
- bullets: default when none of the above strongly fits
- avoid repeating the same layout in adjacent slides when possible

Input slide summaries:
{json.dumps(compact, ensure_ascii=True)}

Return ONLY valid JSON:
{{
  "layouts": ["layout_for_slide_2", "layout_for_slide_3", "..."]
}}
Exactly {len(slides)-1} items in order.
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.9,
            max_tokens=900,
        )
        raw = (resp.choices[0].message.content or "").strip()
        parsed = _parse_json_payload(raw)
        layouts = parsed.get("layouts", []) if isinstance(parsed, dict) else []
        if not isinstance(layouts, list):
            return []
        cleaned = []
        for x in layouts[:len(slides)-1]:
            cleaned.append(_canonical_layout(x, default="bullets"))
        if len(cleaned) != len(slides) - 1:
            return []
        return cleaned
    except Exception:
        return []


# ==========================================================================
# LLM CONTENT GENERATION
# ==========================================================================

def generate_slide_content(topic: str, num_slides: int = 5,
                            tone: str = "Professional") -> dict:
    prompt = f"""
You are an expert presentation designer. Create a {tone} PowerPoint on "{topic}".
Return ONLY valid JSON — no markdown fences, no preamble, no explanation.

=== STRICT RULES ===
1. Generate exactly {num_slides} slides inside a top-level "slides" array.
2. Every slide MUST include: "title", "subtitle", "layout", "icon", "content", and "style".
3. Slide 1 MUST use layout "title_cover" with a big, clear title.
4. For slides 2..N, choose structure dynamically from topic needs; do NOT follow a fixed template order.
5. VARY layouts — do NOT use the same layout twice in a row.
6. VARY visual pattern — do NOT use the same `style.pattern_name` twice in a row.
7. Use table/timeline only when content genuinely benefits from them.
8. Use a clean LIGHT theme only (white and blue/teal family). Avoid dark backgrounds.
9. "content" bullets: 4-7 items, 18-34 words each, use **bold** for key terms.
   Every bullet must cite a REAL company, statistic, or year — never say "many companies".
10. Final slide should be context-appropriate (Thank You, Q&A, Key Takeaways) chosen by topic narrative.
11. Timeline and case-study slides are OPTIONAL; include them only if the topic/sections clearly require them.

=== DESIGN SYSTEM ===
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
  }}
}}

Each slide.style must include:
- pattern_name: unique descriptive label
- surface: one of ["light","tint"]
- header_variant: one of ["solid","split","banded"]
- card_variant: one of ["outline","soft","banded"]
- footer_variant: one of ["solid","line"]
- badge_shape: one of ["oval","rect"]
- accent_rotation: one of ["static","auto"]

=== LAYOUT RULES ===

"title_cover"
  title: large cover title
  subtitle: concise context line
  content: optional 2-3 short highlights

"bullets"
  content: [4-7 bullets]

"two_column"
  left_title:  descriptive string (NOT "Left")
  right_title: descriptive string (NOT "Right")
  left_points:  [4-5 bullets]
  right_points: [4-5 bullets]
  content: []

"big_stat"
  stat:        real value e.g. "40%", "$2.5B", "3x" (NOT "100%")
  stat_label:  <=8 words describing the stat
  stat_source: source and year
  content:     [5-6 supporting bullets]

"timeline"
  steps: [exactly 4 items]
    each: {{"label": "<=4 words", "detail": "15-20 word sentence"}}
  content: []

"icon_grid"
  grid_items: [exactly 4 items]
    each: {{"icon": "keyword", "title": "2-3 words", "detail": "12-18 words"}}
  content: []

"case_study"
  company: real company name (NOT "Case Study" or "Company")
  result:  one-sentence headline outcome
  metrics: [2-3 items, each: {{"label":"metric name","value":integer 1-99}}]
  content: [4-5 supporting bullets]

"table"
  table_columns: [3-5 short column headers]
  table_rows:    [4-6 rows, each matching column count]
  content: []

"chart"
  chart_title: short chart heading
  chart_data: [3-5 items, each: {{"label":"name","value":integer 1-100}}]
  chart_source: optional source/year line
  content: []

"image_text_split"
  image_caption: short visual caption
  image_side: one of ["left","right"]
  content: [4-6 bullets]

"hybrid_insight"
  stat: short KPI value
  stat_label: <=8 words
  chart_data: [3 items, each: {{"label":"name","value":integer 1-100}}]
  content: [4-6 bullets]

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
    }}
  }},
  "slides": [
    {{
      "title": "...",
      "subtitle": "...",
      "layout": "title_cover",
      "icon": "▸",
      "content": ["...", "...", "..."],
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
    data = _parse_json_payload(raw)

    if not isinstance(data, dict):
        data = {"design_system": {}, "slides": []}
    if not isinstance(data.get("design_system"), dict):
        data["design_system"] = {}
    if not isinstance(data.get("slides"), list):
        data["slides"] = []

    allowed_surfaces         = set(_cfg("allowed_surfaces"))
    allowed_header_variants  = set(_cfg("allowed_header_variants"))
    allowed_card_variants    = set(_cfg("allowed_card_variants"))
    allowed_footer_variants  = set(_cfg("allowed_footer_variants"))
    allowed_badge_shapes     = set(_cfg("allowed_badge_shapes"))
    allowed_accent_rotations = set(_cfg("allowed_accent_rotations"))
    min_bullets              = _cfg("min_bullets_by_layout")
    two_col_generic          = _cfg("two_col_generic_headers")
    case_generic             = _cfg("case_study_generic_companies")
    stat_placeholders        = _cfg("stat_placeholder_values")
    fallback_marker          = _cfg("fallback_bullet_marker")
    two_col_min              = _cfg("two_col_min_bullets")

    source_slides = [s for s in data.get("slides", []) if isinstance(s, dict)]
    source_slides = _reorder_and_stamp_by_blueprint(source_slides, topic=topic, tone=tone)
    llm_layout_plan = _plan_layouts_with_llm(topic=topic, tone=tone, slides=source_slides)

    clean_slides = []
    layout_counts = {k: 0 for k in _CONTENT_LAYOUTS}
    caps = _layout_caps(max(1, len(source_slides)))
    prev_content_layout = ""
    for idx, slide in enumerate(source_slides, start=1):
        if not isinstance(slide, dict):
            continue

        slide.setdefault("title", f"Slide {idx}")
        slide.setdefault("subtitle", "")
        layout = _canonical_layout(slide.get("layout", "bullets"))
        slide["icon"] = _safe_icon(slide.get("icon"), default="▸")
        slide.setdefault("content", [])
        if not isinstance(slide.get("style"), dict):
            slide["style"] = {}

        style = slide["style"]
        style.setdefault("pattern_name", f"pattern_{idx}")
        style.setdefault("surface",          "light")
        style.setdefault("header_variant",   "solid")
        style.setdefault("card_variant",     "outline")
        style.setdefault("footer_variant",   "solid")
        style.setdefault("badge_shape",      "oval")
        style.setdefault("accent_rotation",  "static")

        if style["surface"]          not in allowed_surfaces:         style["surface"]          = "light"
        if style["header_variant"]   not in allowed_header_variants:  style["header_variant"]   = "solid"
        if style["card_variant"]     not in allowed_card_variants:    style["card_variant"]     = "outline"
        if style["footer_variant"]   not in allowed_footer_variants:  style["footer_variant"]   = "solid"
        if style["badge_shape"]      not in allowed_badge_shapes:     style["badge_shape"]      = "oval"
        if style["accent_rotation"]  not in allowed_accent_rotations: style["accent_rotation"]  = "static"

        for key in ("content", "left_points", "right_points"):
            if key in slide and isinstance(slide[key], list):
                slide[key] = [
                    str(x.get("text", x)) if isinstance(x, dict) else str(x)
                    for x in slide[key]
                ]
        if not isinstance(slide.get("content"), list):
            slide["content"] = []
        slide["content"] = [str(x).strip() for x in slide["content"] if str(x).strip()]

        if idx == 1:
            layout = "title_cover"
        else:
            blueprint_choice = _canonical_layout(slide.get("layout", "bullets"))
            llm_choice = ""
            if len(llm_layout_plan) >= idx - 1:
                llm_choice = _canonical_layout(llm_layout_plan[idx - 2])
            if blueprint_choice in _CONTENT_LAYOUTS and blueprint_choice != prev_content_layout:
                layout = blueprint_choice
            elif llm_choice in _CONTENT_LAYOUTS and llm_choice != prev_content_layout:
                layout = llm_choice
            else:
                layout = _choose_layout_for_slide(
                    slide=slide,
                    topic=topic,
                    idx=idx,
                    total=len(source_slides),
                    prev_layout=prev_content_layout,
                    used_counts=layout_counts,
                    caps=caps,
                )
            layout = _coerce_layout_by_evidence(slide, layout, topic)
            layout_counts[layout] = layout_counts.get(layout, 0) + 1
            prev_content_layout = layout
        slide["layout"] = layout

        if layout == "two_column":
            for k, default in (("left_title", "Before"), ("right_title", "After")):
                v = slide.get(k, "").strip().lower()
                if not v or v in two_col_generic:
                    slide[k] = default
            if not slide.get("left_points") and not slide.get("right_points"):
                content = slide.get("content", [])
                mid = max(1, len(content) // 2)
                slide["left_points"]  = content[:mid]
                slide["right_points"] = content[mid:]
            for side in ("left_points", "right_points"):
                if not isinstance(slide.get(side), list):
                    slide[side] = []
                slide[side] = [str(x).strip() for x in slide[side] if str(x).strip()]
                while len(slide[side]) < two_col_min:
                    slide[side].append(
                        f"**Operational step**: define owner, timeline, and KPI to move this "
                        f"{topic} track from concept to production impact."
                    )

        if layout == "big_stat":
            stat = str(slide.get("stat", "")).strip()
            if not stat or stat in stat_placeholders:
                slide["stat"] = "—"

        if layout == "timeline":
            clean = [
                s for s in slide.get("steps", [])
                if isinstance(s, dict) and s.get("label") and s.get("detail")
            ]
            if not clean:
                # Do not force synthetic timeline placeholders; downgrade to content slide.
                layout = "bullets"
                slide["layout"] = layout
            else:
                slide["steps"] = clean
                for s in slide["steps"]:
                    s["label"]  = str(s.get("label", ""))
                    s["detail"] = str(s.get("detail", ""))

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

        if layout == "case_study":
            co = slide.get("company", "").strip().lower()
            if not co or co in case_generic:
                t = str(slide.get("title", ""))
                if ":" in t:
                    slide["company"] = t.split(":", 1)[1].strip() or "Featured Organisation"
                else:
                    slide["company"] = "Featured Organisation"
            metrics = slide.get("metrics", [])
            if not isinstance(metrics, list):
                metrics = []
            fixed_metrics = []
            for m in metrics[:3]:
                if not isinstance(m, dict):
                    m = {"label": "Impact", "value": m}
                fixed_metrics.append({
                    "label": str(m.get("label", "Impact")),
                    "value": _coerce_int(m.get("value", 50),
                                         default=50, min_value=1, max_value=99),
                })
            slide["metrics"] = fixed_metrics

        if layout == "table":
            cols = slide.get("table_columns", [])
            rows = slide.get("table_rows", [])
            if not isinstance(cols, list):
                cols = []
            cols = [str(c).strip() for c in cols if str(c).strip()][:_cfg("table_max_cols")]
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
                clean_rows.append(
                    [f"Workstream", f"Baseline for {topic}",
                     "KPI-driven improvement plan"][:len(cols)]
                    + [""] * max(0, len(cols) - 3)
                )
            slide["table_columns"] = cols
            slide["table_rows"]    = clean_rows[:_cfg("table_max_rows")]

        if layout == "chart":
            chart_data = slide.get("chart_data", [])
            if not isinstance(chart_data, list):
                chart_data = []
            fixed = []
            for item in chart_data[:_cfg("chart_max_items")]:
                if not isinstance(item, dict):
                    continue
                label = str(item.get("label", "Metric")).strip()[:40]
                value = _coerce_int(item.get("value", 50), default=50, min_value=1, max_value=100)
                if label:
                    fixed.append({"label": label, "value": value})
            if not fixed:
                fixed = _extract_chart_data_from_content(slide.get("content", []), max_items=_cfg("chart_max_items"))
            if len(fixed) < 3:
                fixed.extend([
                    {"label": "Current", "value": 55},
                    {"label": "Target", "value": 78},
                    {"label": "Potential", "value": 92},
                ][:max(0, 3 - len(fixed))])
            slide["chart_data"] = fixed[:_cfg("chart_max_items")]
            slide["chart_title"] = str(slide.get("chart_title", "")).strip() or "Performance Snapshot"

        if layout == "image_text_split":
            side = str(slide.get("image_side", "")).strip().lower()
            if side not in {"left", "right"}:
                slide["image_side"] = random.choice(["left", "right"])
            slide["image_caption"] = str(slide.get("image_caption", "")).strip() or "Visual Brief"

        if layout == "hybrid_insight":
            stat = str(slide.get("stat", "")).strip()
            if not stat or stat in stat_placeholders:
                extracted = _extract_chart_data_from_content(slide.get("content", []), max_items=1)
                stat = str(extracted[0]["value"]) if extracted else "65"
            slide["stat"] = stat
            slide["stat_label"] = str(slide.get("stat_label", "")).strip() or "Composite impact"

            chart_data = slide.get("chart_data", [])
            if not isinstance(chart_data, list):
                chart_data = []
            fixed = []
            for item in chart_data[:3]:
                if not isinstance(item, dict):
                    continue
                label = str(item.get("label", "Metric")).strip()[:24]
                value = _coerce_int(item.get("value", 50), default=50, min_value=1, max_value=100)
                if label:
                    fixed.append({"label": label, "value": value})
            if len(fixed) < 3:
                generated = _extract_chart_data_from_content(slide.get("content", []), max_items=3)
                for g in generated:
                    if len(fixed) >= 3:
                        break
                    if not any(g["label"] == x["label"] for x in fixed):
                        fixed.append(g)
            while len(fixed) < 3:
                defaults = [("Current", 48), ("Target", 70), ("Potential", 90)]
                label, value = defaults[len(fixed)]
                fixed.append({"label": label, "value": value})
            slide["chart_data"] = fixed[:3]

        min_count = min_bullets.get(layout, 0)
        while len(slide["content"]) < min_count:
            slide["content"].append(
                f"{fallback_marker}: add one measurable KPI, one owner, and one delivery date "
                f"so this {topic} initiative is actionable, auditable, and presentation-ready."
            )

        clean_slides.append(slide)

    # Prevent immediate repeated visual pattern
    for i in range(1, len(clean_slides)):
        cur  = clean_slides[i].get("style", {})
        prev = clean_slides[i - 1].get("style", {})
        if cur.get("pattern_name", "").strip().lower() == prev.get("pattern_name", "").strip().lower():
            cur["pattern_name"] = f'{cur.get("pattern_name", "pattern")}_{i+1}'
            cur["header_variant"] = "banded" if prev.get("header_variant") != "banded" else "split"

    if clean_slides:
        first = clean_slides[0]
        first["layout"] = "title_cover"
        if not first.get("title"):
            first["title"] = topic
        if not isinstance(first.get("content"), list):
            first["content"] = []
        first["content"] = [str(x).strip() for x in first["content"] if str(x).strip()][
            :_cfg("cover_bullet_max")]
    else:
        clean_slides = [{
            "title": topic,
            "subtitle": f"{tone} presentation",
            "layout": "title_cover",
            "icon": "▸",
            "content": [],
            "style": {
                "pattern_name": "cover_default",
                "surface": "light",
                "header_variant": "solid",
                "card_variant": "outline",
                "footer_variant": "line",
                "badge_shape": "oval",
                "accent_rotation": "static",
            },
        }]

    data["slides"] = clean_slides
    return data


# ==========================================================================
# PUBLIC API
# ==========================================================================

def create_ppt(slide_data, topic: str,
               logo_path: str = None,
               tone: str = "Professional") -> str:

    prs = Presentation()
    prs.slide_width  = _IN(_cfg("slide_w"))
    prs.slide_height = _IN(_cfg("slide_h"))

    payload       = slide_data if isinstance(slide_data, dict) else {"slides": slide_data}
    design_system = payload.get("design_system", {}) if isinstance(payload, dict) else {}

    # Start from LLM theme (parse what it gave us)
    theme = _normalize_theme(design_system.get("theme", {}))

    # Pick a RANDOM profile every time — this is the key fix
    deck_profile = _pick_deck_profile()

    # Apply the profile's distinct palette — completely overwrites the LLM theme's colors
    # so every profile looks genuinely different instead of all-blue
    pp = PROFILE_PALETTES.get(deck_profile, PROFILE_PALETTES["classic"])
    theme["primary"]     = pp["primary"]
    theme["secondary"]   = pp["secondary"]
    theme["accent"]      = pp["accent"]
    theme["accent2"]     = pp["accent2"]
    theme["bg_dark"]     = pp["bg_dark"]
    theme["text_muted"]  = pp["text_muted"]
    theme["header_font"] = pp["header_font"]
    theme["body_font"]   = pp["body_font"]

    # These stay constant across all profiles
    theme["bg_light"]   = (255, 255, 255)
    theme["card_bg"]    = (255, 255, 255)
    theme["text_dark"]  = (30, 30, 30)
    theme["text_light"] = (255, 255, 255)

    logo   = logo_path if (logo_path and os.path.exists(logo_path)) else None
    slides = payload.get("slides", [])
    if not isinstance(slides, list):
        slides = []

    for i, spec in enumerate(slides, start=1):
        if not isinstance(spec, dict):
            continue
        spec.setdefault("title", f"Slide {i}")
        spec.setdefault("subtitle", "")
        spec.setdefault("content", [])
        if not isinstance(spec.get("style"), dict):
            spec["style"] = {}
        spec["style"]["_deck_profile"] = deck_profile
        if spec["style"].get("surface") == "dark":
            spec["style"]["surface"] = "light"

        layout   = spec.get("layout", "bullets")
        renderer = _RENDERERS.get(layout, _render_bullets)
        slide    = prs.slides.add_slide(prs.slide_layouts[6])
        renderer(slide, spec, i, theme, logo)

    if len(prs.slides) == 0:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        fallback_marker = _cfg("fallback_bullet_marker")
        _render_bullets(
            slide,
            {
                "title":    topic,
                "subtitle": "Auto-generated summary",
                "layout":   "bullets",
                "icon":     "▸",
                "content":  [
                    f"{fallback_marker}: no valid slide plan was returned for {topic}.",
                    f"{fallback_marker}: regenerate with the same topic for a full deck.",
                    f"{fallback_marker}: add audience and use-case context for richer content.",
                    f"{fallback_marker}: service supports table, timeline, chart, split-view, case-study.",
                    f"{fallback_marker}: fallback prevents empty deck failures.",
                ],
                "style": {"surface": "light", "header_variant": "solid", "card_variant": "outline"},
            },
            1, theme, logo,
        )

    safe = topic.replace(" ", "_").replace("/", "_")[:60]
    os.makedirs("generated", exist_ok=True)
    path = f"generated/{safe}.pptx"
    prs.save(path)
    return path
