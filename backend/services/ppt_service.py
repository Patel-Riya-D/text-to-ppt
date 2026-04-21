"""
ppt_service.py
==============

Core presentation generation engine for the AI PPT system.

This module is responsible for:

* Generating structured slide content using LLMs
* Cleaning and normalizing AI-generated text
* Rendering PowerPoint presentations using python-pptx
* Applying visual themes, layouts, and design systems
* Ensuring output quality (no icons, markdown, or noise in slides)

──────────────────────────────────────────────────────────────

🔧 Key Responsibilities

1. Content Generation

   * Uses LLM to generate structured slide data (title, subtitle, content)
   * Supports multiple layouts (bullets, charts, timeline, icon grid, etc.)

2. Content Cleaning (Critical Layer)

   * Removes unwanted tokens such as:

     * icon names (check-circle, arrow, handshake, etc.)
     * markdown (**bold**, symbols)
     * emojis and special characters
   * Ensures ONLY clean, human-readable text is rendered in slides

3. Content Flattening

   * Converts nested JSON/dict/list structures into flat bullet lists
   * Handles inconsistent LLM outputs robustly

4. Rendering Engine

   * Uses python-pptx to generate slides
   * Supports:

     * headers, footers, page numbers
     * bullet layouts, grids, timelines, case studies
     * adaptive font sizing and spacing

5. Theming System

   * Predefined visual profiles (classic, tech, bold, etc.)
   * Dynamic color extraction from uploaded logos
   * Automatic contrast and readability adjustments

6. Final Safety Layer (Important)

   * Applies aggressive text sanitization before rendering
   * Guarantees no icon tokens or formatting artifacts appear in PPT

──────────────────────────────────────────────────────────────

⚠️ Design Principle

The system follows a strict pipeline:

```
LLM Output → Flatten → Clean → Validate → Render
```

No raw LLM text is ever rendered directly.

──────────────────────────────────────────────────────────────

🚀 Key Functions

* generate_slide_content()
  Generates structured slide JSON using LLM

* create_ppt()
  Converts cleaned slide data into a PowerPoint file

* _clean_bullet_text()
  Core sanitization function for removing unwanted tokens

* _validate_and_clean_content()
  Final gate before rendering

──────────────────────────────────────────────────────────────

🎯 Goal

To produce professional, clean, and visually appealing presentations
without any LLM artifacts such as:

* icon keywords
* markdown syntax
* noisy or malformed content

──────────────────────────────────────────────────────────────
"""

import os
import json
import ast
import re
import random
import colorsys
from PIL import Image
from openai import AzureOpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from io import BytesIO
import base64

from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

# ═══════════════════════════════════════════════════════════════════════════
#  SLIDE DIMENSIONS
# ═══════════════════════════════════════════════════════════════════════════
SW, SH = 13.3, 7.5
HEADER_H   = 1.40
FOOTER_H   = 0.28
CONTENT_Y  = HEADER_H + 0.14
CONTENT_H  = SH - HEADER_H - FOOTER_H - 0.22
CONTENT_X  = 0.42

_LOGO_PATH = None
CONTENT_W  = SW - CONTENT_X - 0.28

# ═══════════════════════════════════════════════════════════════════════════
#  PROFILE SYSTEM
# ═══════════════════════════════════════════════════════════════════════════

PROFILES = {
    "classic":   dict(cover="bands",     header="solid",   footer="bar",  card="outline"),
    "magazine":  dict(cover="sidebar",   header="sidebar", footer="line", card="soft"),
    "executive": dict(cover="centered",  header="band",    footer="bar",  card="outline"),
    "tech":      dict(cover="grid",      header="split",   footer="line", card="banded"),
    "bold":      dict(cover="dark_full", header="band",    footer="bar",  card="soft"),
    "minimal":   dict(cover="centered",  header="split",   footer="line", card="outline"),
}

PALETTES = {
    "classic":   dict(p=(11,95,255),   s=(77,163,255),  a=(0,58,160),   a2=(0,163,163),  dk=(11,42,74),   muted=(100,130,160), hf="Calibri",       bf="Calibri"),
    "magazine":  dict(p=(184,80,66),   s=(231,194,89),  a=(120,45,35),  a2=(167,190,174),dk=(74,28,42),   muted=(140,100,90),  hf="Georgia",        bf="Calibri"),
    "executive": dict(p=(44,95,45),    s=(151,188,98),  a=(20,55,20),   a2=(90,175,150), dk=(22,55,22),   muted=(90,120,90),   hf="Trebuchet MS",   bf="Calibri"),
    "tech":      dict(p=(109,46,70),   s=(162,103,105), a=(70,22,40),   a2=(236,226,208),dk=(44,18,28),   muted=(130,90,100),  hf="Consolas",       bf="Calibri"),
    "bold":      dict(p=(20,20,90),    s=(240,80,60),   a=(10,10,55),   a2=(255,200,0),  dk=(10,10,45),   muted=(110,110,150), hf="Arial Black",    bf="Arial"),
    "minimal":   dict(p=(54,69,79),    s=(120,145,160), a=(28,42,50),   a2=(0,164,180),  dk=(22,35,45),   muted=(110,130,140), hf="Calibri Light",  bf="Calibri"),
}

# ═══════════════════════════════════════════════════════════════════════════
#  ICON / MARKDOWN CLEANING  ← THE CORE FIX
# ═══════════════════════════════════════════════════════════════════════════

# Every token that must NEVER appear in slide text.
_ICON_TOKENS = {
    "check-circle", "check_circle", "checkcircle",
    "check-square", "check_square",
    "thank-you", "thank_you",
    "handshake",
    "trending-up", "trending_up", "trendingup",
    "trending-down", "trending_down",
    "arrow-right", "arrow_right", "arrowright",
    "arrow-left", "arrow_left",
    "arrow-up", "arrow_up",
    "arrow-down", "arrow_down",
    "arrow",
    "star", "heart", "flag", "circle", "check",
    "bullet", "info", "warning", "error", "success",
    "close", "done",
    "thumb-up", "thumb_up", "thumbup",
    "thumb-down", "thumb_down",
    "lightbulb", "rocket", "target", "shield",
    "chart-bar", "chart_bar", "chartbar",
    "pie-chart", "pie_chart", "piechart",
    "bar-chart", "bar_chart", "barchart",
    "line-chart", "line_chart",
    "users", "user", "person", "team",
    "building", "office", "company", "globe",
    "lock", "key", "settings", "gear",
    "plus", "minus", "cross", "tick",
    "checkmark", "checkbox",
    "diamond", "square", "triangle",
    "dot", "dash", "hyphen",
}

# Pre-compiled: matches **any-icon-token** with optional surrounding spaces
_RE_BOLD_ICON = re.compile(
    r'\*\*\s*(' + '|'.join(re.escape(t) for t in sorted(_ICON_TOKENS, key=len, reverse=True)) + r')\s*\*\*',
    re.IGNORECASE
)

# Matches a bare icon token at the VERY START of a string
_RE_LEADING_ICON = re.compile(
    r'^\s*(' + '|'.join(re.escape(t) for t in sorted(_ICON_TOKENS, key=len, reverse=True)) + r')[\s:\-\.]*',
    re.IGNORECASE
)

# Matches **anything** at the very start (catches unknown icon names too)
_RE_LEADING_BOLD_WORD = re.compile(r'^\s*\*\*[^*\s][^*]{0,40}\*\*\s*')

# Remaining **bold** markers (keep text, drop asterisks)
_RE_INLINE_BOLD = re.compile(r'\*\*(.+?)\*\*')

# Leading bullet/arrow/symbol characters
_RE_LEADING_SYMBOLS = re.compile(r'^[\-\–\—\•\▸\▹\►\→\✓\✔\★\☆\◆\◇\»\›\s]+')

# :emoji_name: patterns
_RE_COLON_EMOJI = re.compile(r':[a-z_\-]{2,30}:')

# Actual emoji Unicode ranges
_RE_EMOJI_CHARS = re.compile(
    r'[\U0001F300-\U0001FFFF'
    r'\U00002600-\U000027BF'
    r'\U0001F900-\U0001F9FF'
    r'\U00002700-\U000027BF]'
)

# Heuristic: bare lowercase_token at start followed by a Capital letter
_RE_BARE_ICON_HEURISTIC = re.compile(r'^([a-z][a-z0-9_\-]{2,24})\s+([A-Z])')

_COMMON_STARTERS = {
    'a','an','the','in','on','at','by','to','of','as','is','it','or','if',
    'do','be','no','so','go','use','ai','ml','api','data','with','from',
    'each','per','new','key','top','low','high','more','less','cost','this',
    'that','both','many','most','some','such','all','our','its','not','yet',
    'over','just','also','only','even','well','here','when','then','than',
    'but','and','for','nor','via','due','upon',
}

_SAFE_BULLET_ICONS = {
    "▸", "◆", "✓", "•", "●", "▪", "‣", "→", "➜", "➤", "▶", "►",
}


def _clean_bullet_text(text: str) -> str:
    """
    Guaranteed-clean bullet text.
    Strips **icon-name**, bare icon tokens, markdown, emoji, leading symbols.
    Returns empty string if nothing useful remains.
    """
    if not text or not isinstance(text, str):
        return ""

    t = text.strip()

    # ── Pass 1: remove **known-icon** patterns (e.g. **check-circle**) ──────
    t = _RE_BOLD_ICON.sub('', t)

    # ── Pass 2: remove **any-bold-word** at the very start of the string ─────
    #    This catches unknown icon names the LLM invents.
    t = _RE_LEADING_BOLD_WORD.sub('', t)

    # ── Pass 3: remove known bare icon tokens at start ────────────────────────
    t = _RE_LEADING_ICON.sub('', t)

    # ── Pass 4: inline **bold** → keep text, remove asterisks ────────────────
    t = _RE_INLINE_BOLD.sub(r'\1', t)

    # ── Pass 5: remove leading bullet/arrow/symbol prefixes ──────────────────
    t = _RE_LEADING_SYMBOLS.sub('', t)

    # ── Pass 6: remove :emoji_name: tokens ───────────────────────────────────
    t = _RE_COLON_EMOJI.sub('', t)

    # ── Pass 7: remove actual emoji characters ────────────────────────────────
    t = _RE_EMOJI_CHARS.sub('', t)

    # ── Pass 8: normalise whitespace ─────────────────────────────────────────
    t = re.sub(r'\s+', ' ', t).strip()

    # ── Pass 9: heuristic — bare lowercase_token followed by Capital letter ──
    m = _RE_BARE_ICON_HEURISTIC.match(t)
    if m:
        first = m.group(1).lower()
        if first not in _COMMON_STARTERS:
            # Strip the suspicious token
            remainder = t[m.end(1):].strip()
            if remainder and remainder[0].isupper():
                t = remainder
            elif not remainder:
                t = ""

    return t.strip()


def _validate_and_clean_content(items: list) -> list:
    """Final gate: cleans every item and discards empties."""
    result = []
    for raw in (items or []):
        t = _clean_bullet_text(str(raw) if raw is not None else "")
        if t:
            result.append(t)
    return result


def _sanitize_icon_marker(icon: str, fallback: str = "▸") -> str:
    """
    Keep only safe bullet glyphs. Any word-like token falls back to a glyph.
    """
    token = _safe_str(icon).strip()
    if not token:
        return fallback
    if token in _SAFE_BULLET_ICONS:
        return token
    if len(token) == 1 and not token.isalnum() and not token.isspace():
        return token
    return fallback

# ═══════════════════════════════════════════════════════════════════════════
#  LOW-LEVEL HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _c(t):   return RGBColor(*[max(0,min(255,int(v))) for v in t])
def _IN(v):  return Inches(float(v))
def _mix(c1, c2, t):
    t = max(0.0, min(1.0, float(t)))
    return tuple(int(c1[i]+(c2[i]-c1[i])*t) for i in range(3))
def _lum(c):
    r,g,b = [v/255 for v in c]
    return 0.2126*r + 0.7152*g + 0.0722*b
def _contrast_text(bg, dark=(0,0,0), light=(255,255,255), threshold=0.56):
    return light if _lum(bg) < threshold else dark
def _font_pt(size, role="body"):
    size = float(size)
    if role == "title":    return max(size, round(size * 1.08, 1))
    if role == "subtitle": return max(size, round(size * 1.10, 1))
    return max(size, round(size * 1.09, 1))

def _fit_big_stat_font(stat: str) -> int:
    """
    Size the large metric text so multi-line values do not collide.
    """
    text = _safe_str(stat)
    if not text:
        return 48
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        lines = [text]
    line_count = len(lines)
    longest = max(len(line) for line in lines)

    if line_count >= 4:
        return 28
    if line_count == 3:
        return 32 if longest <= 10 else 28
    if line_count == 2:
        return 40 if longest <= 12 else 36
    if len(text) <= 4:
        return 72
    if len(text) <= 8:
        return 60
    if len(text) <= 14:
        return 52
    return 44
def _is_neutral(c, thr=0.16):
    r,g,b = [v/255 for v in c]
    _,s,_ = colorsys.rgb_to_hsv(r,g,b)
    return s < thr
def _badge_text(bg):
    return _contrast_text(bg, threshold=0.46)

def _parse_color(v, fallback):
    if isinstance(v,(list,tuple)) and len(v)==3:
        try: return tuple(max(0,min(255,int(x))) for x in v)
        except: return fallback
    if isinstance(v,str):
        s = v.strip().lstrip("#")
        if re.fullmatch(r"[0-9A-Fa-f]{6}", s):
            return tuple(int(s[i:i+2],16) for i in (0,2,4))
        parts = [p.strip() for p in s.split(",")]
        if len(parts)==3:
            try: return tuple(max(0,min(255,int(x))) for x in parts)
            except: pass
    return fallback

def _parse_int(v, default=50, lo=0, hi=100):
    if isinstance(v, int): return max(lo,min(hi,v))
    m = re.search(r"-?\d+", str(v))
    return max(lo,min(hi,int(m.group()))) if m else default

def _safe_str(v, default=""):
    return str(v).strip() if v is not None else default

def _safe_list(v):
    return [str(x).strip() for x in v if str(x).strip()] if isinstance(v,list) else []

def _dedupe_keep_order(items):
    out, seen = [], set()
    for item in items or []:
        clean = _safe_str(item)
        key = clean.casefold()
        if not clean or key in seen:
            continue
        seen.add(key)
        out.append(clean)
    return out

def _norm_text(v):
    s = _safe_str(v).casefold()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s

def _looks_like_prompt_title(title: str, topic: str) -> bool:
    tt = _norm_text(title)
    tp = _norm_text(topic)
    if not tt or not tp: return False
    if tt == tp: return True
    if tp in tt and len(tp) >= max(12, int(len(tt) * 0.6)): return True
    return any(tt.startswith(prefix) for prefix in (
        "create a powerpoint deck on","create a presentation on",
        "presentation on","ppt on","deck on",
    ))

def _fallback_slide_title(slide: dict, idx: int) -> str:
    layout = _canon(slide.get("layout", "bullets"))
    if layout == "section_index": return "Contents"
    if layout == "title_cover": return "Executive Overview"
    if layout == "timeline":
        steps = slide.get("steps") or []
        if isinstance(steps, list) and steps:
            first = steps[0] if isinstance(steps[0], dict) else {}
            lbl = _safe_str(first.get("label", "Roadmap"))
            return lbl if lbl else "Roadmap"
        return "Roadmap"
    if layout == "icon_grid":
        items = slide.get("grid_items") or []
        if isinstance(items, list) and items:
            first = items[0] if isinstance(items[0], dict) else {}
            base = _safe_str(first.get("title", "Key Pillars"))
            return base if len(base.split()) <= 4 else "Key Pillars"
        return "Key Pillars"
    if layout == "case_study":
        company = _safe_str(slide.get("company", ""))
        return f"Case Study: {company}" if company else "Case Study"
    if layout == "big_stat":
        return _safe_str(slide.get("stat_label", "")) or "Key Metric"
    content = _safe_list(slide.get("content", []))
    if content:
        first = re.split(r"[:.;-]", content[0], maxsplit=1)[0].strip()
        words = first.split()
        if 1 <= len(words) <= 6: return first
    return f"Slide {idx}"

def _usage_dict(usage):
    if not usage: return None
    return {
        "prompt_tokens": int(getattr(usage, "prompt_tokens", 0) or 0),
        "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
        "total_tokens": int(getattr(usage, "total_tokens", 0) or 0),
    }

def _merge_usage(total, add):
    if not add: return total
    if total is None:
        total = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
    total["prompt_tokens"]     += int(add.get("prompt_tokens", 0) or 0)
    total["completion_tokens"] += int(add.get("completion_tokens", 0) or 0)
    total["total_tokens"]      += int(add.get("total_tokens", 0) or 0)
    return total

# ═══════════════════════════════════════════════════════════════════════════
#  THEME BUILDING
# ═══════════════════════════════════════════════════════════════════════════

_BASE = dict(
    p=(34,83,149), s=(106,161,218), a=(34,83,149), a2=(106,161,218),
    bg=(255,255,255), dk=(34,50,80), td=(34,50,80), tl=(255,255,255),
    muted=(106,161,218), card=(255,255,255), hf="Calibri", bf="Calibri",
)

def _dominant_color_from_pixels(pixels, n_clusters=5):
    if not pixels:
        return None
    if len(pixels) > 5000:
        pixels = random.sample(pixels, 5000)
    try:
        from sklearn.cluster import KMeans
        kmeans = KMeans(n_clusters=min(n_clusters, len(pixels)), random_state=0, n_init=10)
        kmeans.fit(pixels)
        colors = kmeans.cluster_centers_.astype(int)
        best = None
        best_score = -1
        for c in colors:
            r, g, b = c
            if max(r,g,b) - min(r,g,b) < 30:
                continue
            _, s, v = colorsys.rgb_to_hsv(r/255, g/255, b/255)
            score = s * v
            if score > best_score:
                best_score = score
                best = (r, g, b)
        if best is not None:
            return best
    except ImportError:
        pass
    bucket_size = 16
    buckets = {}
    for r, g, b in pixels:
        key = (r // bucket_size, g // bucket_size, b // bucket_size)
        buckets[key] = buckets.get(key, 0) + 1
    sorted_buckets = sorted(buckets.items(), key=lambda x: x[1], reverse=True)[:5]
    best = None
    best_sat = -1
    for (rb, gb, bb), _ in sorted_buckets:
        bucket_pixels = [(r, g, b) for (r, g, b) in pixels
                         if r // bucket_size == rb and g // bucket_size == gb and b // bucket_size == bb]
        if not bucket_pixels:
            continue
        avg_r = sum(p[0] for p in bucket_pixels) // len(bucket_pixels)
        avg_g = sum(p[1] for p in bucket_pixels) // len(bucket_pixels)
        avg_b = sum(p[2] for p in bucket_pixels) // len(bucket_pixels)
        _, s, v = colorsys.rgb_to_hsv(avg_r/255, avg_g/255, avg_b/255)
        if s > 0.2 and v > 0.15 and s > best_sat:
            best_sat = s
            best = (avg_r, avg_g, avg_b)
    if best:
        return best
    n = len(pixels)
    return tuple(int(sum(c[i] for c in pixels) / n) for i in range(3))


def _flatten_to_strings(obj):
    result = []
    if isinstance(obj, str):
        cleaned = _clean_bullet_text(obj)
        return [cleaned] if cleaned else []
    if isinstance(obj, dict):
        if 'points' in obj:
            return _flatten_to_strings(obj['points'])
        for v in obj.values():
            result.extend(_flatten_to_strings(v))
    elif isinstance(obj, list):
        for item in obj:
            result.extend(_flatten_to_strings(item))
    return [t for t in (_clean_bullet_text(x) for x in result) if t]


def _extract_logo_theme(logo_path: str, base: dict) -> dict:
    theme = dict(base)
    if not logo_path or not os.path.exists(logo_path):
        return theme
    try:
        with Image.open(logo_path) as img:
            rgba = img.convert("RGBA")
            bg   = Image.new("RGBA", rgba.size, (255,255,255,255))
            comp = Image.alpha_composite(bg, rgba).convert("RGB")
            comp.thumbnail((220,220))
            q    = comp.quantize(colors=6)
            pal  = q.getpalette() or []
            cnts = q.getcolors() or []
        extracted = []
        for cnt, idx in sorted(cnts, reverse=True):
            b = idx*3
            if b+2 >= len(pal): continue
            col = tuple(pal[b:b+3])
            lm  = _lum(col)
            if lm > 0.96 or lm < 0.04: continue
            extracted.append((cnt,col))
        colors  = [c for _,c in extracted]
        vivid   = [c for c in colors if not _is_neutral(c)]
        neutral = [c for c in colors if _is_neutral(c)]
        vivid.sort(key=lambda c: colorsys.rgb_to_hsv(*[v/255 for v in c])[1], reverse=True)
        primary   = vivid[0] if vivid else theme["p"]
        secondary = vivid[1] if len(vivid)>1 else _mix(primary,(255,255,255),0.35)
        accent    = vivid[2] if len(vivid)>2 else _mix(primary,(0,0,0),0.28)
        accent2   = neutral[0] if neutral else _mix(secondary,(255,255,255),0.18)
        theme.update(dict(
            p=primary, s=secondary, a=accent, a2=accent2,
            dk=_mix(primary,(12,18,28),0.42),
            bg=(255,255,255),
            td=_mix(primary,(15,15,15),0.70) if _lum(primary)>0.35 else (28,36,48),
            tl=(255,255,255),
            muted=_mix((28,36,48),(255,255,255),0.48),
            card=(255,255,255),
        ))
    except Exception as e:
        print(f"[logo-theme] {e}")
    return theme

def _build_theme(palette_name: str, design_system: dict, logo_path: str = None) -> dict:
    pal = PALETTES.get(palette_name, PALETTES["classic"])
    theme = dict(
        p=pal["p"], s=pal["s"], a=pal["a"], a2=pal["a2"],
        dk=pal["dk"], bg=(255,255,255), td=(28,36,48), tl=(255,255,255),
        muted=pal["muted"], card=(255,255,255),
        hf=pal["hf"], bf=pal["bf"],
    )
    raw = design_system.get("theme", {}) if isinstance(design_system, dict) else {}
    def _theme_color_value(value, fallback):
        if isinstance(value, str):
            token = value.strip().lower()
            if token in COLOR_HEX_MAP:
                value = COLOR_HEX_MAP[token]
        return _parse_color(value, fallback)
    for key, tkey in [("primary","p"),("secondary","s"),("accent","a"),("accent2","a2"),
                      ("bg_dark","dk"),("bg_light","bg"),("text_dark","td"),
                      ("text_light","tl"),("text_muted","muted"),("card_bg","card")]:
        if key in raw:
            theme[tkey] = _theme_color_value(raw[key], theme[tkey])
    for fkey in ("header_font","body_font"):
        tkey = "hf" if fkey=="header_font" else "bf"
        if fkey in raw and isinstance(raw[fkey], str) and raw[fkey].strip():
            theme[tkey] = raw[fkey].strip()
    if logo_path:
        theme = _extract_logo_theme(logo_path, theme)
    if _lum(theme.get("td", (0,0,0))) > 0.60:
        theme["td"] = (20, 20, 20)
    if _lum(theme.get("tl", (255,255,255))) < 0.40:
        theme["tl"] = (255, 255, 255)
    return theme

# ═══════════════════════════════════════════════════════════════════════════
#  DRAWING PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════

def _rect(slide, x, y, w, h, fill, line=None, lw=0.75):
    s = slide.shapes.add_shape(1, _IN(x), _IN(y), _IN(w), _IN(h))
    s.fill.solid(); s.fill.fore_color.rgb = _c(fill)
    if line: s.line.color.rgb = _c(line); s.line.width = Pt(lw)
    else: s.line.fill.background()
    return s

def _oval(slide, x, y, w, h, fill):
    s = slide.shapes.add_shape(9, _IN(x), _IN(y), _IN(w), _IN(h))
    s.fill.solid(); s.fill.fore_color.rgb = _c(fill)
    s.line.fill.background()
    return s

def _tb(slide, text, x, y, w, h, size, bold=False, italic=False,
        color=None, face="Calibri", align=PP_ALIGN.LEFT, shrink=False):
    # ── ICON LEAK FIX: clean text before ANY rendering ──────────────────────
    text = _clean_bullet_text(_safe_str(text))
    if not text:
        return None
    bx = slide.shapes.add_textbox(_IN(x), _IN(y), _IN(w), _IN(h))
    tf = bx.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    if shrink: tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]; p.alignment = align
    role = "title" if bold and size >= 18 else ("subtitle" if italic else "body")
    font_size = _font_pt(size, role=role)
    # Split on ** only after cleaning — so no icon name ever becomes a bold run
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part: continue
        r = p.add_run(); r.text = part
        r.font.size = Pt(font_size); r.font.bold = bold or (i % 2 == 1)
        r.font.italic = italic; r.font.name = face
        if color: r.font.color.rgb = _c(color)
    return bx

def _bullets(slide, points, x, y, w, h, size=16, icon="▸",
             ic=None, tc=None, face="Calibri", maxp=6):
    if not points: return
    ic = ic or _BASE["a"]; tc = tc or _BASE["td"]

    icon = _sanitize_icon_marker(icon)
    # ── ICON LEAK FIX: clean every bullet at the last mile ──────────────────
    pts = [_clean_bullet_text(_safe_str(p)) for p in points if _safe_str(p)]
    pts = [p for p in pts if p]          # drop empties after cleaning
    if not pts: return
    maxp = maxp or len(pts)
    pts = pts[:maxp]
    if not pts: return
    n = len(pts)
    if   n <= 2: size = min(26, size+6)
    elif n == 3: size = min(23, size+4)
    elif n == 4: size = min(20, size+2)
    elif n == 5: size = min(18, size+1)
    elif n >= 7: size = max(12, size-2)
    bx = slide.shapes.add_textbox(_IN(x),_IN(y),_IN(w),_IN(h))
    tf = bx.text_frame; tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if n<=3 else MSO_ANCHOR.TOP
    first = False
    sp = 13 if n<=3 else (6 if n>=7 else 8)
    for pt in pts:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(sp); p.space_after = Pt(sp)
        ir = p.add_run(); ir.text = f"{icon}  "
        ir.font.size = Pt(_font_pt(size + 1)); ir.font.bold = True
        ir.font.name = "Segoe UI Symbol"
        ir.font.color.rgb = _c(ic)
        # Split on ** only after the text is already clean
        for i, part in enumerate(pt.split("**")):
            if not part: continue
            r = p.add_run(); r.text = part
            r.font.size=Pt(_font_pt(size)); r.font.bold=(i%2==1)
            r.font.name=face; r.font.color.rgb=_c(tc)

# ═══════════════════════════════════════════════════════════════════════════
#  SHARED CHROME
# ═══════════════════════════════════════════════════════════════════════════

def _add_logo(slide, theme):
    if not _LOGO_PATH or not os.path.exists(_LOGO_PATH):
        return
    try:
        with Image.open(_LOGO_PATH) as img:
            w, h = img.size
            if not w or not h: return
            aspect = h / w
        max_w, max_h = 1.10, 0.60
        width = max_w
        height = width * aspect
        if height > max_h:
            height = max_h
            width = height / aspect
        left = SW - width - 0.35
        top  = 0.18
        slide.shapes.add_picture(_LOGO_PATH, _IN(left), _IN(top),
                                 width=_IN(width), height=_IN(height))
    except Exception:
        return

def _draw_header(slide, title, subtitle, num, theme, profile):
    mode = profile["header"]
    W = SW; H = HEADER_H
    tx, ty, th = 0.45, 0.18, 0.76
    title_fill = theme["dk"]
    if mode == "sidebar":
        _rect(slide, 0, 0, 0.42, H, theme["p"])
        title_fill = _mix(theme["dk"],theme["p"],0.18)
        _rect(slide, 0.42, 0, W-0.42, H, title_fill)
    elif mode == "band":
        _rect(slide, 0, 0, W, H, theme["dk"])
        _rect(slide, 0, 0, W, 0.16, theme["s"])
        title_fill = theme["dk"]
    elif mode == "split":
        _rect(slide, 0, 0, W*0.58, H, theme["dk"])
        _rect(slide, W*0.58, 0, W*0.42, H, theme["p"])
        title_fill = theme["dk"]
    else:
        _rect(slide, 0, 0, W, H, theme["dk"])
        title_fill = theme["dk"]
    logo_space = 1.40 if _LOGO_PATH else 0.50
    title_w = W - tx - logo_space
    tsize = max(22, 32 - max(0, len(title)-42)//8*4)
    _tb(slide, title, tx, ty, title_w, th, tsize,
        bold=True, color=_contrast_text(title_fill), face=theme["hf"], shrink=True)
    if subtitle:
        ssize = max(11, 14 - max(0, len(subtitle)-60)//20)
        _tb(slide, subtitle, tx, 0.96, title_w, 0.38, ssize,
            italic=True, color=_contrast_text(title_fill), face=theme["bf"], shrink=True)
    _add_logo(slide, theme)

def _draw_footer(slide, theme, profile):
    mode = profile["footer"]
    if mode == "line":
        _rect(slide, 0, SH-0.07, SW, 0.07, theme["s"])
    else:
        _rect(slide, 0, SH-FOOTER_H, SW, FOOTER_H, theme["dk"])

def _add_page_number(slide, num, theme, footer_mode="bar"):
    if num is None: return
    ns = str(num)
    if footer_mode == "line":
        y = SH - 0.30
        color = _contrast_text(theme["bg"])
    else:
        y = SH - FOOTER_H + 0.03
        color = _contrast_text(theme["dk"])
    _tb(slide, ns, 0.28, y, 0.60, 0.22, 12,
        bold=True, color=color, face=theme["bf"], align=PP_ALIGN.LEFT)

def _draw_card(slide, theme, profile, x=None, y=None, w=None, h=None):
    cx = x if x is not None else CONTENT_X
    cy = y if y is not None else CONTENT_Y
    cw = w if w is not None else CONTENT_W
    ch = h if h is not None else CONTENT_H
    mode = profile["card"]
    if mode == "soft":
        _rect(slide, cx, cy, cw, ch,
              _mix(theme["card"], theme["s"], 0.12),
              line=_mix(theme["p"],theme["s"],0.5), lw=0.8)
    elif mode == "banded":
        _rect(slide, cx, cy, cw, ch, theme["card"], line=theme["p"], lw=1.0)
        _rect(slide, cx, cy, cw, 0.12, theme["s"])
    else:
        _rect(slide, cx, cy, cw, ch, theme["card"], line=theme["p"], lw=1.0)
    pad = 0.20
    return cx+pad, cy+0.18, cw-pad*2, ch-0.18-0.28

def _slide_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _c(color)

def _accents(theme, n=4):
    base = [theme["p"], theme["a2"], theme["s"], theme["a"]]
    return [base[i % len(base)] for i in range(n)]

# ═══════════════════════════════════════════════════════════════════════════
#  LAYOUT RENDERERS
# ═══════════════════════════════════════════════════════════════════════════

def render_title_cover(slide, spec, num, theme, profile, logo_path=None):
    W, H = SW, SH
    mode = profile["cover"]
    _slide_bg(slide, theme["bg"])
    if mode == "dark_full":
        _slide_bg(slide, theme["dk"])
        _rect(slide, 0, 0, 0.35, H, theme["s"])
        _rect(slide, 0, H-0.12, W, 0.12, theme["a2"])
        tc, sc = _contrast_text(theme["dk"]), _contrast_text(theme["dk"])
    elif mode == "sidebar":
        _rect(slide, 0, 0, 2.05, H, theme["p"])
        _rect(slide, 2.05, 0, 0.10, H, theme["s"])
        tc, sc = _contrast_text(theme["bg"]), _contrast_text(theme["bg"])
    elif mode == "grid":
        _rect(slide, 0, 0, W, 1.25, theme["p"])
        _rect(slide, 0, 1.25, W, 0.08, theme["s"])
        for gx in [1.2,3.0,4.8,6.6,8.4,10.2,12.0]:
            _rect(slide, gx, 0, 0.018, H, _mix(theme["s"],theme["bg"],0.82))
        _rect(slide, 0, H-0.18, W, 0.18, _mix(theme["p"],theme["s"],0.55))
        tc, sc = _contrast_text(theme["bg"]), _contrast_text(theme["bg"])
    elif mode == "centered":
        _rect(slide, 0, 0, W, 1.06, theme["p"])
        _rect(slide, 0, H-0.18, W, 0.18, _mix(theme["p"],theme["s"],0.55))
        tc, sc = _contrast_text(theme["bg"]), _contrast_text(theme["bg"])
    else:
        _rect(slide, 0, 0, W, 1.25, theme["p"])
        _rect(slide, 0, 1.25, W, 0.08, theme["s"])
        _rect(slide, 0, H-0.18, W, 0.18, _mix(theme["p"],theme["s"],0.55))
        tc, sc = _contrast_text(theme["bg"]), _contrast_text(theme["bg"])
    title_x = 2.40 if mode == "sidebar" else 0.70
    title_w  = W - title_x - 0.40
    title    = _safe_str(spec.get("title", "Presentation"))
    tsize    = max(40, 58 - max(0, len(title)-30)//6*4)
    _tb(slide, title, title_x, 1.95, title_w, 2.20, tsize,
        bold=True, color=tc, face=theme["hf"], align=PP_ALIGN.LEFT, shrink=True)
    subtitle = _safe_str(spec.get("subtitle",""))
    if subtitle:
        ssize = max(15, 20 - max(0, len(subtitle)-60)//20)
        _tb(slide, subtitle, title_x+0.02, 4.15, title_w, 0.82, ssize,
            italic=True, color=sc, face=theme["bf"], shrink=True)
    points = [p for p in _validate_and_clean_content(_safe_list(spec.get("content",[]))) if p][:5]
    if points:
        accs   = _accents(theme, 3)
        card_w = max(2.0, (title_w - 0.18*(len(points)-1)) / len(points))
        for i, pt in enumerate(points):
            cx   = title_x + i*(card_w+0.18)
            fill = _mix(theme["card"], accs[i], 0.88 if mode=="dark_full" else 0.82)
            _rect(slide, cx, 5.00, card_w, 1.02, fill, line=accs[i], lw=1.0)
            _rect(slide, cx, 5.00, card_w, 0.10, accs[i])
            _tb(slide, pt, cx+0.12, 5.16, card_w-0.24, 0.76, 15,
                bold=True, color=_contrast_text(fill), face=theme["bf"], shrink=True)

def render_section_index(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title","Contents"), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    sections = _validate_and_clean_content(_safe_list(spec.get("sections", spec.get("content",[]))))
    if not sections:
        _tb(slide, "No sections available.", ix, iy, iw, ih, 15, color=theme["muted"])
        _add_page_number(slide, num, theme, profile["footer"]); return
    sections = sections[:6]
    cols = 2 if len(sections) > 3 else 1
    gap  = 0.18
    rows = (len(sections)+cols-1)//cols
    cw   = (iw - gap*(cols-1)) / cols
    ch   = min(1.08, max(0.72, (ih - gap*(rows-1)) / max(1,rows)))
    accs = _accents(theme, 4)
    for i, sec in enumerate(sections):
        col = i % cols; row = i // cols
        cx  = ix + col*(cw+gap)
        cy  = iy + row*(ch+gap)
        ac  = accs[i % len(accs)]
        fill= _mix(theme["card"],ac,0.90)
        _rect(slide, cx, cy, cw, ch, fill, line=ac, lw=1.0)
        _rect(slide, cx, cy, 0.72, ch, ac)
        _tb(slide, f"{i+1}", cx+0.08, cy+0.14, 0.64, ch-0.24, 17,
            bold=True, color=_contrast_text(ac), face=theme["hf"], align=PP_ALIGN.CENTER)
        _tb(slide, sec, cx+0.86, cy+0.16, cw-1.00, ch-0.28,
            18 if cols==1 else 15, bold=True, color=_contrast_text(fill),
            face=theme["bf"], shrink=True)
    _add_page_number(slide, num, theme, profile["footer"])

def render_bullets(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    clean_content = _validate_and_clean_content(_flatten_to_strings(spec.get("content", [])))
    _bullets(slide, clean_content, ix, iy, iw, ih,
             size=16, icon=spec.get("icon","▸"),
             ic=theme["a"], tc=theme["td"], face=theme["bf"], maxp=20)
    _add_page_number(slide, num, theme, profile["footer"])

def render_two_column(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    gap = 0.20; cw = (iw-gap)/2
    lt = _safe_str(spec.get("left_title","Left"))
    rt = _safe_str(spec.get("right_title","Right"))
    lp = _validate_and_clean_content(_safe_list(spec.get("left_points",[])))
    rp = _validate_and_clean_content(_safe_list(spec.get("right_points",[])))
    if not lp and not rp:
        content = _validate_and_clean_content(_safe_list(spec.get("content",[])))
        mid = max(1, len(content)//2)
        lp, rp = content[:mid], content[mid:]
    icon = spec.get("icon","▸")
    # Use the stronger accent for the bullet glyph in two-column layouts.
    # The lighter secondary accent can disappear against the white card in some
    # themes / export paths, which makes the bullets look missing even though
    # the text is present.
    bullet_ic = theme["a"]
    hh   = 0.46
    _rect(slide, ix, iy, cw, hh, theme["p"])
    _tb(slide, lt, ix+0.10, iy+0.08, cw-0.20, hh-0.12, 15,
        bold=True, color=_contrast_text(theme["p"]), face=theme["hf"])
    _bullets(slide, lp, ix+0.10, iy+hh+0.10, cw-0.20, ih-hh-0.16,
             size=15, icon=icon, ic=bullet_ic, tc=theme["td"], face=theme["bf"], maxp=20)
    rx = ix+cw+gap
    _rect(slide, rx, iy, cw, hh, theme["a2"])
    _tb(slide, rt, rx+0.10, iy+0.08, cw-0.20, hh-0.12, 15,
        bold=True, color=_contrast_text(theme["a2"]), face=theme["hf"])
    _bullets(slide, rp, rx+0.10, iy+hh+0.10, cw-0.20, ih-hh-0.16,
             size=15, icon=icon, ic=bullet_ic, tc=theme["td"], face=theme["bf"], maxp=20)
    _add_page_number(slide, num, theme, profile["footer"])

def render_big_stat(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    pw     = 3.70
    stat   = _safe_str(spec.get("stat","—"))
    label  = _safe_str(spec.get("stat_label",""))
    source = _safe_str(spec.get("stat_source",""))
    pts    = _validate_and_clean_content(_safe_list(spec.get("content",[])))
    stat_lines = [line.strip() for line in stat.splitlines() if line.strip()] or [stat]
    sfont = _fit_big_stat_font(stat)
    stat_box_h = 1.90 if len(stat_lines) == 1 else 2.05 if len(stat_lines) == 2 else 2.25
    _rect(slide, ix, iy, pw, ih, theme["p"])
    _rect(slide, ix, iy, pw, 0.07, theme["a2"])
    _tb(slide, stat, ix+0.10, iy+0.28, pw-0.20, stat_box_h, sfont,
        bold=True, color=_contrast_text(theme["p"]), face=theme["hf"], align=PP_ALIGN.CENTER, shrink=True)
    label_y = iy + (2.15 if len(stat_lines) == 1 else 2.38 if len(stat_lines) == 2 else 2.62)
    _tb(slide, label, ix+0.10, label_y, pw-0.20, 0.62, 14,
        color=_contrast_text(theme["p"]), face=theme["bf"], align=PP_ALIGN.CENTER, shrink=True)
    if source:
        source_y = label_y + 0.62
        _tb(slide, source, ix+0.10, source_y, pw-0.20, 0.46, 10,
            italic=True, color=_contrast_text(theme["p"]), face=theme["bf"], align=PP_ALIGN.CENTER, shrink=True)
    try:
        pct = float(stat.replace("%","").replace("+","").strip())
        if 0 < pct < 100:
            bmar = 0.30; bw = pw-bmar*2; bar_y = iy+ih-0.55; bh = 0.20
            _rect(slide, ix+bmar, bar_y, bw, bh, (255,255,255))
            _rect(slide, ix+bmar, bar_y, bw*(pct/100), bh, theme["a2"])
    except: pass
    bx = ix+pw+0.22; bw2 = iw-pw-0.22
    _bullets(slide, pts, bx, iy+0.10, bw2, ih-0.20,
             size=14, icon=spec.get("icon","▸"),
             ic=theme["a"], tc=theme["td"], face=theme["bf"], maxp=20)
    _add_page_number(slide, num, theme, profile["footer"])

def render_timeline(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    steps = [s for s in (spec.get("steps") or [])
             if isinstance(s,dict) and _safe_str(s.get("label")) and _safe_str(s.get("detail"))]
    if not steps:
        _tb(slide, "No timeline steps provided.", ix, iy, iw, ih, 14, color=theme["muted"])
        _add_page_number(slide, num, theme); return
    n = len(steps); sw = iw/n
    tl_y = iy + ih*0.47
    DR, dr = 0.28, 0.14
    accs = _accents(theme, n)
    _rect(slide, ix, tl_y, iw, 0.06, theme["s"])
    for i, step in enumerate(steps):
        cx = ix + i*sw + sw/2
        ac = accs[i % len(accs)]
        inner_fill = (255,255,255)
        _oval(slide, cx-DR, tl_y-DR, DR*2, DR*2, ac)
        _oval(slide, cx-dr, tl_y-dr, dr*2, dr*2, inner_fill)
        _tb(slide, str(i+1), cx-0.20, tl_y-0.16, 0.40, 0.32, 12,
            bold=True, color=_badge_text(inner_fill), align=PP_ALIGN.CENTER)
        cw_ = sw*0.84; cx_c = cx-cw_/2
        above = (i%2==0)
        if above:
            cy = iy+0.05; ch = max(0.40, tl_y-DR-0.42-cy)
            if tl_y-DR > cy+ch+0.02:
                _rect(slide, cx-0.025, cy+ch, 0.05, tl_y-DR-(cy+ch), theme["s"])
        else:
            cy = tl_y+DR+0.28; ch = max(0.40, iy+ih-0.05-cy)
            if cy > tl_y+DR+0.02:
                _rect(slide, cx-0.025, tl_y+DR, 0.05, cy-(tl_y+DR), theme["s"])
        _rect(slide, cx_c, cy, cw_, ch, theme["card"], line=ac, lw=1.0)
        _tb(slide, step["label"], cx_c+0.10, cy+0.08, cw_-0.20, 0.42, 13,
            bold=True, color=_contrast_text(theme["card"]), face=theme["hf"])
        dh = ch-0.56
        if dh > 0.12:
            _tb(slide, step["detail"], cx_c+0.10, cy+0.52, cw_-0.20, dh, 12,
                color=_contrast_text(theme["card"]), face=theme["bf"])
    _add_page_number(slide, num, theme)

def render_icon_grid(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    items = [g for g in (spec.get("grid_items") or [])
             if isinstance(g,dict) and _safe_str(g.get("title"))][:4]
    if not items:
        _tb(slide, "No grid items provided.", ix, iy, iw, ih, 14, color=theme["muted"])
        _add_page_number(slide, num, theme, profile["footer"]); return
    cols = 2; gap = 0.16
    cw_  = (iw-gap)/cols; ch_ = (ih-gap)/2
    accs = _accents(theme, 4)
    IR   = 0.38
    for i, gi in enumerate(items):
        col = i%cols; row = i//cols
        cx  = ix + col*(cw_+gap)
        cy  = iy + row*(ch_+gap)
        ac  = accs[i % len(accs)]
        _rect(slide, cx, cy, cw_, ch_, theme["card"], line=ac, lw=1.0)
        _rect(slide, cx, cy, 0.09, ch_, ac)
        iox = cx+0.24; ioy = cy+(ch_-IR*2)/2
        _oval(slide, iox, ioy, IR*2, IR*2, ac)
        raw = _safe_str(gi.get("icon", ""))
        ch1 = next((c.upper() for c in raw if c.isascii() and c.isalnum()), _seq_icon(i))
        _tb(slide, ch1, iox+0.04, ioy+0.08, IR*2-0.08, IR*1.6, 17,
            bold=True, color=_contrast_text(ac), face=theme["hf"], align=PP_ALIGN.CENTER)
        tx = iox+IR*2+0.16; tw = cw_-(iox-cx)-IR*2-0.22
        _tb(slide, _clean_bullet_text(_safe_str(gi.get("title",""))), tx, cy+0.12, tw, 0.46, 16,
            bold=True, color=_contrast_text(theme["card"]), face=theme["hf"])
        _tb(slide, _clean_bullet_text(_safe_str(gi.get("detail",""))), tx, cy+0.60, tw, ch_-0.72, 14,
            color=_contrast_text(theme["card"]), face=theme["bf"])
    _add_page_number(slide, num, theme, profile["footer"])

def render_case_study(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    company = _safe_str(spec.get("company","Organisation"))
    result  = _safe_str(spec.get("result",""))
    metrics = [m for m in (spec.get("metrics") or []) if isinstance(m,dict)]
    pts     = _validate_and_clean_content(_safe_list(spec.get("content",[])))
    icon    = spec.get("icon","▸")
    bh = 0.52; rh = 0.48; mh = 0.50
    _rect(slide, ix, iy, iw, bh, theme["p"])
    _tb(slide, f"  {company}", ix+0.18, iy+0.10, iw-0.36, bh-0.16, 17,
        bold=True, color=_contrast_text(theme["p"]), face=theme["hf"])
    cur_y = iy+bh+0.06
    if result:
        _rect(slide, ix, cur_y, iw, rh, theme["a2"])
        _tb(slide, f"  {result}", ix+0.14, cur_y+0.10, iw-0.28, rh-0.16, 13,
            bold=True, color=_contrast_text(theme["a2"]), face=theme["bf"])
        cur_y += rh+0.06
    metric_labels = []
    metrics_used  = 0
    for i, m in enumerate(metrics[:3]):
        lbl = _safe_str(m.get("label","Impact"))
        val = _parse_int(m.get("value",50), default=50, lo=1, hi=99)
        metric_labels.append(lbl.lower())
        my  = cur_y + i*mh
        mlw = 2.90; bar_x = ix+mlw+0.20; bar_w = iw-mlw-0.40
        _tb(slide, lbl, ix+0.14, my+0.04, mlw, 0.30, 13, color=_contrast_text(theme["card"]))
        _rect(slide, bar_x, my+0.06, bar_w, 0.22, (225,232,245))
        _rect(slide, bar_x, my+0.06, bar_w*(val/100), 0.22, theme["a2"])
        label_w  = 0.60
        bar_fill = bar_w * (val / 100)
        if bar_fill >= label_w + 0.10:
            label_x = bar_x + bar_fill - label_w - 0.04
        else:
            label_x = bar_x + 0.04
        _tb(slide, f"{val}", label_x, my+0.01, label_w, 0.22, 11,
            bold=True, color=_contrast_text(theme["card"]))
        metrics_used = (i+1)*mh
    if metric_labels:
        filtered = []
        for p in pts:
            pl = p.lower()
            if any(pl.startswith(lbl) or pl.startswith(lbl+":") for lbl in metric_labels):
                continue
            filtered.append(p)
        pts = filtered
    bul_y = cur_y + (metrics_used if metrics else 0) + 0.10
    bul_h = ih-(bul_y-iy)-0.10
    if bul_h > 0.30 and pts:
        max_bullets = min(4, max(1, int(bul_h / 0.36)))
        _bullets(slide, pts, ix+0.10, bul_y, iw-0.20, bul_h,
                 size=14, icon=icon, ic=theme["a"], tc=theme["td"],
                 face=theme["bf"], maxp=max_bullets)
    _add_page_number(slide, num, theme, profile["footer"])

def render_table(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    cols = _safe_list(spec.get("table_columns",[]))[:5]
    rows = [r for r in (spec.get("table_rows") or []) if isinstance(r,list)][:6]
    if not cols or not rows:
        _tb(slide, "No table data provided.", ix, iy, iw, ih, 14, color=theme["muted"])
        _add_page_number(slide, num, theme, profile["footer"]); return
    nc    = len(cols)
    hh    = 0.55; cg = 0.02; pad = 0.06
    col_w = (iw-(nc-1)*cg)/nc
    row_h = min(0.65, max(0.38, (ih-hh-0.08)/len(rows)))
    for ci, col in enumerate(cols):
        cx = ix + ci*(col_w+cg)
        _rect(slide, cx, iy, col_w, hh, theme["p"], line=theme["s"], lw=0.8)
        _tb(slide, col, cx+pad, iy+pad, col_w-pad*2, hh-pad*2, 13,
            bold=True, color=_contrast_text(theme["p"]), face=theme["hf"],
            align=PP_ALIGN.CENTER, shrink=True)
    for ri, row in enumerate(rows):
        ry   = iy+hh+ri*row_h
        fill = theme["card"] if ri%2==0 else _mix(theme["card"],theme["s"],0.35)
        cells = [_safe_str(v) for v in row[:nc]]
        while len(cells)<nc: cells.append("")
        for ci, cell in enumerate(cells):
            cx = ix + ci*(col_w+cg)
            _rect(slide, cx, ry, col_w, row_h, fill,
                  line=_mix(theme["p"],theme["card"],0.65), lw=0.5)
            _tb(slide, cell, cx+pad, ry+pad, col_w-pad*2, row_h-pad, 12,
                color=_contrast_text(fill), face=theme["bf"], shrink=True)
    _add_page_number(slide, num, theme, profile["footer"])

def render_chart(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    chart_data = [x for x in (spec.get("chart_data") or [])
                  if isinstance(x,dict) and _safe_str(x.get("label"))][:5]
    if not chart_data:
        _tb(slide, "No chart data provided.", ix, iy, iw, ih, 14, color=theme["muted"])
        _add_page_number(slide, num, theme, profile["footer"]); return
    ct = _safe_str(spec.get("chart_title", spec.get("title","")))
    content = _validate_and_clean_content(_flatten_to_strings(spec.get("content", [])))[:4]
    _tb(slide, ct, ix+0.02, iy+0.00, iw-0.04, 0.56, 18,
        bold=True, color=_contrast_text(theme["card"]), face=theme["hf"], shrink=True)
    values = [max(1,_parse_int(d.get("value",0),default=1,lo=1,hi=100)) for d in chart_data]
    vmax   = max(values)
    lw     = max(2.20, iw*0.30); bar_x = ix+lw+0.18; bar_w = iw-lw-0.72
    row_h  = 0.30; row_gap = 0.18
    chart_top = iy + 0.62
    avail_h = ih - 0.62 - (1.65 if content else 0.28)
    needed  = len(chart_data)*row_h + (len(chart_data)-1)*row_gap
    if needed > avail_h and len(chart_data)>1:
        sc = avail_h/needed; row_h *= sc; row_gap *= sc
    accs = _accents(theme, len(chart_data))
    for i, item in enumerate(chart_data):
        y   = chart_top + i*(row_h+row_gap)
        lbl = _safe_str(item.get("label",""))[:40]
        val = max(1, _parse_int(item.get("value",0),default=1,lo=1,hi=100))
        pct = val/vmax
        _tb(slide, lbl, ix+0.02, y+0.01, lw-0.10, row_h-0.02, 15,
            color=_contrast_text(theme["card"]), face=theme["bf"], shrink=True)
        _rect(slide, bar_x, y, bar_w, row_h, _mix(theme["card"],theme["s"],0.90))
        _rect(slide, bar_x, y, bar_w*pct, row_h, accs[i])
        _tb(slide, str(val), bar_x+bar_w+0.08, y-0.01, 0.58, row_h+0.04, 13,
            bold=True, color=_contrast_text(theme["card"]), align=PP_ALIGN.LEFT)
    src = _safe_str(spec.get("chart_source",""))
    if src:
        _tb(slide, src, ix+0.02, iy+ih-0.24, iw-0.04, 0.18, 10,
            italic=True, color=theme["muted"], face=theme["bf"])
    if content:
        take_y = max(iy + 1.95, chart_top + needed + 0.18)
        if take_y + 0.55 < iy + ih - 0.26:
            _tb(slide, "Key Takeaways", ix+0.02, take_y, iw-0.04, 0.26, 13,
                bold=True, color=_contrast_text(theme["card"]), face=theme["hf"])
            _bullets(slide, content, ix+0.00, take_y+0.20, iw-0.04, max(0.30, iy+ih-take_y-0.40),
                     size=13, icon="▸", ic=theme["a"], tc=theme["td"], face=theme["bf"], maxp=4)
    _add_page_number(slide, num, theme, profile["footer"])

def render_hybrid_insight(slide, spec, num, theme, profile, logo_path=None):
    _slide_bg(slide, theme["bg"])
    _draw_header(slide, spec.get("title",""), spec.get("subtitle",""), num, theme, profile)
    _draw_footer(slide, theme, profile)
    ix, iy, iw, ih = _draw_card(slide, theme, profile)
    lw    = iw*0.43; gap = 0.20; rw = iw-lw-gap; rx = ix+lw+gap
    stat  = _safe_str(spec.get("stat","—"))
    label = _safe_str(spec.get("stat_label","Key Indicator"))
    source = _safe_str(spec.get("stat_source",""))
    left_fill = _mix(theme["card"],theme["s"],0.86)
    _rect(slide, ix, iy, lw, ih, left_fill, line=theme["s"], lw=1.0)
    # Break long stat strings into cleaner lines so the KPI never collides
    # with the label or source text.
    stat_lines = stat.replace("  ", " ").strip()
    if " " in stat_lines and "\n" not in stat_lines:
        parts = stat_lines.split()
        if len(parts) >= 3:
            stat_lines = f"{parts[0]}\n{' '.join(parts[1:-1])}\n{parts[-1]}"
        elif len(parts) == 2:
            stat_lines = f"{parts[0]}\n{parts[1]}"
    stat_height = 1.55 if "\n" in stat_lines else 1.20
    sfont = 40 if len(stat_lines) <= 8 else 34
    _tb(slide, stat_lines, ix+0.12, iy+0.22, lw-0.24, stat_height, sfont,
        bold=True, color=_contrast_text(left_fill), face=theme["hf"], align=PP_ALIGN.CENTER)
    _tb(slide, label, ix+0.12, iy+1.86, lw-0.24, 0.34, 12,
        color=_contrast_text(left_fill), face=theme["bf"], align=PP_ALIGN.CENTER)
    if source:
        _tb(slide, source, ix+0.10, iy+2.24, lw-0.20, 0.30, 9,
            italic=True, color=_contrast_text(left_fill), face=theme["bf"], align=PP_ALIGN.CENTER)
    chart = [r for r in (spec.get("chart_data") or []) if isinstance(r,dict)][:3]
    if chart:
        vmax = max(max(1,_parse_int(r.get("value",1),default=1,lo=1,hi=100)) for r in chart)
        by   = iy+2.68; row_h = 0.26; row_gap = 0.16
        for i, row in enumerate(chart):
            y   = by + i*(row_h+row_gap)
            v   = max(1, _parse_int(row.get("value",1),default=1,lo=1,hi=100))
            lbl = _safe_str(row.get("label",""))[:18]
            _tb(slide, lbl, ix+0.12, y-0.01, 1.40, row_h, 11,
                color=_contrast_text(left_fill), face=theme["bf"], shrink=True)
            bar_x = ix+1.55; bar_w = lw-2.08
            _rect(slide, bar_x, y, bar_w, row_h, _mix(theme["card"],theme["s"],0.93))
            _rect(slide, bar_x, y, bar_w*(v/vmax), row_h, theme["a2"])
            vlabel_w = 0.60
            bar_fill = bar_w * (v / vmax)
            if bar_fill >= vlabel_w + 0.10:
                vlabel_x = bar_x + bar_fill - vlabel_w - 0.04
            else:
                vlabel_x = bar_x + 0.04
            _tb(slide, str(v), vlabel_x, y-0.01, vlabel_w, row_h, 10,
                bold=True, color=_contrast_text(left_fill), align=PP_ALIGN.RIGHT, shrink=True)
    _rect(slide, rx, iy, rw, ih, theme["card"], line=theme["p"], lw=1.0)
    clean_content = _validate_and_clean_content(_flatten_to_strings(spec.get("content",[])))
    _bullets(slide, clean_content, rx+0.10, iy+0.10, rw-0.20, ih-0.20,
             size=14, icon=spec.get("icon","▸"),
             ic=theme["a"], tc=theme["td"], face=theme["bf"], maxp=20)
    _add_page_number(slide, num, theme, profile["footer"])

# ═══════════════════════════════════════════════════════════════════════════
#  LAYOUT REGISTRY
# ═══════════════════════════════════════════════════════════════════════════

_RENDERERS = {
    "title_cover":    render_title_cover,
    "section_index":  render_section_index,
    "bullets":        render_bullets,
    "two_column":     render_two_column,
    "big_stat":       render_big_stat,
    "timeline":       render_timeline,
    "icon_grid":      render_icon_grid,
    "case_study":     render_case_study,
    "table":          render_table,
    "chart":          render_chart,
    "hybrid_insight": render_hybrid_insight,
}

_CONTENT_LAYOUTS = list(_RENDERERS.keys())

_ALIASES = {
    "contents":"section_index","agenda":"section_index","index":"section_index",
    "comparison":"two_column","grid":"icon_grid","infographic":"icon_grid",
    "kpi":"big_stat","metric":"big_stat","chart_slide":"chart","graph":"chart",
    "bar_chart":"chart","hybrid":"hybrid_insight","storytelling":"hybrid_insight",
    "two_col":"two_column",
}

def _canon(layout, default="bullets"):
    raw = str(layout or "").strip().lower()
    raw = _ALIASES.get(raw, raw)
    return raw if raw in _RENDERERS else default

# ═══════════════════════════════════════════════════════════════════════════
#  CONTENT VALIDATION & REPAIR
# ═══════════════════════════════════════════════════════════════════════════

def _needs_repair(slide: dict) -> bool:
    layout = _canon(slide.get("layout","bullets"))
    if layout == "two_column":
        lp = _safe_list(slide.get("left_points",[]))
        rp = _safe_list(slide.get("right_points",[]))
        return len(lp)<2 and len(rp)<2
    if layout == "timeline":
        steps = [s for s in (slide.get("steps") or [])
                 if isinstance(s,dict) and _safe_str(s.get("label")) and _safe_str(s.get("detail"))]
        return len(steps) < 3
    if layout == "icon_grid":
        items = [g for g in (slide.get("grid_items") or [])
                 if isinstance(g,dict) and _safe_str(g.get("title"))]
        return len(items) < 3
    if layout == "case_study":
        return not _safe_str(slide.get("company"))
    if layout == "table":
        cols = _safe_list(slide.get("table_columns",[]))
        rows = [r for r in (slide.get("table_rows") or []) if isinstance(r,list)]
        return len(cols)<2 or len(rows)<2
    if layout == "chart":
        data = [x for x in (slide.get("chart_data") or []) if isinstance(x,dict)]
        return len(data) < 2
    if layout == "big_stat":
        return not _safe_str(slide.get("stat"))
    if layout == "hybrid_insight":
        return not _safe_str(slide.get("stat"))
    return False

_REPAIR_PROMPT = {
    "two_column": """
Return ONLY valid JSON for a two_column slide about "{title}".
Fields required: left_title, right_title, left_points (list of 4 plain English strings), right_points (list of 4 plain English strings).
IMPORTANT: No icon names, no markdown bold, no bullet symbols in any string.
Topic context: {context}
""",
    "timeline": """
Return ONLY valid JSON for a timeline slide about "{title}".
Fields required: steps — list of 4 objects each with "label" (short phase name) and "detail" (1-sentence plain English description).
IMPORTANT: No icon names, no markdown, no symbols in any string.
Topic context: {context}
""",
    "icon_grid": """
Return ONLY valid JSON for an icon_grid slide about "{title}".
Fields required: grid_items — list of 4 objects each with "icon" (single letter A-Z), "title" (2-4 plain English words), "detail" (1 plain English sentence).
IMPORTANT: No icon names, no markdown, no symbols in title or detail.
Topic context: {context}
""",
    "case_study": """
Return ONLY valid JSON for a case_study slide about "{title}".
Fields required: company (real org name), result (1 plain English sentence),
metrics (list of 3 objects with "label" and "value" 1-99),
content (list of 4 plain English bullet strings, no icon names).
Topic context: {context}
""",
    "table": """
Return ONLY valid JSON for a table slide about "{title}".
Fields required: table_columns (list of 3-5 column names), table_rows (list of 4-5 rows, each a list of plain strings).
Topic context: {context}
""",
    "chart": """
Return ONLY valid JSON for a chart slide about "{title}".
Fields required: chart_title (string), chart_data (list of 4-5 objects with "label" and "value" 1-100),
chart_source (optional string).
Topic context: {context}
""",
    "big_stat": """
Return ONLY valid JSON for a big_stat slide about "{title}".
Fields required: stat (e.g. "42%"), stat_label (short descriptor), stat_source (source/year),
content (list of 4 plain English bullet strings, no icon names, no markdown).
Topic context: {context}
""",
    "hybrid_insight": """
Return ONLY valid JSON for a hybrid_insight slide about "{title}".
Fields required: stat (e.g. "3.2x"), stat_label (short descriptor),
chart_data (list of 3 objects with "label" and "value" 1-100),
content (list of 4 plain English bullet strings, no icon names, no markdown).
Topic context: {context}
""",
}

def _repair_slide(slide: dict, topic: str):
    layout   = _canon(slide.get("layout","bullets"))
    template = _REPAIR_PROMPT.get(layout)
    if not template: return slide, None
    context = f"{topic} — slide: {slide.get('title','')} — existing content: {slide.get('content',[][:2])}"
    prompt  = template.format(title=slide.get("title",""), context=context)
    usage   = None
    try:
        resp  = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You output ONLY valid JSON. Every text value must be plain English. "
                        "Never include icon names like check-circle, star, arrow, or any markdown syntax."
                    )
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=900,
        )
        usage = _usage_dict(resp.usage)
        raw   = (resp.choices[0].message.content or "").strip()
        raw   = re.sub(r"^```[a-z]*\n?","",raw).rstrip("`").strip()
        data  = json.loads(raw)
        if isinstance(data, dict):
            slide.update(data)
    except Exception as e:
        print(f"[repair:{layout}] {e}")
    return slide, usage

# ═══════════════════════════════════════════════════════════════════════════
#  JSON PARSING
# ═══════════════════════════════════════════════════════════════════════════

def _parse_json(raw: str) -> dict:
    if not raw: return {}
    txt = raw.strip()
    txt = re.sub(r"^```[a-z]*\n?","",txt).rstrip("`").strip()
    try: return json.loads(txt)
    except:
        try: return ast.literal_eval(txt)
        except: pass
        s = txt.find("{"); e = txt.rfind("}")
        if s>=0 and e>s:
            try: return json.loads(txt[s:e+1])
            except:
                try: return ast.literal_eval(txt[s:e+1])
                except: pass
    return {}

def _seq_icon(idx: int) -> str:
    if 0 <= idx < 26: return chr(ord("A") + idx)
    return str(idx + 1)

def _normalize_icon_grid_items(slide: dict) -> list:
    items = slide.get("grid_items", [])
    if not isinstance(items, list): items = []
    cleaned = []
    for i, g in enumerate(items):
        if not isinstance(g, dict): continue
        t = _clean_bullet_text(_safe_str(g.get("title", "")))
        d = _clean_bullet_text(_safe_str(g.get("detail", g.get("description", ""))))
        if not t: continue
        icon = _safe_str(g.get("icon", "")) or _seq_icon(i)
        cleaned.append({"icon": icon, "title": t, "detail": d})
    if cleaned: return cleaned[:4]
    content    = slide.get("content", [])
    raw_text   = "\n".join(str(x) for x in content) if isinstance(content, list) else _safe_str(content)
    parsed     = _parse_json(raw_text) if raw_text.strip() else {}
    parsed_items = None
    if isinstance(parsed, dict) and isinstance(parsed.get("grid_items"), list):
        parsed_items = parsed.get("grid_items")
    elif isinstance(parsed, list):
        parsed_items = parsed
    if isinstance(parsed_items, list):
        for i, g in enumerate(parsed_items):
            if not isinstance(g, dict): continue
            t = _clean_bullet_text(_safe_str(g.get("title", "")))
            d = _clean_bullet_text(_safe_str(g.get("detail", g.get("description", ""))))
            if not t: continue
            icon = _safe_str(g.get("icon", "")) or _seq_icon(i)
            cleaned.append({"icon": icon, "title": t, "detail": d})
        if cleaned: return cleaned[:4]
    lines = []
    if isinstance(content, list):
        lines = [_safe_str(x) for x in content if _safe_str(x)]
    elif raw_text.strip():
        lines = [_safe_str(x) for x in raw_text.splitlines() if _safe_str(x)]
    for i, line in enumerate(lines):
        if len(cleaned) >= 4: break
        parsed_line = _parse_json(line)
        if isinstance(parsed_line, dict):
            t    = _clean_bullet_text(_safe_str(parsed_line.get("title", "")))
            d    = _clean_bullet_text(_safe_str(parsed_line.get("detail", parsed_line.get("description", ""))))
            icon = _safe_str(parsed_line.get("icon", "")) or _seq_icon(i)
            if t:
                cleaned.append({"icon": icon, "title": t, "detail": d})
                continue
        title, detail = line, ""
        if ":" in line:
            title, detail = [part.strip() for part in line.split(":", 1)]
        title  = _clean_bullet_text(_safe_str(title))
        detail = _clean_bullet_text(_safe_str(detail))
        if title:
            cleaned.append({"icon": _seq_icon(i), "title": title[:60], "detail": detail})
    return cleaned[:4]

# ═══════════════════════════════════════════════════════════════════════════
#  SLIDE NORMALIZER
# ═══════════════════════════════════════════════════════════════════════════

def _normalize_slide(slide: dict, idx: int, topic: str) -> dict:
    slide = dict(slide)
    slide.setdefault("title", f"Slide {idx}")
    slide.setdefault("subtitle", "")
    slide.setdefault("content", [])
    slide["icon"] = _sanitize_icon_marker(slide.get("icon", "▸"))
    if not isinstance(slide.get("style"), dict): slide["style"] = {}

    layout = _canon(slide.get("layout","bullets"))
    slide["layout"] = layout
    if _looks_like_prompt_title(slide.get("title", ""), topic):
        slide["title"] = _fallback_slide_title(slide, idx)

    for k in ("content","left_points","right_points","sections"):
        if k in slide:
            slide[k] = _safe_list(slide[k])

    if layout == "two_column":
        if not slide.get("left_title"):  slide["left_title"]  = "Option A"
        if not slide.get("right_title"): slide["right_title"] = "Option B"
        if not slide.get("left_points") and not slide.get("right_points"):
            c   = _safe_list(slide.get("content",[]))
            mid = max(1, len(c)//2)
            slide["left_points"]  = c[:mid]
            slide["right_points"] = c[mid:]
    elif layout == "big_stat":
        slide.setdefault("stat_label","Key Metric")
        slide.setdefault("stat_source","")
    elif layout == "timeline":
        steps = slide.get("steps",[])
        if not isinstance(steps, list): steps = []
        slide["steps"] = [
            {
                "label":  _clean_bullet_text(_safe_str(s.get("label"))),
                "detail": _clean_bullet_text(_safe_str(s.get("detail"))),
            }
            for s in steps
            if isinstance(s,dict) and _safe_str(s.get("label")) and _safe_str(s.get("detail"))
        ][:5]
    elif layout == "icon_grid":
        slide["grid_items"] = _normalize_icon_grid_items(slide)
        slide["content"]    = [
            f"{item['title']}: {item['detail']}".rstrip(": ").strip()
            for item in slide["grid_items"]
        ]
    elif layout == "case_study":
        slide.setdefault("company","Organisation")
        slide.setdefault("result","")
        metrics = slide.get("metrics",[])
        if not isinstance(metrics,list): metrics = []
        slide["metrics"] = [
            {"label":_safe_str(m.get("label","Metric")),
             "value":_parse_int(m.get("value",50),default=50,lo=1,hi=99)}
            for m in metrics if isinstance(m,dict)
        ][:3]
    elif layout == "table":
        cols = _safe_list(slide.get("table_columns",[]))[:5]
        rows = [r for r in (slide.get("table_rows") or []) if isinstance(r,list)][:6]
        slide["table_columns"] = cols
        slide["table_rows"]    = []
        for row in rows:
            cells = [_safe_str(v) for v in row[:len(cols)]]
            while len(cells)<len(cols): cells.append("")
            if any(cells): slide["table_rows"].append(cells)
    elif layout == "chart":
        data = slide.get("chart_data",[])
        if not isinstance(data,list): data = []
        slide["chart_data"] = [
            {"label":_safe_str(d.get("label",""))[:40],
             "value":_parse_int(d.get("value",50),default=50,lo=1,hi=100)}
            for d in data if isinstance(d,dict) and _safe_str(d.get("label",""))
        ][:5]
        slide.setdefault("chart_title", slide["title"])
        slide.setdefault("chart_source","")
    elif layout == "hybrid_insight":
        slide.setdefault("stat_label","Key Indicator")
        data = slide.get("chart_data",[])
        if not isinstance(data,list): data = []
        slide["chart_data"] = [
            {"label":_safe_str(d.get("label",""))[:24],
             "value":_parse_int(d.get("value",50),default=50,lo=1,hi=100)}
            for d in data if isinstance(d,dict) and _safe_str(d.get("label",""))
        ][:3]
    elif layout == "section_index":
        secs = _safe_list(slide.get("sections", slide.get("content",[])))
        slide["sections"] = secs[:6]
        slide["content"]  = secs[:6]

    # ── Final content cleaning gate (all fields) ─────────────────────────────
    slide["content"] = _validate_and_clean_content(_flatten_to_strings(slide.get("content", [])))

    if slide.get("layout") == "two_column":
        slide["left_points"]  = _validate_and_clean_content(_flatten_to_strings(slide.get("left_points", [])))
        slide["right_points"] = _validate_and_clean_content(_flatten_to_strings(slide.get("right_points", [])))

    if slide.get("layout") == "section_index":
        slide["sections"] = _validate_and_clean_content(slide.get("sections", []))
        slide["content"]  = slide["sections"]

    if slide.get("layout") == "timeline" and "steps" in slide:
        for step in slide["steps"]:
            if isinstance(step, dict):
                step["label"]  = _clean_bullet_text(str(step.get("label", "")))
                step["detail"] = _clean_bullet_text(str(step.get("detail", "")))

    if slide.get("layout") == "icon_grid" and "grid_items" in slide:
        for item in slide["grid_items"]:
            if isinstance(item, dict):
                item["title"]  = _clean_bullet_text(str(item.get("title", "")))
                item["detail"] = _clean_bullet_text(str(item.get("detail", "")))

    return slide


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN CONTENT GENERATION
# ═══════════════════════════════════════════════════════════════════════════

COLOR_HEX_MAP = {
    "red": "#E53935", "blue": "#1565C0", "green": "#2E7D32",
    "yellow": "#F9A825", "orange": "#E65100", "purple": "#6A1B9A",
    "pink": "#AD1457", "cyan": "#00838F", "teal": "#00695C",
    "white": "#FFFFFF", "black": "#212121", "gray": "#546E7A",
    "grey": "#546E7A", "navy": "#0D1B4B", "gold": "#F59E0B",
    "silver": "#9E9E9E", "brown": "#5D4037", "indigo": "#283593",
    "violet": "#4527A0", "magenta": "#AD1457", "lime": "#558B2F",
    "dark blue": "#0D2B6B", "light blue": "#42A5F5",
    "dark green": "#1B5E20", "light green": "#66BB6A",
    "dark red": "#B71C1C", "light red": "#EF9A9A",
}

def _hex_to_rgb(hex_color: str):
    s = str(hex_color or "").strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", s):
        return None
    return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))

def _rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*[max(0, min(255, int(v))) for v in rgb])

def _mix_hex(hex_a: str, hex_b: str, t: float) -> str:
    rgb_a = _hex_to_rgb(hex_a)
    rgb_b = _hex_to_rgb(hex_b)
    if not rgb_a or not rgb_b:
        return hex_a or hex_b or "#000000"
    t = max(0.0, min(1.0, float(t)))
    mixed = tuple(int(rgb_a[i] + (rgb_b[i] - rgb_a[i]) * t) for i in range(3))
    return _rgb_to_hex(mixed)

def _resolve_theme_color(name: str):
    token = str(name or "").strip().lower()
    if not token:
        return ""
    if token in COLOR_HEX_MAP:
        return COLOR_HEX_MAP[token]
    if re.fullmatch(r"#?[0-9A-Fa-f]{6}", token):
        return token if token.startswith("#") else f"#{token}"
    return ""

def _prioritize_user_theme(primary_name: str, secondary_name: str):
    """
    Prefer a vivid color as the primary theme color.

    Requests such as "white and green" should render as a white background
    with green accents, not a gray-first palette.
    """
    neutral = {"white", "black", "gray", "grey", "silver"}
    p = str(primary_name or "").strip().lower()
    s = str(secondary_name or "").strip().lower()
    if p in neutral and s and s not in neutral:
        return secondary_name, primary_name
    return primary_name, secondary_name

def generate_slide_content(topic: str, num_slides: int = 6,
                           tone: str = "Professional", theme_colors: dict = None) -> dict:
    target       = max(3, int(num_slides or 6))
    palette_name = random.choice(list(PALETTES.keys()))

    color_hint = ""
    if theme_colors and isinstance(theme_colors, dict):
        primary_name   = str(theme_colors.get('primary', '') or '').strip()
        secondary_name = str(theme_colors.get('secondary', '') or '').strip()
        primary_name, secondary_name = _prioritize_user_theme(primary_name, secondary_name)
        primary_hex    = _resolve_theme_color(primary_name)
        secondary_hex  = _resolve_theme_color(secondary_name)
        extra_names    = theme_colors.get("accent_candidates", []) if isinstance(theme_colors.get("accent_candidates", []), list) else []
        primary_name   = str(theme_colors.get("primary", "") or "").strip()
        secondary_name = str(theme_colors.get("secondary", "") or "").strip()
        primary_name, secondary_name = _prioritize_user_theme(primary_name, secondary_name)
        primary_str    = primary_hex or primary_name or '(not specified)'
        secondary_str  = secondary_hex or secondary_name or '(not specified)'
        if primary_hex or secondary_hex:
            color_hint = f"""
=== CUSTOM COLOR THEME ===
Primary color: {primary_str}  (name: {primary_name})
Secondary color: {secondary_str}  (name: {secondary_name})
Rules:
1. Set "primary" in design_system.theme to {primary_str}.
2. Set "secondary" in design_system.theme to {secondary_str}.
3. Derive "accent", "accent2", and "bg_dark" from the primary color.
4. Ensure text contrast is readable on all backgrounds.
"""

    prompt = f"""
You are an expert presentation strategist.
Create a {tone} PowerPoint deck on: "{topic}".
Return ONLY valid JSON. No markdown fences. No commentary.
Generate exactly {target} slides.

{color_hint}

=== REQUIRED TOP-LEVEL STRUCTURE ===
{{
  "design_system": {{
    "theme": {{
      "primary": "#RRGGBB",
      "secondary": "#RRGGBB",
      "accent": "#RRGGBB",
      "accent2": "#RRGGBB",
      "bg_dark": "#RRGGBB",
      "bg_light": "#FFFFFF",
      "text_dark": "#RRGGBB",
      "text_light": "#FFFFFF",
      "card_bg": "#FFFFFF",
      "header_font": "font name",
      "body_font": "font name"
    }}
  }},
  "slides": [ ... exactly {target} slide objects ... ]
}}

=== SLIDE RULES ===
Slide 1: layout = "title_cover"
Slide 2: layout = "section_index"
Slides 3-N: choose the BEST layout for each content type.
- Avoid repeating the same layout in consecutive slides.
- Use concrete facts, real companies, real numbers.
- Generate smart, concise slide titles. Never copy the user's raw prompt verbatim.

=== EVERY SLIDE MUST HAVE ===
title, subtitle, layout, icon (single char ▸ ◆ ✓), content (list of strings), style object.

style object fields (all required):
  pattern_name, surface ("light"|"tint"),
  header_variant ("solid"|"split"|"banded"),
  card_variant ("outline"|"soft"|"banded"),
  footer_variant ("solid"|"line"),
  badge_shape ("oval"|"rect"),
  accent_rotation ("static"|"auto")

=== LAYOUT-SPECIFIC REQUIRED FIELDS ===

"title_cover":
  content: [2-3 short highlight strings]

"section_index":
  sections: [4-7 short section title strings]
  content: same as sections

"bullets":
  content: [5-6 substantive bullet strings]

"two_column":
  left_title, right_title,
  left_points: [4 strings], right_points: [4 strings], content: []

"big_stat":
  stat, stat_label, stat_source,
  content: [4-5 supporting strings]

"timeline":
  steps: [{{"label":"Phase","detail":"1-sentence description"}}, ...4 items]
  content: []

"icon_grid":
  grid_items: [{{"icon":"A","title":"2-3 words","detail":"1 sentence"}}, ...4 items]
  content: []

"case_study":
  company, result,
  metrics: [{{"label":"name","value":75}}, ...2-3 items],
  content: [3 strings]

"table":
  table_columns: [3-5 strings], table_rows: [[...], ...4-5 rows], content: []

"chart":
  chart_title, chart_data: [{{"label":"name","value":75}}, ...4-5 items],
  chart_source, content: []

"hybrid_insight":
  stat, stat_label,
  chart_data: [{{"label":"name","value":75}}, ...3 items],
  content: [4-5 strings]

=== CONTENT FORMAT — ABSOLUTE RULES ===
Every string in content, left_points, right_points, sections, step details MUST be:
- Plain English only. Start with a capital letter or digit.
- 5 to 20 words long.
- NO icon names of any kind: never write check-circle, check_circle, star, arrow,
  bullet, circle, heart, flag, check, trending_up, handshake, thank_you, or ANY
  other icon/UI keyword — not even as a prefix.
- NO markdown syntax: no **bold**, no *italic*, no `code`.
- NO bullet prefixes: no •, -, –, →, ✓, ▸ at the start of strings.
- NO emoji characters.

BAD (never generate):
  "**check-circle** AI is transforming industries"
  "check-circle AI continues to transform..."
  "• Strategic investment is key"
  "arrow Collaboration drives growth"

GOOD (always generate):
  "AI is transforming industries with unprecedented speed and impact."
  "Strategic investment in AI technologies is crucial for competitive advantage."
  "Collaboration across sectors will drive sustainable growth."

=== FORMATTING SAFETY RULES ===
- Keep text concise enough to fit on a slide.
- Never generate duplicate bullets.
- Keep metric labels 2-4 words max.
- Output ONLY valid JSON. No markdown fences. No commentary. No extra keys.
"""

    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a JSON-only presentation content generator. "
                    "You output ONLY valid JSON — no markdown, no fences, no commentary. "
                    "Every text value must be plain English prose starting with a capital letter. "
                    "You MUST NEVER include icon names (check-circle, check_circle, star, arrow, "
                    "trending_up, handshake, thank_you, bullet, circle, heart, flag, etc.) "
                    "or markdown (**bold**, *italic*) anywhere in your output. "
                    "No emoji. No bullet symbol prefixes."
                )
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.6,
        max_tokens=7000,
    )
    usage_total = _usage_dict(resp.usage)
    raw         = (resp.choices[0].message.content or "").strip()
    data        = _parse_json(raw)

    if not isinstance(data, dict): data = {}
    data.setdefault("design_system", {})
    data.setdefault("slides", [])

    # Override theme colors (highest priority)
    if theme_colors and isinstance(theme_colors, dict):
        design = data["design_system"].setdefault("theme", {})
        data["design_system"]["theme_meta"] = {"source": "user"}
        primary_name   = str(theme_colors.get("primary", "") or "").strip()
        secondary_name = str(theme_colors.get("secondary", "") or "").strip()
        primary_hex    = _resolve_theme_color(primary_name)
        secondary_hex  = _resolve_theme_color(secondary_name)

        if primary_hex:
            design["primary"] = primary_hex
            design["accent"] = _mix_hex(primary_hex, "#000000", 0.16)
            design["bg_dark"] = _mix_hex(primary_hex, "#000000", 0.45)
        if secondary_hex:
            design["secondary"] = secondary_hex
            design["accent2"] = _mix_hex(secondary_hex, "#FFFFFF", 0.20)
        elif primary_hex:
            design["secondary"] = _mix_hex(primary_hex, "#FFFFFF", 0.28)
            design["accent2"] = _mix_hex(primary_hex, "#FFFFFF", 0.18)

        if secondary_name.lower() in ("white", "light") or primary_name.lower() in ("white", "light"):
            design["bg_light"] = "#FFFFFF"
            design["card_bg"] = "#FFFFFF"
        if extra_names:
            design["accent_candidates"] = [
                _resolve_theme_color(name) for name in extra_names
                if _resolve_theme_color(name)
            ][:3]

    slides     = [s for s in data["slides"] if isinstance(s,dict)]
    normalized = []
    for i, s in enumerate(slides, start=1):
        ns = _normalize_slide(s, i, topic)
        normalized.append(ns)

    if normalized:
        normalized[0]["layout"] = "title_cover"
    if len(normalized) > 1 and _canon(normalized[1].get("layout","")) != "section_index":
        normalized.insert(1, {
            "title":"Contents","subtitle":f"{tone} overview",
            "layout":"section_index","icon":"▸",
            "sections":[s.get("title","") for s in normalized[2:] if s.get("title")][:6],
            "content":[], "style":{}
        })
    if len(normalized) > 1:
        normalized[1]["layout"]   = "section_index"
        secs = [s.get("title","") for s in normalized[2:] if _safe_str(s.get("title",""))][:6]
        normalized[1]["sections"] = secs
        normalized[1]["content"]  = secs

    # Clean thank-you slide
    if normalized and normalized[-1].get("title", "").lower() in ["thank you", "thanks", "q&a"]:
        normalized[-1] = {
            "title": "Thank You", "subtitle": "Questions?",
            "layout": "title_cover", "icon": "▸",
            "content": [], "style": {}
        }
    normalized = normalized[:target]

    for i, slide in enumerate(normalized):
        if any(w in slide.get("title", "").lower() for w in ["thank", "thanks", "q&a", "questions"]):
            slide.update({"title": "Thank You", "subtitle": "Questions?",
                          "layout": "title_cover", "content": []})
            break

    # Convert temporal trend bullets → chart
    for slide in normalized:
        if slide.get("layout") == "bullets":
            content = slide.get("content", [])
            pattern = re.compile(r'^\s*(\d{4})\s*[:：]\s*(\d+(?:\.\d+)?)%?\s*$')
            matches = [(m.group(1), m.group(2)) for line in content for m in [pattern.match(line.strip())] if m]
            if len(matches) >= 3:
                slide["layout"]      = "chart"
                slide["chart_title"] = slide.get("title", "Trend Over Years")
                slide["chart_data"]  = [{"label": y, "value": int(float(v))} for y, v in matches]
                slide["content"]     = []

    repaired = []
    for s in normalized:
        if _needs_repair(s):
            s, repair_usage = _repair_slide(s, topic)
            usage_total = _merge_usage(usage_total, repair_usage)
            s = _normalize_slide(s, 0, topic)
        repaired.append(s)

    data["slides"] = repaired
    data["usage"]  = usage_total
    return data

def _derive_sections(slides):
    return [_safe_str(s.get("title","")) for s in slides
            if _canon(s.get("layout","")) not in ("title_cover","section_index")
            and _safe_str(s.get("title",""))][:6]

# ═══════════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════

def create_ppt(slide_data, topic, logo_path=None, tone="Professional", content_image_path=None):
    payload = slide_data if isinstance(slide_data, dict) else {"slides": slide_data}
    design  = payload.get("design_system", {}) if isinstance(payload, dict) else {}
    slides  = payload.get("slides", [])
    if not isinstance(slides, list): slides = []

    global _LOGO_PATH
    _LOGO_PATH = logo_path

    theme_meta = design.get("theme_meta", {}) if isinstance(design, dict) else {}
    if isinstance(theme_meta, dict) and theme_meta.get("source") == "user":
        palette_name = "classic"
    else:
        palette_name = random.choice(list(PALETTES.keys()))
    theme        = _build_theme(palette_name, design, logo_path)
    profile      = PROFILES[palette_name]

    prs = Presentation()
    prs.slide_width  = _IN(SW)
    prs.slide_height = _IN(SH)

    for i, spec in enumerate(slides, start=1):
        if not isinstance(spec, dict): continue
        spec     = _normalize_slide(spec, i, topic)
        layout   = _canon(spec.get("layout","bullets"))
        renderer = _RENDERERS.get(layout, render_bullets)
        pslide   = prs.slides.add_slide(prs.slide_layouts[6])
        renderer(pslide, spec, i, theme, profile, logo_path)

    if len(prs.slides) == 0:
        pslide = prs.slides.add_slide(prs.slide_layouts[6])
        render_bullets(pslide, {
            "title": topic, "subtitle": "Auto-generated",
            "layout":"bullets","icon":"▸",
            "content": [f"No slide data was returned for: {topic}. Please try again."],
        }, 1, theme, profile, logo_path)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
