"""
app.py
======

Frontend conversational interface for the AI PPT Generator system.

Built using Streamlit, this module provides a fully interactive,
chat-based experience for creating, editing, and managing presentations.

──────────────────────────────────────────────────────────────

💬 Core Features

1. Conversational PPT Creation

   * Users can create presentations using natural language
   * Example: "Create a PPT on Generative AI with 5 slides"

2. Multi-Turn Chat Intelligence

   * Maintains session memory
   * Supports follow-up edits:

     * "Edit slide 3"
     * "Add more points"
     * "Change tone to formal"

3. Slot-Filling Workflow

   * Dynamically collects required inputs:

     * topic
     * number of slides
     * tone
   * Avoids hardcoded flows

4. Live PPT Preview

   * Displays slides inside chat UI
   * Shows structured content before final generation

5. PPT Editing System

   * Users can modify:

     * slide titles
     * bullet points
     * layouts (chart, timeline, etc.)
   * Changes reflected instantly

6. PPT History & Switching

   * Maintains multiple generated presentations
   * Allows switching between them using semantic search

──────────────────────────────────────────────────────────────

🧹 Content Cleaning Layer

To ensure clean output:

* Removes icon tokens (check-circle, handshake, etc.)
* Strips markdown and emojis
* Ensures only plain text is displayed in UI

Functions:

* clean_icon_tokens()
* flatten_slide_content()

──────────────────────────────────────────────────────────────

🔗 Backend Integration

Communicates with FastAPI backend:

* /generate-outline → Generates slide structure
* /build-ppt       → Generates final PPT file

Handles:

* API requests
* Response parsing
* Error handling

──────────────────────────────────────────────────────────────

📦 Session State Management

Uses Streamlit session_state to manage:

* chat history
* current presentation
* slide data
* token usage
* user inputs

──────────────────────────────────────────────────────────────

🎯 Design Principle

The UI follows a strict conversational model:

```
User Input → Intent Detection → Slot Filling → API Call → Preview → Edit → Download
```

No static flows — everything is dynamic and context-aware.

──────────────────────────────────────────────────────────────

🚀 Key Functions

* init_state()
  Initializes session variables

* flatten_slide_content()
  Converts structured content into displayable bullets

* render_inline_preview()
  Displays slides in chat interface

* save_editor_content()
  Handles user edits to slides

──────────────────────────────────────────────────────────────

🎯 Goal

To provide a seamless, intelligent, and interactive experience for:

* creating presentations
* editing slides
* previewing content
* downloading final PPT

All through natural conversation.

──────────────────────────────────────────────────────────────
"""

import base64
import copy
import os
import ast
import json
import re
from uuid import uuid4
from typing import Dict, List, Optional, Any, Tuple
from openai import AzureOpenAI
import requests
import streamlit as st
from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)
from backend.services.ppt_service import generate_slide_content

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

OUTLINE_URL = "http://127.0.0.1:9000/generate-outline"
BUILD_URL   = "http://127.0.0.1:9000/build-ppt"

st.set_page_config(page_title="AI PPT Chat Builder", page_icon="💬", layout="wide")

# ------------------------------------------------------------------------------
#  GLOBAL ICON / MARKDOWN CLEANING (applies everywhere)
# ------------------------------------------------------------------------------
def clean_icon_tokens(text: str) -> str:
    """
    Remove all known icon keywords, markdown bold, emojis, and leading symbols.
    Returns cleaned text or empty string.
    """
    if not text or not isinstance(text, str):
        return ""
    t = str(text)
    # Remove **any-word** patterns (including **check-circle**)
    t = re.sub(r'\*\*.*?\*\*', '', t)
    # Remove known icon words (with underscores, hyphens, spaces)
    icon_pattern = r'\b(check[-_ ]?circle|handshake|thank_you|trending_up|trending_down|arrow|star|heart|flag|bullet|check|circle)\b'
    t = re.sub(icon_pattern, '', t, flags=re.IGNORECASE)
    # Remove :emoji_name:
    t = re.sub(r':[a-z_\-]{2,30}:', '', t)
    # Remove actual emoji characters
    t = re.sub(r'[\U0001F300-\U0001F6FF]', '', t)
    # Remove leading symbols (bullets, arrows, etc.)
    t = re.sub(r'^[\-\–\—\•\▸\▹\►\→\✓\✔\s]+', '', t)
    # Normalise spaces
    t = re.sub(r'\s+', ' ', t).strip()
    return t


_SAFE_BULLET_ICONS = {"▸", "◆", "✓", "•", "●", "▪", "‣", "→", "➜", "➤", "▶", "►"}


def clean_bullet_icon(icon: str, fallback: str = "▸") -> str:
    """Keep bullet icons to a safe glyph set so text tokens never render literally."""
    token = str(icon).strip() if icon is not None else ""
    if not token:
        return fallback
    if token in _SAFE_BULLET_ICONS:
        return token
    if len(token) == 1 and not token.isalnum() and not token.isspace():
        return token
    return fallback

# ------------------------------------------------------------------------------
#  COLOR EXTRACTION UTILITIES
# ------------------------------------------------------------------------------
_COLOR_NAMES = {
    "red", "blue", "green", "yellow", "orange", "purple", "pink",
    "cyan", "teal", "white", "black", "gray", "grey", "navy", "gold",
    "silver", "brown", "indigo", "violet", "magenta", "lime",
    "dark blue", "light blue", "dark green", "light green",
    "dark red", "light red", "mint", "lavender", "coral"
}

def extract_theme_colors(text: str) -> dict:
    """Extract primary/secondary color names from natural language."""
    if not text:
        return {}
    lower = text.lower()
    colors = {}
    # Pattern: "X and Y theme/color"
    m = re.search(
        r'(\b(?:dark|light)?\s*\w+)\s+and\s+(\b(?:dark|light)?\s*\w+)\s*(?:theme|color|colors|palette)',
        lower
    )
    if m:
        c1, c2 = m.group(1).strip(), m.group(2).strip()
        if c1 in _COLOR_NAMES:
            colors["primary"] = c1
        if c2 in _COLOR_NAMES:
            colors["secondary"] = c2
    if not colors:
        m = re.search(r'(\b(?:dark|light)?\s*\w+)\s+(?:theme|color|colors)', lower)
        if m:
            c = m.group(1).strip()
            if c in _COLOR_NAMES:
                colors["primary"] = c
    if not colors:
        m = re.search(r'with\s+(\w+)\s+and\s+(\w+)', lower)
        if m:
            c1, c2 = m.group(1), m.group(2)
            if c1 in _COLOR_NAMES:
                colors["primary"] = c1
            if c2 in _COLOR_NAMES:
                colors["secondary"] = c2
    return colors

# ------------------------------------------------------------------------------
#  JSON CONTENT FLATTENING WITH CLEANING
# ------------------------------------------------------------------------------
def _flatten_content_item(item) -> list:
    """Recursively flatten a content item and clean icon tokens."""
    if item is None:
        return []
    if isinstance(item, str):
        s = item.strip()
        if not s:
            return []
        if s.startswith("{") or s.startswith("["):
            try:
                parsed = json.loads(s)
                return _flatten_content_item(parsed)
            except Exception:
                pass
            try:
                parsed = ast.literal_eval(s)
                return _flatten_content_item(parsed)
            except Exception:
                pass
        cleaned = clean_icon_tokens(s)
        return [cleaned] if cleaned else []
    if isinstance(item, dict):
        results = []
        header = item.get("header") or item.get("title") or item.get("label") or ""
        points = item.get("points") or item.get("items") or item.get("content") or []
        detail = item.get("detail") or item.get("description") or item.get("text") or ""
        if header and points:
            cleaned_header = clean_icon_tokens(header)
            if cleaned_header:
                results.append(cleaned_header)
            for p in (points if isinstance(points, list) else [points]):
                results.extend(_flatten_content_item(p))
        elif header and detail:
            combined = f"{header}: {detail}"
            cleaned = clean_icon_tokens(combined)
            if cleaned:
                results.append(cleaned)
        elif header:
            cleaned = clean_icon_tokens(header)
            if cleaned:
                results.append(cleaned)
        elif detail:
            cleaned = clean_icon_tokens(detail)
            if cleaned:
                results.append(cleaned)
        elif points:
            for p in (points if isinstance(points, list) else [points]):
                results.extend(_flatten_content_item(p))
        else:
            for v in item.values():
                if v and isinstance(v, (str, int, float)):
                    results.append(clean_icon_tokens(str(v)))
        return [r for r in results if r]
    if isinstance(item, list):
        results = []
        for sub in item:
            results.extend(_flatten_content_item(sub))
        return results
    return [clean_icon_tokens(str(item))] if item else []

def flatten_slide_content(content) -> list:
    """Convert any content value into a clean list of plain bullet strings."""
    if content is None:
        return []
    if isinstance(content, str):
        stripped = content.strip()
        if stripped.startswith("[") or stripped.startswith("{"):
            try:
                parsed = json.loads(stripped)
                return flatten_slide_content(parsed)
            except Exception:
                pass
            try:
                parsed = ast.literal_eval(stripped)
                return flatten_slide_content(parsed)
            except Exception:
                pass
        lines = [clean_icon_tokens(l) for l in stripped.splitlines() if clean_icon_tokens(l)]
        return lines if lines else ([clean_icon_tokens(stripped)] if stripped else [])
    if isinstance(content, list):
        result = []
        for item in content:
            result.extend(_flatten_content_item(item))
        return [clean_icon_tokens(r) for r in result if clean_icon_tokens(r)]
    if isinstance(content, dict):
        return flatten_slide_content(list(content.values()))
    return [clean_icon_tokens(str(content))] if content else []

# ------------------------------------------------------------------------------
#  STATE INIT
# ------------------------------------------------------------------------------
def init_state():
    defaults = {
        "messages": [{
            "role": "assistant",
            "content": (
                "👋 Hi! Tell me the **topic** for your presentation and I'll get started.\n\n"
                "Example: *'Create a deck on Generative AI in Healthcare'*"
            ),
        }],
        "outline_payload": None,
        "ppt_bytes":       None,
        "ppt_filename":    None,
        "topic":           "",
        "tone":            "Professional",
        "num_slides":      6,
        "slide_count":     6,
        "logo_bytes":      None,
        "logo_name":       None,
        "sections":        "",
        "ppt_history":     [],
        "current_ppt_id":  None,
        "current_step":    "topic",
        "asked_steps":     [],
        "token_usage":     {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
        "last_usage":      None,
        "session_id":      uuid4().hex,
        "last_question":   None,
        "pending_intent":  None,
        "active_slide_index": None,
        "last_build_error": None,
        "session_memory":  {"ppts": [], "current_ppt": None},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ------------------------------------------------------------------------------
#  HELPER FUNCTIONS (slide utilities, history, etc.)
# ------------------------------------------------------------------------------
def add_message(role, content):
    st.session_state.messages.append({"role": role, "content": content})

def add_message_with_preview(role: str, text: str, preview_slides: list = None,
                              ppt_label: str = "", ppt_bytes: bytes = None,
                              ppt_filename: str = ""):
    entry = {
        "role": role,
        "content": text,
    }
    if preview_slides:
        entry["preview_slides"] = preview_slides
        entry["ppt_label"] = ppt_label
    if ppt_bytes:
        entry["ppt_bytes"] = ppt_bytes
        entry["ppt_filename"] = ppt_filename
    st.session_state.messages.append(entry)

def extract_chat_input(value):
    if value is None:
        return None, []
    if isinstance(value, str):
        return value, []
    text = getattr(value, "text", None)
    files = getattr(value, "files", None) or []
    return text, files

def usage_to_dict(usage):
    if not usage:
        return None
    return {
        "prompt_tokens":     int(getattr(usage, "prompt_tokens", 0) or 0),
        "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
        "total_tokens":      int(getattr(usage, "total_tokens", 0) or 0),
    }

def record_usage(kind: str, usage: dict):
    if not usage:
        return
    totals = st.session_state.token_usage
    for k in ("prompt_tokens", "completion_tokens", "total_tokens"):
        totals[k] += int(usage.get(k, 0) or 0)
    st.session_state.last_usage = {"kind": kind, **usage}

def ensure_editor_id(slide: dict) -> dict:
    slide = dict(slide or {})
    slide.setdefault("_editor_id", uuid4().hex)
    slide.setdefault("title", "Untitled Slide")
    slide.setdefault("subtitle", "")
    slide.setdefault("layout", "bullets")
    slide["icon"] = clean_bullet_icon(slide.get("icon", "▸"))
    slide.setdefault("content", [])
    slide.setdefault("style", {})
    if not isinstance(slide["content"], list):
        slide["content"] = flatten_slide_content(slide["content"])
    else:
        slide["content"] = flatten_slide_content(slide["content"])
    if not isinstance(slide["style"], dict):
        slide["style"] = {}
    return slide

def normalize_slide(slide: dict) -> dict:
    return ensure_editor_id(slide or {})

def reset_editor_widget_state():
    keys = [k for k in st.session_state.keys() if re.match(r"^[tslc]_[0-9a-f]{32}$", k)]
    for key in keys:
        st.session_state.pop(key, None)

def prime_editor_widget_state(slide: dict, force: bool = False):
    sid = slide.get("_editor_id")
    if not sid:
        return
    values = {
        f"t_{sid}": slide.get("title", ""),
        f"s_{sid}": slide.get("subtitle", ""),
        f"l_{sid}": slide.get("layout", "bullets"),
        f"c_{sid}": get_editor_content(slide),
    }
    for key, value in values.items():
        if force or key not in st.session_state:
            st.session_state[key] = value

def sync_all_editor_widgets(slides: list):
    reset_editor_widget_state()
    for slide in slides:
        if isinstance(slide, dict):
            prime_editor_widget_state(slide, force=True)

def bullets_to_text(items) -> str:
    if not isinstance(items, list):
        return ""
    return "\n".join(str(x).strip() for x in items if str(x).strip())

def text_to_bullets(value: str) -> list:
    bullets = []
    for line in str(value or "").splitlines():
        c = line.strip().lstrip("-").lstrip("•").strip()
        if c:
            bullets.append(c)
    return bullets

def seq_icon(idx: int) -> str:
    if 0 <= idx < 26:
        return chr(ord("A") + idx)
    return str(idx + 1)

def normalize_icon_grid_items(slide: dict, text: str = "") -> list:
    items = slide.get("grid_items", []) if isinstance(slide, dict) else []
    cleaned = []
    if isinstance(items, list):
        for i, item in enumerate(items):
            if not isinstance(item, dict):
                continue
            title = str(item.get("title", "")).strip()
            detail = str(item.get("detail", item.get("description", ""))).strip()
            if not title:
                continue
            icon = str(item.get("icon", "")).strip() or seq_icon(i)
            cleaned.append({"icon": icon, "title": title, "detail": detail})
    if cleaned:
        return cleaned[:4]
    source = str(text or "").strip()
    parsed_items = None
    if source:
        try:
            parsed = json.loads(source)
            if isinstance(parsed, dict) and isinstance(parsed.get("grid_items"), list):
                parsed_items = parsed["grid_items"]
            elif isinstance(parsed, list):
                parsed_items = parsed
        except Exception:
            try:
                parsed = ast.literal_eval(source)
                if isinstance(parsed, dict) and isinstance(parsed.get("grid_items"), list):
                    parsed_items = parsed["grid_items"]
                elif isinstance(parsed, list):
                    parsed_items = parsed
            except Exception:
                parsed_items = None
    if isinstance(parsed_items, list):
        for i, item in enumerate(parsed_items):
            if not isinstance(item, dict):
                continue
            title = str(item.get("title", "")).strip()
            detail = str(item.get("detail", item.get("description", ""))).strip()
            if not title:
                continue
            icon = str(item.get("icon", "")).strip() or seq_icon(i)
            cleaned.append({"icon": icon, "title": title, "detail": detail})
        if cleaned:
            return cleaned[:4]
    for i, line in enumerate(text_to_bullets(source)):
        if len(cleaned) >= 4:
            break
        parsed_line = None
        try:
            parsed_line = json.loads(line)
        except Exception:
            try:
                parsed_line = ast.literal_eval(line)
            except Exception:
                parsed_line = None
        if isinstance(parsed_line, dict):
            title = str(parsed_line.get("title", "")).strip()
            detail = str(parsed_line.get("detail", parsed_line.get("description", ""))).strip()
            icon = str(parsed_line.get("icon", "")).strip() or seq_icon(i)
            if title:
                cleaned.append({"icon": icon, "title": title, "detail": detail})
                continue
        if ":" in line:
            title, detail = [p.strip() for p in line.split(":", 1)]
        else:
            title, detail = line[:30].strip(), line.strip()
        if title:
            cleaned.append({"icon": seq_icon(i), "title": title, "detail": detail})
    return cleaned[:4]

def get_editor_content(slide: dict) -> str:
    layout = slide.get("layout", "bullets")
    if layout == "two_column":
        lp = slide.get("left_points", []) or []
        rp = slide.get("right_points", []) or []
        combined = list(lp) + list(rp)
        if combined:
            return bullets_to_text(combined)
    elif layout == "timeline":
        steps = slide.get("steps", []) or []
        if steps:
            return "\n".join(
                f"{s.get('label', ''):<20}{s.get('detail', '')}"
                for s in steps if isinstance(s, dict)
            )
    elif layout == "icon_grid":
        items = slide.get("grid_items", []) or []
        if items:
            return "\n".join(
                f"{g.get('title', '')}: {g.get('detail', '')}"
                for g in items if isinstance(g, dict)
            )
    elif layout == "table":
        cols = slide.get("table_columns", []) or []
        rows = slide.get("table_rows", []) or []
        if cols:
            lines = [" | ".join(str(c) for c in cols)]
            for row in rows:
                lines.append(" | ".join(str(v) for v in row))
            return "\n".join(lines)
    elif layout == "chart":
        data = slide.get("chart_data", []) or []
        if data:
            return "\n".join(
                f"{d.get('label', '')} : {d.get('value', '')}"
                for d in data if isinstance(d, dict)
            )
    elif layout == "case_study":
        lines = []
        if slide.get("company"):
            lines.append(f"Company: {slide['company']}")
        if slide.get("result"):
            lines.append(f"Result: {slide['result']}")
        for m in (slide.get("metrics") or []):
            if isinstance(m, dict):
                lines.append(f"  {m.get('label', '')}: {m.get('value', '')}")
        lines.extend(slide.get("content", []) or [])
        return "\n".join(lines)
    elif layout in ("big_stat", "hybrid_insight"):
        lines = []
        if slide.get("stat"):
            lines.append(f"STAT: {slide['stat']}")
        if slide.get("stat_label"):
            lines.append(f"LABEL: {slide['stat_label']}")
        lines.extend(slide.get("content", []) or [])
        return "\n".join(lines)
    elif layout == "section_index":
        secs = slide.get("sections", []) or slide.get("content", []) or []
        return bullets_to_text(secs)
    return bullets_to_text(flatten_slide_content(slide.get("content", [])))

def save_editor_content(slide: dict, text: str) -> dict:
    layout = slide.get("layout", "bullets")
    if layout in ("section_index", "title_cover") and len(text.strip()) > 10:
        if "thank" not in slide.get("title", "").lower():
            slide["layout"] = "bullets"
            layout = "bullets"
    lines = text_to_bullets(text)
    if layout == "two_column":
        mid = max(1, len(lines) // 2)
        slide["left_points"] = lines[:mid]
        slide["right_points"] = lines[mid:]
        slide["content"] = lines
    elif layout == "timeline":
        steps = []
        for line in lines:
            if ":" in line:
                parts = line.split(":", 1)
                steps.append({"label": parts[0].strip(), "detail": parts[1].strip()})
            elif len(line) > 25:
                steps.append({"label": line[:20].strip(), "detail": line[20:].strip()})
            else:
                steps.append({"label": line, "detail": ""})
        slide["steps"] = steps
        slide["content"] = lines
    elif layout == "icon_grid":
        slide["grid_items"] = normalize_icon_grid_items(slide, text)
        slide["content"] = [
            f"{item['title']}: {item['detail']}".rstrip(": ").strip()
            for item in slide["grid_items"]
        ]
    elif layout == "table":
        if lines:
            slide["table_columns"] = [c.strip() for c in lines[0].split("|") if c.strip()]
            slide["table_rows"] = [[c.strip() for c in row.split("|")] for row in lines[1:]]
        slide["content"] = lines
    elif layout == "chart":
        data = []
        for line in lines:
            if ":" in line:
                p = line.split(":", 1)
                try:
                    data.append({"label": p[0].strip(), "value": int(float(p[1].strip()))})
                except Exception:
                    data.append({"label": p[0].strip(), "value": 50})
            else:
                data.append({"label": line, "value": 50})
        slide["chart_data"] = data[:5]
        slide["content"] = lines
    elif layout == "case_study":
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("company:"):
                slide["company"] = line.split(":", 1)[1].strip()
            elif ll.startswith("result:"):
                slide["result"] = line.split(":", 1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines
    elif layout in ("big_stat", "hybrid_insight"):
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("stat:"):
                slide["stat"] = line.split(":", 1)[1].strip()
            elif ll.startswith("label:"):
                slide["stat_label"] = line.split(":", 1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines
    elif layout == "section_index":
        slide["sections"] = lines
        slide["content"] = lines
    else:
        slide["content"] = lines
    return slide

# ------------------------------------------------------------------------------
#  CHAT-INLINE PREVIEW RENDERER (simplified – content already cleaned)
# ------------------------------------------------------------------------------
def render_inline_preview(slides: list, label: str = "", ppt_bytes: bytes = None,
                           ppt_filename: str = "presentation.pptx", key_suffix: str = ""):
    if not slides:
        st.caption("No slides to preview.")
        return
    header = f"📊 **{label}**" if label else "📊 **Presentation Preview**"
    st.markdown(header)
    st.caption(f"{len(slides)} slides")
    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        title = slide.get("title", f"Slide {idx}")
        layout = slide.get("layout", "bullets")
        subtitle = slide.get("subtitle", "")
        with st.expander(f"Slide {idx}: {title}  `{layout}`", expanded=(idx == 1)):
            if subtitle:
                st.markdown(f"*{subtitle}*")
            content = flatten_slide_content(slide.get("content", []))
            if content:
                for item in content:
                    st.markdown(f"- {item}")
            if layout == "two_column":
                col1, col2 = st.columns(2)
                with col1:
                    st.caption(f"**{slide.get('left_title', 'Left')}**")
                    for pt in flatten_slide_content(slide.get("left_points", []))[:4]:
                        st.markdown(f"- {pt}")
                with col2:
                    st.caption(f"**{slide.get('right_title', 'Right')}**")
                    for pt in flatten_slide_content(slide.get("right_points", []))[:4]:
                        st.markdown(f"- {pt}")
            elif layout == "big_stat":
                st.metric(
                    label=clean_icon_tokens(slide.get("stat_label", "")),
                    value=clean_icon_tokens(slide.get("stat", "—"))
                )
            elif layout == "timeline":
                steps = slide.get("steps", []) or []
                for s in steps[:5]:
                    if isinstance(s, dict):
                        label_clean = clean_icon_tokens(s.get("label", ""))
                        detail_clean = clean_icon_tokens(s.get("detail", ""))
                        if label_clean or detail_clean:
                            st.markdown(f"**{label_clean}** — {detail_clean}")
            elif layout == "icon_grid":
                items = slide.get("grid_items", []) or []
                for g in items[:4]:
                    if isinstance(g, dict):
                        title_clean = clean_icon_tokens(g.get("title", ""))
                        detail_clean = clean_icon_tokens(g.get("detail", ""))
                        if title_clean or detail_clean:
                            st.markdown(f"{g.get('icon', '•')} **{title_clean}** — {detail_clean}")
            elif layout == "chart":
                data = slide.get("chart_data", []) or []
                chart_title = clean_icon_tokens(slide.get("chart_title", slide.get("title", "")))
                for d in data[:5]:
                    if isinstance(d, dict):
                        lbl = clean_icon_tokens(d.get("label", ""))
                        val = d.get("value", 0)
                        if lbl:
                            st.markdown(f"- {lbl}: **{val}**")
                if chart_title:
                    st.caption(f"Chart: {chart_title}")
                extra = flatten_slide_content(slide.get("content", []))
                if extra:
                    st.markdown("**Key Takeaways**")
                    for item in extra[:4]:
                        st.markdown(f"- {item}")
            elif layout == "case_study":
                company = clean_icon_tokens(slide.get("company", ""))
                result = clean_icon_tokens(slide.get("result", ""))
                if company:
                    st.markdown(f"🏢 **{company}**")
                if result:
                    st.markdown(f"✅ {result}")
    if ppt_bytes is not None:
        dl_key = f"dl_{key_suffix}" if key_suffix else f"dl_{abs(hash(ppt_filename + str(len(slides))))}"
        st.download_button(
            label="⬇️ Download PPT",
            data=ppt_bytes,
            file_name=ppt_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary",
            key=dl_key,
        )

def render_message(msg: dict, msg_index: int = 0):
    role = msg["role"]
    with st.chat_message(role):
        st.markdown(msg["content"])
        if msg.get("preview_slides"):
            render_inline_preview(
                slides=msg["preview_slides"],
                label=msg.get("ppt_label", ""),
                ppt_bytes=msg.get("ppt_bytes"),
                ppt_filename=msg.get("ppt_filename", "presentation.pptx"),
                key_suffix=f"msg{msg_index}",
            )

# ------------------------------------------------------------------------------
#  PPT HISTORY & SEMANTIC SEARCH
# ------------------------------------------------------------------------------
def next_ppt_id() -> str:
    return f"ppt_{len(st.session_state.get('ppt_history', [])) + 1}"

def _upsert_ppt_history(ppt_id: str, **fields):
    history = st.session_state.setdefault("ppt_history", [])
    for item in history:
        if item.get("id") == ppt_id:
            item.update(fields)
            return
    record = {"id": ppt_id}
    record.update(fields)
    history.append(record)

def get_ppt_by_id(ppt_id: str):
    for item in st.session_state.get("ppt_history", []):
        if item.get("id") == ppt_id:
            return item
    return None

def ppt_sections_preview(item: dict):
    sections = (item or {}).get("sections") or ""
    if sections:
        return sections
    outline = (item or {}).get("outline_payload") or {}
    slides = outline.get("slides", []) if isinstance(outline, dict) else []
    titles = [str(s.get("title", "")).strip() for s in slides if isinstance(s, dict)]
    preview = ", ".join([t for t in titles[2:6] if t]) if len(titles) > 2 else ", ".join([t for t in titles[:4] if t])
    return preview or "—"

def semantic_search_ppts(query: str) -> List[Dict]:
    history = st.session_state.get("ppt_history", [])
    if not history:
        return []
    ppt_summaries = []
    for ppt in history:
        summary = {
            "id": ppt.get("id"),
            "topic": ppt.get("topic", ""),
            "sections": ppt.get("sections", ""),
            "slide_titles": []
        }
        outline = ppt.get("outline_payload", {})
        slides = outline.get("slides", [])
        for s in slides[:8]:
            if isinstance(s, dict):
                summary["slide_titles"].append(s.get("title", ""))
        ppt_summaries.append(summary)
    prompt = f"""
You are a semantic search engine. Given the user query, rank the following PPTs by relevance.
Return a JSON list of objects with "id" and "score" (float 0-1) in descending order of relevance.
Only return the JSON list, no extra text.

Query: "{query}"

PPTs:
{json.dumps(ppt_summaries, indent=2)}
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=500,
        )
        raw = resp.choices[0].message.content.strip()
        match = re.search(r"\[.*\]", raw, re.DOTALL)
        if match:
            results = json.loads(match.group())
            ranked = []
            for r in results:
                if isinstance(r, dict) and "id" in r:
                    ranked.append(r)
                elif isinstance(r, str):
                    ranked.append({"id": r, "score": 1.0})
            return ranked
    except Exception:
        pass
    return []

def find_ppt_semantically(query: str, disambiguate: bool = False) -> Tuple[Optional[str], Optional[str]]:
    ranked = semantic_search_ppts(query)
    if not ranked:
        return None, None
    high_conf = [r for r in ranked if r.get("score", 0) >= 0.5]
    if not high_conf:
        high_conf = ranked[:1]
    if disambiguate and len(high_conf) > 1:
        options = []
        for r in high_conf[:3]:
            ppt = get_ppt_by_id(r["id"])
            if ppt:
                topic = ppt.get("topic", f"PPT {r['id']}")
                options.append(f"**{topic}** ({r['id']})")
        if options:
            question = "I found several presentations matching your request:\n\n" + "\n".join(f"- {opt}" for opt in options) + "\n\nWhich one did you mean?"
            return None, question
    return high_conf[0]["id"], None

def switch_active_ppt(ppt_id: str) -> bool:
    item = get_ppt_by_id(ppt_id)
    if not item:
        return False
    st.session_state.current_ppt_id = ppt_id
    st.session_state.outline_payload = item.get("outline_payload")
    st.session_state.ppt_bytes = item.get("ppt_bytes")
    st.session_state.ppt_filename = item.get("ppt_filename")
    st.session_state.topic = item.get("topic", "")
    st.session_state.sections = item.get("sections", "")
    st.session_state.num_slides = item.get("slide_count", st.session_state.num_slides)
    st.session_state.slide_count = st.session_state.num_slides
    if item.get("outline_payload"):
        sync_all_editor_widgets(item["outline_payload"].get("slides", []))
    st.session_state.pending_intent = None
    return True

def resolve_ppt_reference(user_input: str) -> Tuple[bool, Optional[str], Optional[str]]:
    override_patterns = [
        r"(?:in|to|switch to|open)\s+(?:the\s+)?(?:ppt|presentation)?\s*['\"]?([^'\"]+)['\"]?",
        r"(?:do this in|apply to)\s+(\d+(?:st|nd|rd|th)?\s*ppt)",
        r"(?:go to|select)\s+ppt\s+(\d+)",
    ]
    for pat in override_patterns:
        m = re.search(pat, user_input, re.IGNORECASE)
        if m:
            ref = m.group(1).strip()
            num_match = re.search(r"(\d+)", ref)
            if num_match:
                idx = int(num_match.group(1)) - 1
                history = st.session_state.get("ppt_history", [])
                if 0 <= idx < len(history):
                    new_id = history[idx]["id"]
                    switch_active_ppt(new_id)
                    return True, new_id, None
            best_id, _ = find_ppt_semantically(ref, disambiguate=False)
            if best_id:
                switch_active_ppt(best_id)
                return True, best_id, None
    if re.search(r"(?:which ppt|what about|show me|open|switch to|go back|previous ppt)", user_input, re.IGNORECASE):
        best_id, clarification = find_ppt_semantically(user_input, disambiguate=True)
        if clarification:
            return False, None, clarification
        if best_id and best_id != st.session_state.get("current_ppt_id"):
            switch_active_ppt(best_id)
            return True, best_id, None
    if re.search(r"\b(?:go back|previous ppt|last presentation)\b", user_input, re.IGNORECASE):
        last_ppt = st.session_state.get("last_ppt")
        if last_ppt and get_ppt_by_id(last_ppt):
            switch_active_ppt(last_ppt)
            return True, last_ppt, None
        else:
            return False, None, "No previous presentation found."
    return False, None, None

# ------------------------------------------------------------------------------
#  SLOT-FILLING CONVERSATIONAL AGENT
# ------------------------------------------------------------------------------
class ConversationalAgent:
    INTENT_SLOTS = {
        "create_ppt": ["topic", "slide_count", "sections"],
        "edit_slide": ["slide_number", "change_content"],
        "add_slide": ["slide_content", "position"],
        "delete_slide": ["slide_number"],
        "download_ppt": ["ppt_ref"],
        "preview_ppt": ["ppt_ref"],
        "ppt_info": ["info_type", "ppt_ref"],
        "suggest_topic": ["scope"],
        "greeting": [],
        "smalltalk": [],
    }
    def __init__(self):
        self.state = {
            "intent": None,
            "slots": {},
            "missing_slots": [],
            "next_question": None,
            "action": "ask",
        }
    def load_state(self):
        pending = st.session_state.get("pending_intent")
        if pending:
            self.state = pending
    def save_state(self):
        st.session_state["pending_intent"] = self.state
    def clear_state(self):
        self.state = {
            "intent": None,
            "slots": {},
            "missing_slots": [],
            "next_question": None,
            "action": "ask",
        }
        st.session_state["pending_intent"] = None
    def process(self, user_input: str, context: dict) -> Tuple[str, Optional[Dict], bool]:
        self.load_state()
        prompt = self._build_prompt(user_input, context)
        try:
            resp = _client.chat.completions.create(
                model=AZURE_DEPLOYMENT,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
                max_tokens=500,
            )
            raw = resp.choices[0].message.content.strip()
            if not raw.startswith("{"):
                m = re.search(r"\{.*\}", raw, re.DOTALL)
                if m:
                    raw = m.group()
            data = json.loads(raw)
            new_intent = data.get("intent")
            new_slots = data.get("slots", {})
            missing = data.get("missing_slots", [])
            next_q = data.get("next_question")
            action = data.get("action", "ask")
            if self.state["intent"] == new_intent:
                merged_slots = {**self.state["slots"], **new_slots}
            else:
                merged_slots = new_slots
                self.state["intent"] = new_intent
            self.state["slots"] = merged_slots
            self.state["missing_slots"] = missing
            self.state["next_question"] = next_q
            self.state["action"] = action
            if action == "ask" and missing:
                self.save_state()
                return next_q, None, False
            if action == "execute" or not missing:
                action_data = {
                    "intent": self.state["intent"],
                    "slots": self.state["slots"],
                }
                self.clear_state()
                return None, action_data, True
            else:
                self.save_state()
                return next_q or "Could you please provide more information?", None, False
        except Exception as e:
            st.error(f"Agent error: {e}")
            return "Sorry, I had trouble understanding. Could you rephrase?", None, False
    def _build_prompt(self, user_input: str, context: dict) -> str:
        current_intent = self.state.get("intent")
        current_slots = self.state.get("slots", {})
        current_missing = self.state.get("missing_slots", [])
        slides_summary = summarize_slides_for_llm(context.get("slides", []))
        has_deck = context.get("has_deck", False)
        prompt = f"""
You are a conversational AI assistant for building presentations. Your job is to fill in missing information (slots) for the user's intent.

Current conversation state:
- Intent: {current_intent if current_intent else "none"}
- Already filled slots: {json.dumps(current_slots)}
- Still missing slots: {current_missing}

User just said: "{user_input}"

Context:
- Does the user have an existing presentation? {has_deck}
- Current slide titles (if any): {json.dumps([s.get('title') for s in slides_summary])}

Your task: Update the state based on the user's input. Return a JSON object with:

{{
  "intent": "one of: create_ppt, edit_slide, add_slide, delete_slide, download_ppt, preview_ppt, ppt_info, suggest_topic, greeting, smalltalk, unknown",
  "slots": {{ ... }},
  "missing_slots": [...],
  "next_question": "string or null",
  "action": "ask" or "execute"
}}

Required slots for each intent:
- create_ppt: topic (string), slide_count (integer), sections (string, optional)
- edit_slide: slide_number (integer), change_content (string)
- add_slide: slide_content (string), position (string or integer, optional)
- delete_slide: slide_number (integer)
- download_ppt: ppt_ref (string or integer, optional)
- preview_ppt: ppt_ref (string or integer, optional)
- ppt_info: info_type (string: "count", "list", "topic"), ppt_ref (optional)
- suggest_topic: scope (string: "current", "new")
- greeting / smalltalk / unknown: no slots

Rules:
- If the user's input indicates a new intent, reset slots and set intent accordingly.
- If the user is continuing a previous intent, merge new information.
- Only mark a slot as missing if it's required and not yet filled.
- For edit_slide, if user says "edit slide 3" without change content, set slide_number=3, change_content=null, missing_slots=["change_content"].
- For edit_slide, if user says "add 2 points about AI to slide 3", set slide_number=3, change_content="add 2 points about AI", missing_slots=[].
- For add_slide, if user says "add a summary slide", set slide_content="summary", position=null, missing_slots=["position"].
- For add_slide, if user says "add a summary slide at the end", set slide_content="summary", position="end", missing_slots=[].
- For create_ppt, if user says "make a ppt on AI", set topic="AI", slide_count=null, sections=null, missing_slots=["slide_count"].
- For smalltalk like "thanks", "ok", "yes", "no", set intent="smalltalk".
- For "hi", "hello" set intent="greeting".

Be conversational: next_question should ask ONLY for the next missing slot (one at a time).

Return ONLY valid JSON. No extra text.
"""
        return prompt

# ------------------------------------------------------------------------------
#  ACTION EXECUTORS
# ------------------------------------------------------------------------------
def commit_changes(updated_slides, success_msg):
    st.session_state.outline_payload["slides"] = updated_slides
    sync_all_editor_widgets(updated_slides)
    ppt_bytes = rebuild_ppt_from_outline()
    ppt_id = st.session_state.get("current_ppt_id", "")
    topic = st.session_state.get("topic", "Presentation")
    label = f"Updated Preview — {topic} ({ppt_id})"
    add_message_with_preview(
        role="assistant",
        text=success_msg,
        preview_slides=updated_slides,
        ppt_label=label,
        ppt_bytes=ppt_bytes,
        ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
    )
    st.rerun()

def execute_action(intent: str, slots: dict, slides: list) -> Tuple[str, bool]:
    if intent in ("edit_slide", "add_slide", "delete_slide", "preview_ppt", "download_ppt"):
        if not st.session_state.get("current_ppt_id"):
            return "No presentation is currently active. Please create or switch to a presentation first.", False
        if not st.session_state.get("outline_payload"):
            return "The current presentation has no slides. Please create a presentation first.", False

    if intent == "create_ppt":
        topic = slots.get("topic")
        slide_count = slots.get("slide_count")
        sections = slots.get("sections", "")
        if not topic or not slide_count:
            return "Missing topic or slide count for creation.", False
        try:
            slide_count = int(slide_count)
        except (ValueError, TypeError):
            return "Slide count must be a number. Please provide a valid number.", False

        # Extract colors from the full original user message
        all_messages = st.session_state.get("messages", [])
        last_user_msg = next((m["content"] for m in reversed(all_messages) if m["role"] == "user"), "")
        theme_colors = extract_theme_colors(last_user_msg) or extract_theme_colors(topic)

        # Remove color phrases from topic to avoid duplication in slide titles
        color_pattern = r'\b(?:in\s+)?(?:dark\s+|light\s+)?(?:' + '|'.join(_COLOR_NAMES) + r')\s*(?:and\s+(?:dark\s+|light\s+)?(?:' + '|'.join(_COLOR_NAMES) + r'))?\s*(?:theme|color|colors|palette|scheme)?\b'
        topic_clean = re.sub(color_pattern, '', topic, flags=re.IGNORECASE).strip().strip(',').strip()
        topic_final = topic_clean if len(topic_clean) > 3 else topic

        generate_outline_and_reply(topic_final, slide_count, st.session_state.tone, sections or None, theme_colors)
        return "Generating presentation...", True

    elif intent == "edit_slide":
        slide_num = slots.get("slide_number")
        change_content = slots.get("change_content")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number.", False
        if not change_content:
            return "Missing change content for editing.", False
        if slide_num < 1 or slide_num > len(slides):
            return f"Slide {slide_num} doesn't exist. Deck has {len(slides)} slides.", False

        # SPECIAL HANDLE: "add bullet points" - use direct fallback, skip LLM
        if "add" in change_content.lower() and re.search(r'\d+\s*points?', change_content.lower()):
            num_match = re.search(r'(\d+)\s*points?', change_content.lower())
            num = int(num_match.group(1)) if num_match else 2
            slide = slides[slide_num - 1]
            if slide.get("layout") != "bullets":
                slide["layout"] = "bullets"
            prompt = f"Generate {num} short bullet points about: {slide.get('title')}. Return only as JSON list of strings."
            resp = _client.chat.completions.create(
                model=AZURE_DEPLOYMENT,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.4,
                max_tokens=300,
            )
            raw = resp.choices[0].message.content.strip()
            if raw.startswith('['):
                new_points = json.loads(raw)
            else:
                match = re.search(r'\[.*\]', raw, re.DOTALL)
                new_points = json.loads(match.group()) if match else []
            if new_points:
                # Clean the generated points immediately
                new_points = [clean_icon_tokens(pt) for pt in new_points if clean_icon_tokens(pt)]
                existing = slide.get("content", [])
                for pt in new_points:
                    if pt not in existing:
                        existing.append(pt)
                slide["content"] = existing
                updated_slides = refresh_section_index_slide(slides)
                commit_changes(updated_slides, f"✅ Added {num} bullet points to slide {slide_num}.")
                return f"Added {num} bullet points to slide {slide_num}.", True
            else:
                return "Failed to generate bullet points. Please try again.", False

        # For all other edits (remove, change, etc.), use the LLM
        edit_prompt = f"Edit slide {slide_num}: {change_content}"
        target_layout = slides[slide_num - 1].get("layout", "bullets")
        if target_layout in ("section_index", "title_cover") and "point" in change_content.lower():
            edit_prompt += " IMPORTANT: Change layout to 'bullets' and add content as bullet points."
        try:
            result, usage = interpret_edit_with_llm(edit_prompt, slides)
            record_usage("edit", usage)
            if result.get("action") == "edit":
                raw_slides = result.get("slides", [])
                # Clean content of each edited slide
                for s in raw_slides:
                    if "content" in s and isinstance(s["content"], list):
                        s["content"] = [clean_icon_tokens(item) for item in s["content"] if clean_icon_tokens(item)]
                    if "left_points" in s:
                        s["left_points"] = [clean_icon_tokens(p) for p in s["left_points"] if clean_icon_tokens(p)]
                    if "right_points" in s:
                        s["right_points"] = [clean_icon_tokens(p) for p in s["right_points"] if clean_icon_tokens(p)]
                cleaned = [ensure_editor_id(s) for s in raw_slides]
                updated_slides = refresh_section_index_slide(cleaned)
                commit_changes(updated_slides, f"✅ Slide {slide_num} updated.")
                return f"Slide {slide_num} updated.", True
            else:
                if "add" in change_content.lower() and re.search(r'\d+\s*points?', change_content.lower()):
                    num_match = re.search(r'(\d+)\s*points?', change_content.lower())
                    num = int(num_match.group(1)) if num_match else 2
                    slide = slides[slide_num - 1]
                    prompt = f"Generate {num} short bullet points about: {slide.get('title')}. Return only as JSON list of strings."
                    resp = _client.chat.completions.create(
                        model=AZURE_DEPLOYMENT,
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.4,
                        max_tokens=300,
                    )
                    raw = resp.choices[0].message.content.strip()
                    if raw.startswith('['):
                        new_points = json.loads(raw)
                    else:
                        match = re.search(r'\[.*\]', raw, re.DOTALL)
                        new_points = json.loads(match.group()) if match else []
                    if new_points:
                        new_points = [clean_icon_tokens(pt) for pt in new_points if clean_icon_tokens(pt)]
                        slide['content'].extend(new_points)
                        updated_slides = refresh_section_index_slide(slides)
                        commit_changes(updated_slides, f"✅ Added {num} bullet points to slide {slide_num}.")
                        return f"Added {num} bullet points to slide {slide_num}.", True
                return result.get("question", "Edit failed. Please rephrase."), False
        except Exception as e:
            return f"Edit failed: {e}", False

    elif intent == "add_slide":
        slide_content = slots.get("slide_content")
        position = slots.get("position", "end")
        if not slide_content:
            return "Missing slide content.", False
        try:
            new_slide_data = draft_slide_from_request(slide_content, slides)
            new_slide_data = normalize_slide(new_slide_data)
            if position is None:
                position = "end"
            if isinstance(position, str):
                pos_lower = position.lower()
                if pos_lower in ("start", "first", "top"):
                    idx = 1
                elif pos_lower in ("end", "last", "bottom"):
                    idx = len(slides) + 1
                else:
                    after_match = re.search(r"after slide (\d+)", pos_lower)
                    before_match = re.search(r"before slide (\d+)", pos_lower)
                    if after_match:
                        idx = int(after_match.group(1)) + 1
                    elif before_match:
                        idx = int(before_match.group(1))
                    else:
                        idx = len(slides) + 1
            else:
                try:
                    idx = min(max(1, int(position)), len(slides) + 1)
                except (ValueError, TypeError):
                    idx = len(slides) + 1
            updated_slides = list(slides)
            updated_slides.insert(idx - 1, new_slide_data)
            updated_slides = refresh_section_index_slide(updated_slides)
            commit_changes(updated_slides, f"✅ New slide added at position {idx}.")
            return f"Slide added at position {idx}.", True
        except Exception as e:
            return f"Add slide failed: {e}", False

    elif intent == "delete_slide":
        slide_num = slots.get("slide_number")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number to delete.", False
        if 1 <= slide_num <= len(slides):
            updated_slides = [s for i, s in enumerate(slides, start=1) if i != slide_num]
            updated_slides = refresh_section_index_slide(updated_slides)
            commit_changes(updated_slides, f"✅ Slide {slide_num} removed.")
            return f"Slide {slide_num} removed.", True
        else:
            return f"Slide {slide_num} does not exist. The deck has {len(slides)} slide(s).", False

    elif intent == "download_ppt":
        ppt_ref = slots.get("ppt_ref")
        if ppt_ref:
            best_id, clarification = find_ppt_semantically(str(ppt_ref), disambiguate=True)
            if clarification:
                return clarification, False
            if best_id:
                switch_active_ppt(best_id)
                return f"Switched to {best_id}. You can download it using the button in the preview above.", True
            else:
                return "I couldn't find that PPT. Try 'ppt 1', 'second ppt', or 'previous ppt'.", False
        else:
            if st.session_state.get("ppt_bytes"):
                return "You can download the current PPT using the button in the preview above.", True
            else:
                return "No PPT has been built yet. Please create a presentation first.", False

    elif intent == "preview_ppt":
        ppt_ref = slots.get("ppt_ref")
        if ppt_ref:
            best_id, clarification = find_ppt_semantically(str(ppt_ref), disambiguate=True)
            if clarification:
                return clarification, False
            if best_id:
                switch_active_ppt(best_id)
                item = get_ppt_by_id(best_id)
                if item and item.get("outline_payload"):
                    preview_slides = item["outline_payload"].get("slides", [])
                    topic = item.get("topic", "Presentation")
                    label = f"Preview — {topic} ({best_id})"
                    add_message_with_preview(
                        role="assistant",
                        text=f"Showing preview for **{best_id}** — {topic}.",
                        preview_slides=preview_slides,
                        ppt_label=label,
                        ppt_bytes=item.get("ppt_bytes"),
                        ppt_filename=item.get("ppt_filename", "presentation.pptx"),
                    )
                    st.rerun()
                return f"Showing preview for **{best_id}**.", True
            else:
                return "I couldn't find that PPT.", False
        else:
            outline = st.session_state.get("outline_payload")
            if outline and outline.get("slides"):
                preview_slides = outline["slides"]
                ppt_id = st.session_state.get("current_ppt_id", "")
                topic = st.session_state.get("topic", "Presentation")
                label = f"Preview — {topic} ({ppt_id})"
                add_message_with_preview(
                    role="assistant",
                    text="Here is the current presentation:",
                    preview_slides=preview_slides,
                    ppt_label=label,
                    ppt_bytes=st.session_state.get("ppt_bytes"),
                    ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
                )
                st.rerun()
            return "The current presentation is shown above.", True

    elif intent == "ppt_info":
        info_type = slots.get("info_type")
        if info_type == "count":
            total = len(st.session_state.get("ppt_history", []))
            return f"You have created {total} presentation(s) so far.", True
        elif info_type == "list":
            history = st.session_state.get("ppt_history", [])
            if not history:
                return "You haven't created any presentations yet.", True
            lines = ["Here are all your presentations:"]
            for i, p in enumerate(history, 1):
                topic = p.get("topic", f"PPT {i}")
                slide_count = p.get("slide_count", "?")
                ppt_id = p.get("id", f"ppt_{i}")
                lines.append(f"**{i}.** {ppt_id}: {slide_count} slides — {topic}")
            return "\n".join(lines), True
        else:
            return "I can tell you how many PPTs you have or list them. What would you like?", False

    elif intent == "suggest_topic":
        scope = slots.get("scope")
        if scope == "new":
            suggestions = ["Responsible AI adoption in enterprises", "AI copilots and productivity", "Cybersecurity risks in the age of AI"]
        else:
            suggestions = ["Next-gen use cases and ROI", "Governance, safety, and compliance", "Roadmap and implementation plan"]
        return "Here are a few topic ideas:\n\n" + "\n".join(f"- {s}" for s in suggestions), True

    elif intent == "greeting":
        return "👋 Hello! I'm your PPT assistant. You can ask me to create a new presentation, edit slides, add content, or switch between decks. What would you like to do?", True
    elif intent == "smalltalk":
        return "Got it! I'm here when you're ready to work on your presentation. Just tell me what you'd like to do.", True
    else:
        return "I'm not sure how to help with that. Could you rephrase?", False

# ------------------------------------------------------------------------------
#  LLM HELPERS (edit, draft)
# ------------------------------------------------------------------------------
def summarize_slides_for_llm(slides: list) -> list:
    summary = []
    for i, slide in enumerate(slides or [], start=1):
        if not isinstance(slide, dict):
            continue
        content = slide.get("content", []) or []
        if not isinstance(content, list):
            content = []
        summary.append({
            "slide_number": i,
            "title": str(slide.get("title", "")).strip(),
            "subtitle": str(slide.get("subtitle", "")).strip(),
            "layout": str(slide.get("layout", "bullets")).strip(),
            "content_preview": [str(x).strip() for x in content[:3] if str(x).strip()],
        })
    return summary

def extract_first_json(text: str) -> dict:
    text = text.strip()
    start = text.find('{')
    if start == -1:
        raise ValueError("No JSON object found")
    brace_count = 0
    in_string = False
    escape = False
    for i, ch in enumerate(text[start:], start=start):
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '"' and not escape:
            in_string = not in_string
        if not in_string:
            if ch == '{':
                brace_count += 1
            elif ch == '}':
                brace_count -= 1
                if brace_count == 0:
                    return json.loads(text[start:i + 1])
    raise ValueError("Unbalanced braces")

def interpret_edit_with_llm(user_text, slides):
    prompt = f"""
You are an AI presentation editor. You MUST preserve existing content and only modify as requested.

User request: {user_text}

Current slides (FULL DECK):
{json.dumps(slides, indent=2)}

RULES:
1. Identify the slide number from the request (e.g., "slide 9").
2. For that slide, **KEEP ALL EXISTING CONTENT**.
3. If the request says "add X bullet points", generate X new, relevant bullet points based on the slide's title and existing content, and **APPEND** them to the existing content list.
4. Do NOT remove or replace any existing bullet points.
5. If the request says "change" or "replace", then you may modify, but for "add" you must only append.
6. Return the ENTIRE updated slides array with all other slides unchanged.
7. Content fields must be plain string arrays – no nested dicts.

Return ONLY valid JSON:
{{
  "action": "edit",
  "slides": [ ... full updated slides array ... ]
}}
"""
    response = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    usage = usage_to_dict(response.usage)
    raw = response.choices[0].message.content
    try:
        data = extract_first_json(raw)
        # Clean content in the edited slides
        if "slides" in data:
            for s in data["slides"]:
                if "content" in s and isinstance(s["content"], list):
                    s["content"] = [clean_icon_tokens(item) for item in s["content"] if clean_icon_tokens(item)]
                if "left_points" in s:
                    s["left_points"] = [clean_icon_tokens(p) for p in s["left_points"] if clean_icon_tokens(p)]
                if "right_points" in s:
                    s["right_points"] = [clean_icon_tokens(p) for p in s["right_points"] if clean_icon_tokens(p)]
        return data, usage
    except Exception as e:
        st.error(f"JSON parse error: {e}\nRaw output: {raw[:500]}")
        raise ValueError("Invalid JSON from LLM")

def draft_slide_from_request(user_text: str, slides: list):
    titles = [str(s.get("title", "")).strip() for s in (slides or []) if isinstance(s, dict)]
    prompt = f"""
Create a single new slide for a presentation.
User request: {user_text}
Existing slide titles: {titles}

Return ONLY valid JSON for a single slide object with fields:
title, subtitle, layout, icon, content (list of plain strings only — no nested dicts), style (dict).
Choose the most appropriate layout.
"""
    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4,
    )
    raw = (resp.choices[0].message.content or "").strip()
    if not raw.startswith("{"):
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        raw = m.group() if m else raw
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Invalid slide JSON")
    # Clean the content immediately
    if "content" in data and isinstance(data["content"], list):
        data["content"] = [clean_icon_tokens(item) for item in data["content"] if clean_icon_tokens(item)]
    if "left_points" in data:
        data["left_points"] = [clean_icon_tokens(p) for p in data["left_points"] if clean_icon_tokens(p)]
    if "right_points" in data:
        data["right_points"] = [clean_icon_tokens(p) for p in data["right_points"] if clean_icon_tokens(p)]
    data.setdefault("title", "New Slide")
    data.setdefault("subtitle", "")
    data.setdefault("layout", "bullets")
    data.setdefault("icon", "▸")
    data.setdefault("content", ["Add your main point here."])
    data.setdefault("style", {})
    return data

# ------------------------------------------------------------------------------
#  OUTLINE GENERATION
# ------------------------------------------------------------------------------
def has_asked(step: str) -> bool:
    return step in st.session_state.get("asked_steps", [])
def mark_asked(step: str):
    asked = st.session_state.get("asked_steps", [])
    if step not in asked:
        asked.append(step)
        st.session_state["asked_steps"] = asked
def build_outline_topic(topic: str, sections) -> str:
    base = (topic or "").strip() or "Presentation"
    if sections:
        return f"{base}. Include sections: {sections}."
    return base

def request_outline(topic: str, num_slides: int, tone: str, sections=None, theme_colors=None) -> dict:
    topic_for_outline = build_outline_topic(topic, sections)
    data = {"topic": topic_for_outline, "num_slides": num_slides, "tone": tone}
    if theme_colors:
        data["theme_colors"] = json.dumps(theme_colors)
    resp = requests.post(OUTLINE_URL, data=data, timeout=(10, 180))
    if resp.status_code != 200:
        raise RuntimeError(resp.text)
    payload = resp.json()
    usage = payload.get("usage")
    if isinstance(usage, dict):
        record_usage("outline", usage)
    slides = [ensure_editor_id(s) for s in payload.get("slides", [])]
    slides = enforce_slide_count(slides, num_slides, topic, tone, theme_colors)
    payload["slides"] = slides
    return payload

def summarize_outline(slides: list) -> str:
    if not slides:
        return "I couldn't generate an outline yet. Try another prompt."
    lines = ["✅ I drafted these sections for the deck:\n"]
    for i, s in enumerate(slides, 1):
        lines.append(f"**{i}.** {s.get('title', f'Slide {i}')}")
    lines.append("\nAsk me in chat if you want any changes.")
    return "\n".join(lines)

def enforce_slide_count(slides, required_count, topic, tone, theme_colors=None):
    slides = slides or []
    if len(slides) > required_count:
        return slides[:required_count]
    max_attempts = 5
    attempts = 0
    while len(slides) < required_count and attempts < max_attempts:
        missing = required_count - len(slides)
        extra_data = generate_slide_content(topic, missing, tone, theme_colors)
        extra_slides = extra_data.get("slides", [])
        if len(extra_slides) > missing:
            extra_slides = extra_slides[-missing:]
        if not extra_slides:
            break
        slides.extend(extra_slides)
        attempts += 1
    return slides[:required_count]

def generate_outline_and_reply(topic: str, count: int, tone: str, sections=None, theme_colors=None):
    with st.spinner(f"✍️ Building {count}-slide deck on '{topic}'…"):
        try:
            payload = request_outline(topic, count, tone, sections, theme_colors)
            payload["slides"] = payload.get("slides", [])[:count]
            st.session_state.outline_payload = payload
            st.session_state.topic = topic
            st.session_state.sections = sections or ""
            st.session_state.num_slides = count
            st.session_state.slide_count = count
            ppt_id = next_ppt_id()
            st.session_state.current_ppt_id = ppt_id
            _upsert_ppt_history(
                ppt_id,
                topic=topic,
                slide_count=count,
                sections=sections or "",
                outline_payload=payload,
                ppt_bytes=None,
                ppt_filename=None,
                file_path=None,
            )
            sync_all_editor_widgets(payload.get("slides", []))
            ppt_bytes = rebuild_ppt_from_outline()
            reply_text = summarize_outline(payload.get("slides", []))
            slides_for_preview = payload.get("slides", [])
            label = f"Preview — {topic} ({ppt_id})"
            add_message_with_preview(
                role="assistant",
                text=reply_text,
                preview_slides=slides_for_preview,
                ppt_label=label,
                ppt_bytes=ppt_bytes,
                ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
            )
            st.rerun()
        except Exception as e:
            st.error(f"Outline generation failed: {e}")

def rebuild_ppt_from_outline() -> bytes | None:
    outline = st.session_state.outline_payload
    if not outline or not outline.get("slides"):
        return None
    build_payload = sanitize_outline_for_build(copy.deepcopy(outline))
    files = None
    if st.session_state.get("logo_bytes"):
        files = {
            "logo": (
                st.session_state.get("logo_name", "logo.png"),
                st.session_state.logo_bytes,
                "image/png",
            )
        }
    try:
        resp = requests.post(
            BUILD_URL,
            data={
                "topic": st.session_state.topic or "Presentation",
                "tone": st.session_state.tone,
                "slides_json": json.dumps(build_payload),
            },
            files=files,
            timeout=(10, 180),
        )
        if resp.status_code == 200:
            result = resp.json()
            ppt_bytes = base64.b64decode(result["ppt_base64"])
            slides = outline.get("slides", [])
            first_title = ""
            if isinstance(slides, list) and slides:
                first = slides[0] if isinstance(slides[0], dict) else {}
                first_title = str(first.get("title", "")).strip()
            file_topic = first_title or st.session_state.topic or "presentation"
            safe_name = re.sub(r"[^\w\-]+", "_", file_topic).strip("_") or "presentation"
            filename = f"{safe_name[:80]}.pptx"
            st.session_state.ppt_bytes = ppt_bytes
            st.session_state.ppt_filename = filename
            ppt_id = st.session_state.get("current_ppt_id")
            if ppt_id:
                _upsert_ppt_history(
                    ppt_id,
                    ppt_bytes=ppt_bytes,
                    ppt_filename=filename,
                    file_path=None,
                    outline_payload=outline,
                    topic=st.session_state.topic,
                    sections=st.session_state.sections,
                    slide_count=len(slides),
                )
            return ppt_bytes
        else:
            st.session_state.last_build_error = resp.text
            return None
    except Exception as e:
        st.session_state.last_build_error = str(e)
        return None

def sanitize_outline_for_build(payload: dict) -> dict:
    safe = {
        "design_system": payload.get("design_system", {}) if isinstance(payload, dict) else {},
        "slides": [],
    }
    for slide in (payload.get("slides", []) if isinstance(payload, dict) else []):
        if not isinstance(slide, dict):
            continue
        layout = str(slide.get("layout", "bullets")).strip() or "bullets"
        s = {
            "title": str(slide.get("title", "")).strip(),
            "subtitle": str(slide.get("subtitle", "")).strip(),
            "layout": layout,
            "icon": clean_bullet_icon(slide.get("icon", "▸")),
            "content": flatten_slide_content(slide.get("content", [])),
            "style": slide.get("style", {}) if isinstance(slide.get("style"), dict) else {},
        }
        if layout == "two_column":
            s["left_title"] = slide.get("left_title", "Left")
            s["right_title"] = slide.get("right_title", "Right")
            lp = flatten_slide_content(slide.get("left_points", []))
            rp = flatten_slide_content(slide.get("right_points", []))
            if not lp and not rp:
                mid = max(1, len(s["content"]) // 2)
                lp, rp = s["content"][:mid], s["content"][mid:]
            s["left_points"] = lp
            s["right_points"] = rp
        elif layout == "big_stat":
            s["stat"] = slide.get("stat", "—")
            s["stat_label"] = slide.get("stat_label", "")
            s["stat_source"] = slide.get("stat_source", "")
        elif layout == "timeline":
            s["steps"] = slide.get("steps", []) or []
        elif layout == "icon_grid":
            s["grid_items"] = normalize_icon_grid_items(slide, bullets_to_text(flatten_slide_content(slide.get("content", []))))
            s["content"] = [f"{item['title']}: {item['detail']}".rstrip(": ").strip() for item in s["grid_items"]]
        elif layout == "case_study":
            s["company"] = slide.get("company", "")
            s["result"] = slide.get("result", "")
            s["metrics"] = slide.get("metrics", []) or []
        elif layout == "table":
            s["table_columns"] = slide.get("table_columns", []) or []
            s["table_rows"] = slide.get("table_rows", []) or []
        elif layout == "chart":
            s["chart_title"] = slide.get("chart_title", "")
            s["chart_data"] = slide.get("chart_data", []) or []
            s["chart_source"] = slide.get("chart_source", "")
        elif layout == "hybrid_insight":
            s["stat"] = slide.get("stat", "—")
            s["stat_label"] = slide.get("stat_label", "")
            s["chart_data"] = slide.get("chart_data", []) or []
        elif layout == "section_index":
            s["sections"] = slide.get("sections", s["content"])
        safe["slides"].append(s)
    return safe

def refresh_section_index_slide(slides: list) -> list:
    normalized = [ensure_editor_id(s) for s in (slides or []) if isinstance(s, dict)]
    if len(normalized) < 2:
        return normalized
    index_slide = normalized[1]
    if str(index_slide.get("layout", "")).strip() != "section_index":
        return normalized
    sections = [str(s.get("title", "")).strip() for s in normalized[2:] if isinstance(s, dict) and str(s.get("title", "")).strip()][:6]
    index_slide["sections"] = sections
    index_slide["content"] = sections
    normalized[1] = index_slide
    return normalized

def make_thank_you_slide(slide: dict) -> dict:
    slide = ensure_editor_id(slide)
    slide["title"] = "Thank You"
    slide["subtitle"] = "Questions?"
    slide["layout"] = "title_cover"
    slide["content"] = []
    return slide

def apply_add_action(old_slides, slide, position):
    new_slide = ensure_editor_id(slide)
    new_slide["_editor_id"] = uuid4().hex
    slides = list(old_slides)
    if isinstance(position, dict):
        if "before" in position:
            target = position["before"].lower()
            for i, s in enumerate(slides):
                title = s.get("title", "").lower()
                layout = s.get("layout", "").lower()
                if target in title or target == layout:
                    slides.insert(i, new_slide)
                    return slides
        elif "after" in position:
            target = position["after"].lower()
            for i, s in enumerate(slides):
                title = s.get("title", "").lower()
                layout = s.get("layout", "").lower()
                if target in title or target == layout:
                    slides.insert(i + 1, new_slide)
                    return slides
        slides.append(new_slide)
        return slides
    if isinstance(position, int):
        idx = max(1, min(len(slides) + 1, position))
        slides.insert(idx - 1, new_slide)
        return slides
    pos = str(position or "").strip().lower()
    if pos in ("end", "last", "bottom"):
        slides.append(new_slide)
    elif pos in ("start", "first", "top"):
        slides.insert(0, new_slide)
    else:
        slides.append(new_slide)
    return slides

# ------------------------------------------------------------------------------
#  MAIN UI — CHAT-FIRST LAYOUT
# ------------------------------------------------------------------------------
st.title("💬 AI PPT Chat Builder")
st.caption("Describe your deck, then edit it in chat and download.")

# Sidebar settings
st.sidebar.title("⚙️ Settings")
tone = st.sidebar.selectbox(
    "Presentation Style",
    ["Professional", "Creative", "Educational"],
    index=["Professional", "Creative", "Educational"].index(st.session_state.tone),
)
st.session_state.tone = tone

st.sidebar.markdown("---")
st.sidebar.subheader("Token Usage")
last_usage = st.session_state.get("last_usage")
if last_usage:
    st.sidebar.write(f"Last: {last_usage.get('kind', '')}")
    st.sidebar.write(f"Input: {last_usage.get('prompt_tokens', 0)}")
    st.sidebar.write(f"Output: {last_usage.get('completion_tokens', 0)}")
    st.sidebar.write(f"Total: {last_usage.get('total_tokens', 0)}")
else:
    st.sidebar.write("No usage yet.")
totals = st.session_state.get("token_usage", {})
st.sidebar.write("**Session total**")
st.sidebar.write(f"Input: {totals.get('prompt_tokens', 0)}")
st.sidebar.write(f"Output: {totals.get('completion_tokens', 0)}")
st.sidebar.write(f"Total: {totals.get('total_tokens', 0)}")

st.sidebar.markdown("---")
st.sidebar.subheader("Session")
st.sidebar.code(st.session_state.get("session_id", ""))
msgs = st.session_state.get("messages", [])
st.sidebar.write(f"Turns: {len(msgs)}")

# FIX 1: Sidebar-only PPT history list (no duplicate bottom section)
history = st.session_state.get("ppt_history", [])
if history:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Presentations")
    for idx, item in enumerate(history, start=1):
        topic_label = item.get("topic") or f"PPT {idx}"
        is_active = item.get("id") == st.session_state.get("current_ppt_id")
        badge = " 🟢" if is_active else ""
        st.sidebar.caption(f"**{idx}. {topic_label}**{badge}")
        ppt_bytes = item.get("ppt_bytes")
        ppt_name = item.get("ppt_filename") or f"{item.get('id')}.pptx"
        col1, col2 = st.sidebar.columns([2, 1])
        with col1:
            st.download_button(
                label="⬇️",
                data=ppt_bytes or b"",
                file_name=ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                disabled=not bool(ppt_bytes),
                key=f"sdl_{item.get('id')}",
                use_container_width=True,
            )
        with col2:
            if not is_active:
                if st.sidebar.button("↩", key=f"ssw_{item.get('id')}"):
                    switch_active_ppt(item.get("id"))
                    st.rerun()

# ------------------------------------------------------------------------------
#  CHAT HISTORY — replay with embedded previews
# ------------------------------------------------------------------------------
for msg_index, msg in enumerate(st.session_state.messages):
    render_message(msg, msg_index=msg_index)

# Chat input
chat_value = None
try:
    chat_value = st.chat_input(
        "Describe a topic or ask to edit a slide…",
        key="chat_prompt",
        accept_file=True,
        file_type=["png", "jpg", "jpeg"],
    )
except TypeError:
    chat_value = st.chat_input(
        "Describe a topic or ask to edit a slide…",
        key="chat_prompt",
    )

prompt, files = extract_chat_input(chat_value)

# Handle logo upload
if files:
    uploaded_logo = files[0]
    if uploaded_logo:
        logo_bytes = uploaded_logo.read()
        if logo_bytes != st.session_state.get("logo_bytes"):
            st.session_state.logo_bytes = logo_bytes
            st.session_state.logo_name = uploaded_logo.name
            if st.session_state.get("outline_payload"):
                rebuild_ppt_from_outline()

if prompt is not None:
    prompt = prompt.strip()
    if not prompt:
        st.stop()

    # If this is a create_ppt request and no file was uploaded, clear any leftover logo
    if (files is None or len(files) == 0) and re.search(r'\b(?:make|create|generate|new)\s+(?:a\s+)?(?:ppt|presentation|deck)', prompt, re.IGNORECASE):
        st.session_state.logo_bytes = None
        st.session_state.logo_name = None

    # Render user message immediately
    with st.chat_message("user"):
        st.markdown(prompt)
    add_message("user", prompt)

    # 1. Resolve PPT reference (possibly switch active PPT)
    def should_resolve_ppt(user_input: str) -> bool:
        return bool(re.search(r"(ppt|presentation|deck)\s*\d+|switch|open|go to|previous ppt", user_input.lower()))
    if should_resolve_ppt(prompt):
        changed, new_ppt_id, clarification = resolve_ppt_reference(prompt)
    else:
        changed, new_ppt_id, clarification = False, None, None
    if clarification:
        with st.chat_message("assistant"):
            st.markdown(clarification)
        add_message("assistant", clarification)
        st.rerun()
    if changed:
        with st.chat_message("assistant"):
            st.markdown(f"✅ Switched to presentation **{new_ppt_id}**. How can I help?")
        add_message("assistant", f"Switched to {new_ppt_id}.")

    # Prepare context for the agent
    outline_payload = st.session_state.outline_payload
    slides = outline_payload.get("slides", []) if outline_payload else []
    context = {
        "has_deck": bool(outline_payload),
        "slides": slides,
    }

    # Initialize agent
    if "agent" not in st.session_state:
        st.session_state.agent = ConversationalAgent()
    agent = st.session_state.agent

    # Process with agent (intent detection + slot filling)
    response_text, action_data, should_execute = agent.process(prompt, context)

    if should_execute and action_data:
        intent = action_data["intent"]
        slots = action_data["slots"]
        result_msg, success = execute_action(intent, slots, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        agent.clear_state()
        st.rerun()
    elif response_text:
        with st.chat_message("assistant"):
            st.markdown(response_text)
        add_message("assistant", response_text)
        st.rerun()
    else:
        with st.chat_message("assistant"):
            st.markdown("I'm not sure how to help. Could you rephrase?")
        add_message("assistant", "I'm not sure how to help. Could you rephrase?")
        st.rerun()
