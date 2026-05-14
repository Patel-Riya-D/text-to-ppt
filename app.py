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
from backend.services.intelligence_layer import (
    _should_execute_directly,
    _confidence_preface,
    _natural_success_message,
)
from backend.services.intent_classifier import classify_ppt_intent
from backend.services.assistant_orchestrator import (
    create_topic_followup as orchestrator_create_topic_followup,
    extract_conversational_create_topic as orchestrator_extract_conversational_create_topic,
    has_explicit_deterministic_command as orchestrator_has_explicit_deterministic_command,
    is_contextual_followup_candidate as orchestrator_is_contextual_followup_candidate,
    is_generic_presentation_context as orchestrator_is_generic_presentation_context,
    is_implicit_create_request as orchestrator_is_implicit_create_request,
    is_readonly_conversation_request as orchestrator_is_readonly_conversation_request,
    pre_route_message,
    resolve_contextual_intent as orchestrator_resolve_contextual_intent,
    resolve_open_conversation as orchestrator_resolve_open_conversation,
)
from backend.services.ppt_service import generate_slide_content, merge_slides as merge_slide_payloads

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

OUTLINE_URL = "http://127.0.0.1:9000/generate-outline"
BUILD_URL   = "http://127.0.0.1:9000/build-ppt"

st.set_page_config(page_title="AI PPT Generator", page_icon="💬", layout="wide")

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
_MAX_POINTS_PER_REQUEST = 5
_MAX_SLIDES_WITHOUT_CONFIRMATION = 15


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


def _norm_text(v):
    """Normalize text for fuzzy PPT/topic matching."""
    s = str(v or "").casefold()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s

def _safe_str(v, default=""):
    return str(v).strip() if v is not None else default

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

_COLOR_NAMES_ORDERED = sorted(_COLOR_NAMES, key=len, reverse=True)
_COLOR_NAME_PATTERN = re.compile(
    r"\b(" + "|".join(re.escape(name) for name in _COLOR_NAMES_ORDERED) + r")\b",
    re.IGNORECASE,
)
_NEUTRAL_THEME_COLORS = {"white", "black", "gray", "grey", "silver"}

def _prioritize_theme_colors(colors: list[str]) -> list[str]:
    """
    Prefer a vivid color as the primary theme color.

    This keeps requests like "white and green" from turning the deck into
    a gray-first palette, which would make headers and accents feel dull.
    """
    if len(colors) < 2:
        return colors
    primary = str(colors[0]).strip().lower()
    secondary = str(colors[1]).strip().lower()
    if primary in _NEUTRAL_THEME_COLORS and secondary not in _NEUTRAL_THEME_COLORS:
        reordered = [colors[1], colors[0]]
        if len(colors) > 2:
            reordered.extend(colors[2:])
        return reordered
    return colors


def extract_theme_colors(text: str) -> dict:
    """Extract primary/secondary color names from natural language."""
    if not text:
        return {}

    lower = text.lower()
    colors = []   # ✅ FIX: initialize

    # 🎯 Detect "blue and white" type patterns
    paired_patterns = [
        r"(?P<first>" + _COLOR_NAME_PATTERN.pattern[2:-2] + r")\s+and\s+(?P<second>" + _COLOR_NAME_PATTERN.pattern[2:-2] + r")",
    ]

    for pat in paired_patterns:
        m = re.search(pat, lower, re.IGNORECASE)
        if m:
            first = m.group("first").strip().lower()
            second = m.group("second").strip().lower()
            for c in (first, second):
                if c in _COLOR_NAMES and c not in colors:
                    colors.append(c)
            break

    # 🎯 fallback: detect any color words
    if not colors:
        for m in _COLOR_NAME_PATTERN.finditer(lower):
            c = m.group(1).strip().lower()
            if c not in colors:
                colors.append(c)

    if not colors:
        return {}

    colors = _prioritize_theme_colors(colors)

    result = {"primary": colors[0]}
    if len(colors) > 1:
        result["secondary"] = colors[1]

    return result   # ✅ IMPORTANT

def extract_theme_colors_from_messages(messages: list) -> dict:
    """Find the first user message that explicitly mentions theme colors."""
    for msg in reversed(messages or []):
        if not isinstance(msg, dict) or msg.get("role") != "user":
            continue
        colors = extract_theme_colors(msg.get("content", ""))
        if colors:
            return colors
    return {}

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
        "ppt_store":       {},
        "current_ppt_id":  None,
        "current_step":    "topic",
        "asked_steps":     [],
        "token_usage":     {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
        "last_usage":      None,
        "session_id":      uuid4().hex,
        "last_question":   None,
        "pending_intent":  None,
        "pending_action":  None,
        "pending_new_ppt": None,
        "active_slide_index": None,
        "last_slide_index": None,
        "last_action_type": None,
        "last_ppt_id": None,
        "previous_ppt_id": None,
        "recent_ppt_ids": [],
        "last_build_error": None,
        "active_topic_domain": "",
        "pending_topic_switch": None,
        "conversation_objects": {},
        "session_memory":  {"ppts": [], "current_ppt": None},
        "active_edit_context": None,
        "conversation_state": {
            "active_ppt_id": None,
            "active_slide_index": None,
            "last_topic": "",
            "last_action_type": None,
            "last_action_text": "",
            "pending_create": None,
            "pending_edit": None,
            "turn_summary": "",
        },
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ------------------------------------------------------------------------------
#  HELPER FUNCTIONS (slide utilities, history, etc.)
# ------------------------------------------------------------------------------
def _conv_state() -> dict:
    state = st.session_state.get("conversation_state")
    if not isinstance(state, dict):
        state = {}
        st.session_state["conversation_state"] = state
    state.setdefault("active_ppt_id", st.session_state.get("current_ppt_id"))
    state.setdefault("active_slide_index", st.session_state.get("active_slide_index"))
    state.setdefault("last_topic", st.session_state.get("topic", ""))
    state.setdefault("last_action_type", st.session_state.get("last_action_type"))
    state.setdefault("last_action_text", "")
    state.setdefault("pending_create", st.session_state.get("pending_new_ppt"))
    state.setdefault("pending_edit", st.session_state.get("pending_intent"))
    state.setdefault("pending_action", st.session_state.get("pending_action"))
    state.setdefault("turn_summary", "")
    return state

def _set_conv_state(**kwargs):
    state = _conv_state()
    state.update(kwargs)
    st.session_state["conversation_state"] = state
    if "active_ppt_id" in kwargs:
        st.session_state["last_ppt_id"] = kwargs["active_ppt_id"]
    if "active_slide_index" in kwargs:
        st.session_state["active_slide_index"] = kwargs["active_slide_index"]
        st.session_state["last_slide_index"] = kwargs["active_slide_index"]
    if "last_topic" in kwargs:
        st.session_state["topic"] = kwargs["last_topic"]
    if "last_action_type" in kwargs:
        st.session_state["last_action_type"] = kwargs["last_action_type"]
    if "pending_create" in kwargs:
        st.session_state["pending_new_ppt"] = kwargs["pending_create"]
    if "pending_edit" in kwargs:
        st.session_state["pending_intent"] = kwargs["pending_edit"]
    if "pending_action" in kwargs:
        st.session_state["pending_action"] = kwargs["pending_action"]

def _clear_transient_conversation_state():
    """
    Clear turn-specific state after an action completes.

    We keep the active PPT itself, but remove stale intent/slide context so the
    next user message is evaluated as a fresh request unless it explicitly
    continues the previous turn.
    """
    current_ppt_id = st.session_state.get("current_ppt_id")
    current_topic = st.session_state.get("topic", "")
    state = _conv_state()
    state.update({
        "active_ppt_id": current_ppt_id,
        "active_slide_index": None,
        "last_topic": current_topic,
        "last_action_type": None,
        "last_action_text": "",
        "pending_create": None,
        "pending_edit": None,
        "turn_summary": "",
    })
    st.session_state["conversation_state"] = state
    st.session_state["pending_intent"] = None
    st.session_state["pending_action"] = None
    st.session_state["pending_new_ppt"] = None
    st.session_state["active_edit_context"] = None
    st.session_state["last_action_type"] = None
    st.session_state["active_slide_index"] = None
    st.session_state["last_slide_index"] = None
    st.session_state["last_ppt_id"] = current_ppt_id

def _build_context_badge() -> str:
    state = _conv_state()
    parts = []
    if state.get("active_ppt_id"):
        parts.append(_ppt_display_label(state["active_ppt_id"]))
    if state.get("active_slide_index"):
        parts.append(f"slide {state['active_slide_index']}")
    if state.get("last_action_type"):
        parts.append(state["last_action_type"].replace("_", " "))
    return " | ".join(parts) if parts else "No active deck"

def add_message(role, content):
    st.session_state.messages.append({"role": role, "content": content})

def _conversation_objects() -> dict:
    objects = st.session_state.get("conversation_objects")
    if not isinstance(objects, dict):
        objects = {}
        st.session_state["conversation_objects"] = objects
    return objects

def _remember_conversation_object(key: str, value: dict) -> None:
    if not key or not isinstance(value, dict):
        return
    objects = _conversation_objects()
    objects[key] = copy.deepcopy(value)
    st.session_state["conversation_objects"] = objects

def _extract_numbered_suggestions(text: str) -> list[dict]:
    suggestions = []
    for line in str(text or "").splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        match = re.match(r"^(?:[-*]\s*)?(?:(\d+)[\.\)]\s+)?(.+)$", cleaned)
        if not match:
            continue
        body = re.sub(r"^\*\*(.+?)\*\*:?\s*", r"\1: ", match.group(2).strip())
        if len(body.split()) < 3:
            continue
        if not match.group(1) and not re.match(r"^[-*]\s+", cleaned):
            continue
        suggestions.append({
            "index": int(match.group(1) or len(suggestions) + 1),
            "type": "suggestion",
            "content": body,
        })
    return suggestions[:8]

def _remember_answer_objects(user_input: str, answer: str, target_slide: Optional[int] = None) -> None:
    answer = _safe_str(answer, "")
    if not answer:
        return
    suggestions = _extract_numbered_suggestions(answer)
    if suggestions:
        for item in suggestions:
            item["slide_number"] = target_slide
            item["source_user_input"] = user_input
        _remember_conversation_object("last_suggestions", {"type": "suggestion_list", "items": suggestions})
        _remember_conversation_object("last_suggestion", suggestions[-1])
    _remember_conversation_object(
        "last_generated_content",
        {
            "type": "assistant_answer",
            "content": answer,
            "slide_number": target_slide,
            "source_user_input": user_input,
        },
    )

def _ordinal_to_int(value: str) -> Optional[int]:
    text = _safe_str(value, "").lower()
    words = {
        "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
        "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
    }
    if text in words:
        return words[text]
    match = re.match(r"(\d+)(?:st|nd|rd|th)?$", text)
    if match:
        return int(match.group(1))
    return None

def _selected_suggestion_from_text(user_input: str) -> Optional[dict]:
    text = _safe_str(user_input, "")
    match = re.search(
        r"\b((?:\d+)(?:st|nd|rd|th)?|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)\s+suggestion\b",
        text,
        re.IGNORECASE,
    )
    idx = _ordinal_to_int(match.group(1)) if match else None
    objects = _conversation_objects()
    suggestions = (objects.get("last_suggestions") or {}).get("items") or []
    if idx and 1 <= idx <= len(suggestions):
        selected = copy.deepcopy(suggestions[idx - 1])
        _remember_conversation_object("selected_recommendation", selected)
        return selected
    if re.search(r"\b(?:that|this|the|last)\s+suggestion\b", text, re.IGNORECASE):
        suggestion = objects.get("last_suggestion")
        return copy.deepcopy(suggestion) if isinstance(suggestion, dict) else None
    return None

def _resolve_referenced_content(user_input: str) -> Optional[dict]:
    selected = _selected_suggestion_from_text(user_input)
    if selected:
        return selected
    if not re.search(r"\b(?:this|that|it|above|previous|last|suggestion|recommendation)\b", user_input, re.IGNORECASE):
        return None
    objects = _conversation_objects()
    keys = (
        ("selected_recommendation", "last_suggestion", "last_generated_content")
        if re.search(r"\b(?:suggestion|recommendation)\b", user_input, re.IGNORECASE)
        else ("selected_recommendation", "last_generated_content", "last_suggestion")
    )
    for key in keys:
        obj = objects.get(key)
        if isinstance(obj, dict) and _safe_str(obj.get("content"), ""):
            return copy.deepcopy(obj)
    return None

def _extract_reference_target_slide(user_input: str, referenced: Optional[dict] = None) -> Optional[int]:
    slide_num = _extract_explicit_slide_number_from_text(user_input)
    if slide_num is not None:
        return slide_num
    if re.search(r"\b(?:title|cover|first)\s+slide\b|\bslide\s+(?:title|cover)\b", user_input, re.IGNORECASE):
        return 1
    if referenced and referenced.get("slide_number"):
        try:
            return int(referenced.get("slide_number"))
        except (TypeError, ValueError):
            return None
    return None

def _extract_pending_slide_answer(user_input: str, slides: list) -> Optional[int | str]:
    text = _safe_str(user_input, "")
    if re.fullmatch(r"\s*(?:all|all slides?|every slide|each slide)\s*", text, re.IGNORECASE):
        return "all"
    slide_num = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
    if slide_num is not None:
        return slide_num
    ordinal = _ordinal_to_int(text.strip().lower())
    if ordinal and 1 <= ordinal <= len(slides or []):
        return ordinal
    return None

def _is_logo_insertion_request(user_input: str) -> bool:
    text = _safe_str(user_input, "")
    return bool(
        re.search(r"\blogo\b", text, re.IGNORECASE)
        and re.search(r"\b(?:insert|add|place|put|apply|use|include)\b", text, re.IGNORECASE)
    )

def _is_all_slides_target(user_input: str) -> bool:
    return bool(re.search(r"\b(?:all slides?|every slide|each slide|across all|throughout|whole deck|entire deck|all)\b", user_input, re.IGNORECASE))

def _handle_logo_insertion_request(user_input: str, slides: list) -> bool:
    if not _is_logo_insertion_request(user_input):
        return False
    if not st.session_state.get("logo_bytes"):
        reply = "Please upload the logo image first, then tell me where to place it."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    if not st.session_state.get("current_ppt_id") or not slides:
        reply = "No presentation is currently active. Please create or switch to a presentation first."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    if _is_all_slides_target(user_input):
        ppt_bytes = rebuild_ppt_from_outline()
        if not ppt_bytes:
            reply = f"I couldn't rebuild the PPT with the logo. {st.session_state.get('last_build_error') or ''}".strip()
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        topic = st.session_state.get("topic", "Presentation")
        label = f"Updated Preview — {topic} ({st.session_state.get('current_ppt_id')})"
        add_message_with_preview(
            role="assistant",
            text="Inserted the uploaded logo across all slides.",
            preview_slides=slides,
            ppt_label=label,
            ppt_bytes=ppt_bytes,
            ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
        )
        st.rerun()
    slide_num = _extract_explicit_slide_number_from_text(user_input) or _resolve_slide_reference_text(user_input, slides)
    if slide_num is None:
        reply = "Should I place the logo on slide 1 or across all slides?"
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    reply = "Logo placement is currently applied when the PPT is rebuilt. Say `insert this logo into all slides` to add it deck-wide."
    with st.chat_message("assistant"):
        st.markdown(reply)
    add_message("assistant", reply)
    st.rerun()

def _is_reference_execution_request(user_input: str) -> bool:
    return bool(
        re.search(
            r"\b(?:place|put|insert|add|apply|use|move|copy)\b.*\b(?:this|that|it|suggestion|recommendation|\d+(?:st|nd|rd|th)?\s+suggestion|first\s+suggestion|second\s+suggestion|third\s+suggestion|fourth\s+suggestion|fifth\s+suggestion)\b"
            r"|\b(?:yes|yeah|yep|ok|okay|sure)\b.*\bapply\b.*\bsuggestion\b",
            user_input,
            re.IGNORECASE,
        )
    )

def _handle_reference_execution_request(user_input: str, slides: list) -> bool:
    if not _is_reference_execution_request(user_input):
        return False
    referenced = _resolve_referenced_content(user_input)
    if not referenced:
        reply = "I'm not sure what content you mean by \"this.\" Please paste the text or tell me which suggestion to use."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    content = _safe_str(referenced.get("content"), "")
    if not slides:
        reply = "No presentation is currently active. Please create or switch to a presentation first."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    slide_num = _extract_reference_target_slide(user_input, referenced)
    if slide_num is None:
        question = "Which slide should I apply that suggestion to?"
        st.session_state.pending_intent = {
            "intent": "edit_slide",
            "slots": {"slide_number": None, "change_content": content},
            "missing_slots": ["slide_number"],
            "next_question": question,
            "action": "ask",
        }
        st.session_state.pending_action = st.session_state.pending_intent
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()
    if not (1 <= int(slide_num) <= len(slides or [])):
        reply = f"Slide {slide_num} does not exist. The deck has {len(slides or [])} slide(s)."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()
    instruction = content
    if referenced.get("type") in {"suggestion", "assistant_answer"}:
        instruction = f"Apply this recommendation: {content}"
    result_msg, success = execute_action("edit_slide", {"slide_number": int(slide_num), "change_content": instruction}, slides)
    with st.chat_message("assistant"):
        st.markdown(result_msg)
    add_message("assistant", result_msg)
    if success:
        st.rerun()
    return True

def add_message_with_preview(role: str, text: str, preview_slides: list = None,
                              ppt_label: str = "", ppt_bytes: bytes = None,
                              ppt_filename: str = ""):
    entry = {
        "role": role,
        "content": text,
    }
    if preview_slides:
        entry["preview_slides"] = copy.deepcopy(preview_slides)
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

def _is_placeholder_slide_title(title: str, idx: Optional[int] = None) -> bool:
    cleaned = _safe_str(title, "").strip()
    norm = _norm_text(cleaned)
    placeholders = {"", "slide", "untitled", "untitled slide", "title", "new slide"}
    if norm in placeholders:
        return True
    if re.fullmatch(r"slide\s*\d+", norm):
        if idx is None:
            return True
        return norm == f"slide {idx}"
    return False

def _derive_meaningful_slide_title(slide: dict, idx: int = 1) -> str:
    if not isinstance(slide, dict):
        return "Key Takeaways"
    layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()
    if layout == "section_index":
        return "Contents"
    if layout == "title_cover":
        topic = _safe_str(st.session_state.get("topic", ""), "")
        return topic or "Executive Overview"
    if layout == "table":
        cols = flatten_slide_content(slide.get("table_columns", []))
        if len(cols) >= 3:
            return f"{cols[1]} vs {cols[2]}"
        return "Comparison Overview"
    if layout == "two_column":
        left = _safe_str(slide.get("left_title", ""), "")
        right = _safe_str(slide.get("right_title", ""), "")
        if left and right and _norm_text(left) not in {"left", "key points"}:
            return f"{left} vs {right}" if _norm_text(right) not in {"right", "supporting points"} else left
    if layout == "timeline":
        steps = [s for s in (slide.get("steps", []) or []) if isinstance(s, dict)]
        if steps:
            return _safe_str(steps[0].get("label", ""), "") or "Roadmap"
        return "Roadmap"
    if layout == "icon_grid":
        items = [g for g in (slide.get("grid_items", []) or []) if isinstance(g, dict)]
        if items:
            title = _safe_str(items[0].get("title", ""), "")
            if title:
                return title[:60]
        return "Key Pillars"
    candidates = []
    candidates.extend(flatten_slide_content(slide.get("content", [])))
    candidates.extend(flatten_slide_content(slide.get("left_points", [])))
    candidates.extend(flatten_slide_content(slide.get("right_points", [])))
    for raw in candidates:
        text = clean_icon_tokens(_safe_str(raw, ""))
        text = re.split(r"[:.;|–-]", text, maxsplit=1)[0].strip()
        words = [w for w in text.split() if w]
        if 2 <= len(words) <= 7:
            return " ".join(words)[:64]
        if len(words) > 7:
            return " ".join(words[:6])[:64]
    return "Key Takeaways"

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
    if _is_placeholder_slide_title(slide.get("title", ""), None):
        slide["title"] = _derive_meaningful_slide_title(slide, 1)
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
    structured_layouts = {
        "two_column",
        "timeline",
        "icon_grid",
        "case_study",
        "chart",
        "big_stat",
        "hybrid_insight",
        "section_index",
        "table",
    }
    def _preview_column_title(title_value: str, points: list, fallback: str) -> str:
        title_text = _safe_str(title_value, "").strip()
        if title_text and _norm_text(title_text) not in {"left", "right", "left title", "right title", "option a", "option b", "title"}:
            return title_text
        first_point = _safe_str((points or [fallback])[0], "")
        first_point = re.sub(r"^\s*[-•▸►▶➤]+\s*", "", first_point).strip()
        if first_point:
            words = first_point.split()
            derived = " ".join(words[:4]).strip()
            if derived:
                return derived[:40]
        return fallback
    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        title = slide.get("title", f"Slide {idx}")
        layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()
        if layout in {"comparison", "two_col"}:
            layout = "two_column"
        subtitle = slide.get("subtitle", "")
        with st.expander(f"Slide {idx}: {title}  `{layout}`", expanded=(idx == 1)):
            if subtitle:
                st.markdown(f"*{subtitle}*")
            content = flatten_slide_content(slide.get("content", []))
            if layout not in structured_layouts and content:
                for item in content:
                    st.markdown(f"- {item}")
            if layout == "two_column":
                col1, col2 = st.columns(2)
                with col1:
                    left_points = flatten_slide_content(slide.get("left_points", []))
                    st.caption(f"**{_preview_column_title(slide.get('left_title', ''), left_points, 'Key Points')}**")
                    for pt in left_points:
                        st.markdown(f"- {pt}")
                with col2:
                    right_points = flatten_slide_content(slide.get("right_points", []))
                    st.caption(f"**{_preview_column_title(slide.get('right_title', ''), right_points, 'Supporting Points')}**")
                    for pt in right_points:
                        st.markdown(f"- {pt}")
            elif layout == "big_stat":
                st.metric(
                    label=clean_icon_tokens(slide.get("stat_label", "")),
                    value=clean_icon_tokens(slide.get("stat", "—"))
                )
            elif layout == "timeline":
                steps = slide.get("steps", []) or []
                for s in steps:
                    if isinstance(s, dict):
                        label_clean = clean_icon_tokens(s.get("label", ""))
                        detail_clean = clean_icon_tokens(s.get("detail", ""))
                        if label_clean or detail_clean:
                            st.markdown(f"**{label_clean}** — {detail_clean}")
            elif layout == "icon_grid":
                items = slide.get("grid_items", []) or []
                for g in items:
                    if isinstance(g, dict):
                        title_clean = clean_icon_tokens(g.get("title", ""))
                        detail_clean = clean_icon_tokens(g.get("detail", ""))
                        if title_clean or detail_clean:
                            st.markdown(f"{g.get('icon', '•')} **{title_clean}** — {detail_clean}")
            elif layout == "chart":
                data = slide.get("chart_data", []) or []
                chart_title = clean_icon_tokens(slide.get("chart_title", slide.get("title", "")))
                for d in data:
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
                    for item in extra:
                        st.markdown(f"- {item}")
            elif layout == "case_study":
                company = clean_icon_tokens(slide.get("company", ""))
                result = clean_icon_tokens(slide.get("result", ""))
                if company:
                    st.markdown(f"🏢 **{company}**")
                if result:
                    st.markdown(f"✅ {result}")
                extra = flatten_slide_content(slide.get("content", []))
                if extra:
                    st.markdown("**Key Points**")
                    for item in extra:
                        st.markdown(f"- {item}")
            elif layout == "table":
                cols = flatten_slide_content(slide.get("table_columns", []))
                rows = [
                    [_safe_str(cell, "") for cell in row]
                    for row in (slide.get("table_rows", []) or [])
                    if isinstance(row, list)
                ]
                if cols and rows:
                    st.table([dict(zip(cols, row + [""] * max(0, len(cols) - len(row)))) for row in rows])
                elif cols:
                    st.caption("No table rows yet.")
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

def _ppt_store() -> dict:
    store = st.session_state.setdefault("ppt_store", {})
    if not isinstance(store, dict):
        store = {}
        st.session_state["ppt_store"] = store
    return store

def _deepcopy_value(value):
    return copy.deepcopy(value) if value is not None else None

def _upsert_ppt_history(ppt_id: str, **fields):
    history = st.session_state.setdefault("ppt_history", [])
    for item in history:
        if item.get("id") == ppt_id:
            for key, value in fields.items():
                item[key] = copy.deepcopy(value)
            return
    record = {"id": ppt_id}
    for key, value in fields.items():
        record[key] = copy.deepcopy(value)
    history.append(record)

def get_ppt_by_id(ppt_id: str):
    store_item = _ppt_store().get(ppt_id)
    if store_item:
        return store_item
    for item in st.session_state.get("ppt_history", []):
        if item.get("id") == ppt_id:
            return item
    return None

def _current_ppt_item():
    current_id = st.session_state.get("action_target_ppt_id") or st.session_state.get("current_ppt_id")
    if not current_id:
        return None
    return get_ppt_by_id(current_id)

def _current_outline_payload():
    item = _current_ppt_item()
    if item and item.get("outline_payload"):
        return copy.deepcopy(item["outline_payload"])
    outline = st.session_state.get("outline_payload")
    return copy.deepcopy(outline) if outline else None

def _slide_looks_like_thank_you(slide: dict) -> bool:
    if not isinstance(slide, dict):
        return False
    text = " ".join([
        _norm_text(slide.get("title", "")),
        _norm_text(slide.get("subtitle", "")),
    ]).strip()
    if not text:
        return False
    return bool(re.search(r"\b(thank you|thanks|q and a|q a|questions)\b", text))

def _slide_search_text(slide: dict) -> str:
    if not isinstance(slide, dict):
        return ""
    parts = [
        _safe_str(slide.get("title", "")),
        _safe_str(slide.get("subtitle", "")),
    ]
    for item in flatten_slide_content(slide.get("content", []))[:3]:
        parts.append(_safe_str(item))
    return _norm_text(" ".join(parts))

def _resolve_slide_reference_text(ref: str, slides: list) -> Optional[int]:
    """
    Resolve a slide reference like "slide 4", "conclusion slide", or "thank you".
    Returns a 1-based slide index.
    """
    text = _norm_text(ref)
    if not text:
        return None

    num_match = re.search(r"\b(?:slide\s*)?#?(\d+)\b", ref or "", re.IGNORECASE)
    if num_match:
        try:
            idx = int(num_match.group(1))
            if 1 <= idx <= len(slides):
                return idx
        except ValueError:
            pass

    candidates = []
    for i, slide in enumerate(slides or [], start=1):
        if not isinstance(slide, dict):
            continue
        hay = _slide_search_text(slide)
        if not hay:
            continue
        if _slide_looks_like_thank_you(slide) and re.search(r"\b(thank you|thanks|q and a|q a|questions)\b", text):
            candidates.append((100, i))
            continue
        if text == hay or text in hay or hay in text:
            candidates.append((len(hay), i))
            continue
        words = [w for w in text.split() if w not in {"slide", "ppt", "presentation", "deck", "the", "a", "an", "of", "and", "to", "after", "before"}]
        slide_words = [w for w in hay.split() if w not in {"slide", "ppt", "presentation", "deck", "the", "a", "an", "of", "and", "to", "after", "before"}]
        overlap = len(set(words) & set(slide_words))
        if overlap >= 1:
            candidates.append((overlap, i))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]
    return None

def _compute_insert_index(source_idx: int, anchor_idx: int, relation: str) -> int:
    relation = (relation or "").strip().lower()
    if relation == "before":
        return max(0, anchor_idx - 1 - (1 if source_idx < anchor_idx else 0))
    if relation == "after":
        return max(0, anchor_idx - (1 if source_idx < anchor_idx else 0))
    return max(0, anchor_idx - 1)

def _parse_move_swap_request(user_input: str, slides: list) -> Optional[dict]:
    text = user_input or ""
    lower = text.lower()
    if not re.search(r"\b(move|reorder|re-arrange|rearrange|swap|switch positions|change position)\b", lower):
        return None

    slide_count = len(slides or [])
    # Swap requests
    if re.search(r"\b(swap|switch positions|change position of both|exchange)\b", lower):
        refs = re.split(r"(?i)\b(?:swap|exchange|between)\b", text, maxsplit=1)
        if len(refs) == 2:
            tail = refs[1]
            parts = re.split(r"(?i)\b(?:and|with|vs\.?|versus)\b", tail, maxsplit=1)
            if len(parts) == 2:
                a_ref = parts[0].strip(" ,.:;")
                b_ref = parts[1].strip(" ,.:;")
                a_idx = _resolve_slide_reference_text(a_ref, slides)
                b_idx = _resolve_slide_reference_text(b_ref, slides)
                if a_idx and b_idx and a_idx != b_idx:
                    return {
                        "intent": "swap_slides",
                        "slots": {"slide_a": a_idx, "slide_b": b_idx},
                        "missing_slots": [],
                    }

        refs = re.findall(r"\bslide\s+\d+\b", lower)
        if len(refs) >= 2:
            a_idx = _resolve_slide_reference_text(refs[0], slides)
            b_idx = _resolve_slide_reference_text(refs[1], slides)
            if a_idx and b_idx and a_idx != b_idx:
                return {
                    "intent": "swap_slides",
                    "slots": {"slide_a": a_idx, "slide_b": b_idx},
                    "missing_slots": [],
                }
        return {
            "intent": "swap_slides",
            "slots": {},
            "missing_slots": ["slide_a", "slide_b"],
        }

    move_match = re.search(
        r"(?i)\bmove\s+(?P<src>.+?)\s+(?:before|after|to\s+position|to)\s+(?P<dst>.+)$",
        text,
    )
    if move_match:
        src_ref = move_match.group("src").strip(" ,.:;")
        dst_ref = move_match.group("dst").strip(" ,.:;")
        relation = "before" if "before" in lower else "after" if "after" in lower else "position"
        source_idx = _resolve_slide_reference_text(src_ref, slides)
        if relation == "position":
            pos_match = re.search(r"(?i)\bto\s+position\s+(\d+)\b|\bto\s+(\d+)\b", dst_ref)
            if not pos_match:
                pos_match = re.search(r"\b(\d+)\b", dst_ref)
            if source_idx and pos_match:
                try:
                    position = int(next(g for g in pos_match.groups() if g))
                except StopIteration:
                    position = None
                except ValueError:
                    position = None
                if position is not None:
                    return {
                        "intent": "move_slide",
                        "slots": {"slide_number": source_idx, "position": position},
                        "missing_slots": [],
                    }
        else:
            anchor_idx = _resolve_slide_reference_text(dst_ref, slides)
            if source_idx and anchor_idx:
                return {
                    "intent": "move_slide",
                    "slots": {
                        "slide_number": source_idx,
                        "anchor_slide": anchor_idx,
                        "relation": relation,
                    },
                    "missing_slots": [],
                }

        slots = {}
        if source_idx:
            slots["slide_number"] = source_idx
        else:
            slots["source_ref"] = src_ref
        if relation == "position":
            slots["position"] = dst_ref
            missing = ["position"] if source_idx else ["slide_number", "position"]
        else:
            slots["anchor_ref"] = dst_ref
            slots["relation"] = relation
            missing = ["anchor_slide"] if source_idx else ["slide_number", "anchor_slide"]
        return {
            "intent": "move_slide",
            "slots": slots,
            "missing_slots": missing,
        }

    return None

def _move_slide_in_list(slides: list, source_idx: int, relation: str, anchor_idx: Optional[int] = None, position: Optional[int] = None) -> list:
    updated = [copy.deepcopy(s) for s in (slides or []) if isinstance(s, dict)]
    if not (1 <= source_idx <= len(updated)):
        return updated

    slide = updated.pop(source_idx - 1)
    relation = (relation or "").strip().lower()
    if relation == "position" and position is not None:
        insert_at = max(0, min(len(updated), int(position) - 1))
    elif anchor_idx is not None:
        insert_at = _compute_insert_index(source_idx, anchor_idx, relation)
        insert_at = max(0, min(len(updated), insert_at))
    else:
        insert_at = len(updated)

    updated.insert(insert_at, slide)
    return updated

def _resolve_move_swap_slots(intent_data: dict, slides: list) -> Optional[dict]:
    intent = intent_data.get("intent")
    slots = dict(intent_data.get("slots", {}))
    missing = list(intent_data.get("missing_slots", []))

    if intent == "move_slide":
        if slots.get("slide_number") is None and slots.get("source_ref"):
            idx = _resolve_slide_reference_text(slots["source_ref"], slides)
            if idx:
                slots["slide_number"] = idx
        if slots.get("anchor_slide") is None and slots.get("anchor_ref"):
            idx = _resolve_slide_reference_text(slots["anchor_ref"], slides)
            if idx:
                slots["anchor_slide"] = idx
        if slots.get("position") is not None:
            try:
                slots["position"] = int(slots["position"])
            except (ValueError, TypeError):
                pass
        if slots.get("slide_number") and (slots.get("anchor_slide") is not None or slots.get("position") is not None):
            return {"intent": intent, "slots": slots, "missing_slots": []}
        missing = []
        if not slots.get("slide_number"):
            missing.append("slide_number")
        if slots.get("position") is None and slots.get("anchor_slide") is None:
            missing.append("anchor_slide")
        return {"intent": intent, "slots": slots, "missing_slots": missing}

    if intent == "swap_slides":
        if slots.get("slide_a") is None and slots.get("slide_a_ref"):
            idx = _resolve_slide_reference_text(slots["slide_a_ref"], slides)
            if idx:
                slots["slide_a"] = idx
        if slots.get("slide_b") is None and slots.get("slide_b_ref"):
            idx = _resolve_slide_reference_text(slots["slide_b_ref"], slides)
            if idx:
                slots["slide_b"] = idx
        if slots.get("slide_a") and slots.get("slide_b") and slots.get("slide_a") != slots.get("slide_b"):
            return {"intent": intent, "slots": slots, "missing_slots": []}
        missing = []
        if not slots.get("slide_a"):
            missing.append("slide_a")
        if not slots.get("slide_b"):
            missing.append("slide_b")
        return {"intent": intent, "slots": slots, "missing_slots": missing}

    return intent_data

def _dedupe_thank_you_slides(slides: list) -> list:
    cleaned = [copy.deepcopy(s) for s in (slides or []) if isinstance(s, dict)]
    thank_you_indexes = [i for i, slide in enumerate(cleaned) if _slide_looks_like_thank_you(slide)]
    if len(thank_you_indexes) <= 1:
        return cleaned
    last_index = thank_you_indexes[-1]
    thank_you_slide = cleaned[last_index]
    non_thank_you = [slide for i, slide in enumerate(cleaned) if i != last_index and not _slide_looks_like_thank_you(slide)]
    return non_thank_you + [thank_you_slide]

def _build_ppt_record(ppt_id: str, *, outline_payload=None, slides=None, topic=None, sections=None, slide_count=None, ppt_bytes=None, ppt_filename=None, file_path=None):
    store = _ppt_store()
    record = copy.deepcopy(store.get(ppt_id, {"id": ppt_id}))
    if outline_payload is not None:
        record["outline_payload"] = copy.deepcopy(outline_payload)
    if slides is not None:
        record["slides"] = copy.deepcopy(slides)
    if record.get("outline_payload") is None and record.get("slides") is not None:
        record["outline_payload"] = {"slides": copy.deepcopy(record["slides"])}
    if record.get("slides") is None and isinstance(record.get("outline_payload"), dict):
        record["slides"] = copy.deepcopy(record["outline_payload"].get("slides", []))
    if topic is not None:
        record["topic"] = topic
    if sections is not None:
        record["sections"] = sections
    if slide_count is None:
        slide_count = len(record.get("slides", []) or [])
    record["slide_count"] = slide_count
    if ppt_bytes is not None:
        record["ppt_bytes"] = ppt_bytes
    if ppt_filename is not None:
        record["ppt_filename"] = ppt_filename
    if file_path is not None:
        record["file_path"] = file_path
    store[ppt_id] = record
    _upsert_ppt_history(
        ppt_id,
        topic=record.get("topic", ""),
        slide_count=record.get("slide_count", 0),
        sections=record.get("sections", ""),
        outline_payload=record.get("outline_payload", {}),
        slides=record.get("slides", []),
        ppt_bytes=record.get("ppt_bytes"),
        ppt_filename=record.get("ppt_filename"),
        file_path=record.get("file_path"),
    )
    return record

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
    query = query or ""
    if _is_deictic_ppt_reference(query):
        item = get_ppt_by_id(st.session_state.get("current_ppt_id"))
        if item and item.get("id"):
            return item["id"], None
    direct = _resolve_ppt_history_item(query)
    if direct and direct.get("id"):
        return direct["id"], None
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

def _is_deictic_ppt_reference(ref: str) -> bool:
    text = _norm_text(ref)
    return text in {
        "this", "it", "that", "current", "selected", "active", "current ppt",
        "this ppt", "current presentation", "current deck", "active ppt",
        "active presentation", "active deck", "the current", "the current ppt",
    }

def _resolve_deictic_ppt_ref() -> Optional[dict]:
    current_id = st.session_state.get("current_ppt_id")
    if current_id:
        item = get_ppt_by_id(current_id)
        if item:
            return item
    history = st.session_state.get("ppt_history", [])
    if len(history) == 1:
        return history[0]
    return None

def _match_ppt_by_topic_text(query: str) -> Optional[dict]:
    text = _norm_text(query)
    if not text:
        return None
    history = st.session_state.get("ppt_history", [])
    generic_words = {
        "add", "edit", "delete", "remove", "update", "change", "modify",
        "preview", "download", "insert", "slide", "ppt", "presentation",
        "deck", "open", "show", "go", "to", "this", "that", "it", "current",
    }
    text_keywords = [w for w in text.split() if w and w not in generic_words]
    if not text_keywords:
        return None

    exact_matches = []
    partial_matches = []
    for item in history:
        if not isinstance(item, dict):
            continue
        topic = _norm_text(item.get("topic", ""))
        if not topic:
            continue
        topic_keywords = [w for w in topic.split() if w and w not in generic_words]
        overlap = len(set(text_keywords) & set(topic_keywords))
        if topic == text or text in topic or topic in text:
            exact_matches.append((len(topic), item))
        elif overlap >= 2:
            partial_matches.append((overlap, item))
    if exact_matches:
        return sorted(exact_matches, key=lambda x: x[0], reverse=True)[0][1]
    if partial_matches:
        return sorted(partial_matches, key=lambda x: x[0], reverse=True)[0][1]
    return None

def _is_preview_request(user_input: str) -> bool:
    text = user_input or ""
    if re.search(r"\bslide\s*\d+\b", text, re.IGNORECASE):
        return False
    return bool(re.search(
        r"\b(?:show|open|view|get|display|preview)\b.*\b(?:preview|ppt|presentation|deck)\b|"
        r"\bpreview\b",
        text,
        re.IGNORECASE,
    ))

def _extract_slide_navigation_target(user_input: str) -> Optional[int]:
    text = user_input or ""
    patterns = [
        r"\b(?:go to|jump to|open|show|switch to|focus on)\s+slide\s+(\d+)\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                return None
    return None

def _extract_deck_refinement_request(user_input: str) -> Optional[dict]:
    text = (user_input or "").strip()
    if not text:
        return None
    lower = text.lower()
    tone = None
    style_hints = []
    if re.search(r"\b(?:more\s+)?professional\b|\bformal\b", lower):
        tone = "Professional"
        style_hints.append("Make the deck more polished, businesslike, and concise.")
    if re.search(r"\bcasual\b|\bfriendly\b|\bconversational\b", lower):
        tone = "Creative"
        style_hints.append("Make the deck feel casual, friendly, and conversational.")
    if re.search(r"\b(?:shorter|less text|concise|clean|clearer|simpler|important points only|key points only|remove fluff)\b", lower):
        style_hints.append("Use shorter slides with only the most important points.")
    if re.search(r"\bremove everything\b", lower):
        style_hints.append("Remove extra text and leave only the core content.")
    if not style_hints and not tone:
        return None
    return {
        "tone": tone,
        "style_hint": " ".join(style_hints).strip(),
    }

def _is_cancel_command(user_input: str) -> bool:
    text = _safe_str(user_input)
    return bool(re.fullmatch(r"(?:stop|cancel|nevermind|never mind|abort|quit)", text, re.IGNORECASE))

def _is_affirmative_command(user_input: str) -> bool:
    text = _safe_str(user_input)
    return bool(re.fullmatch(r"(?:yes|yep|yeah|sure|ok|okay|go ahead|do it|please do|continue)", text, re.IGNORECASE))

def _is_negative_command(user_input: str) -> bool:
    text = _safe_str(user_input)
    return bool(re.fullmatch(r"(?:no|nope|not now|don't|do not|cancel)", text, re.IGNORECASE))

_NUMBER_WORDS = {
    "zero": 0, "one": 1, "a": 1, "an": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    "thirty": 30, "forty": 40, "fifty": 50,
}

def _parse_small_number_token(token: str) -> Optional[int]:
    token = _safe_str(token).lower()
    if not token:
        return None
    if token.isdigit():
        return int(token)
    return _NUMBER_WORDS.get(token)

def _is_null_add_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.fullmatch(r"(?:add nothing|nothing to add|no change|no changes|nothing)", text, re.IGNORECASE))

def _is_clear_slide_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.search(
        r"\b(?:remove everything|clear slide|clear all|erase all|wipe slide|wipe all|delete all content|remove all content)\b",
        text,
        re.IGNORECASE,
    ))

def _is_blank_slide_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.search(r"\b(?:make blank slide|blank slide|empty slide)\b", text, re.IGNORECASE))

def _extract_point_count(change_content: str) -> Optional[int]:
    text = (change_content or "").lower().strip()
    if not text or "add" not in text:
        return None
    text = text.replace("point(s)", "points").replace("bullet(s)", "bullets")
    num_match = re.search(r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)\s*(?:bullet\s*)?(?:point|points|bullet|bullets|item|items)\b", text)
    if num_match:
        parsed = _parse_small_number_token(num_match.group(1))
        return max(1, parsed or 1)
    if re.search(r"\b(?:one|a|an)\s+(?:bullet\s*)?(?:point|points|bullet|bullets)\b", text):
        return 1
    if re.search(r"\b(?:point|points|bullet|bullets)\b", text):
        return 2
    return None

def _blank_slide_payload(slide: dict) -> dict:
    if not isinstance(slide, dict):
        return {}
    cleared = copy.deepcopy(slide)
    cleared["title"] = ""
    cleared["subtitle"] = ""
    cleared["layout"] = "bullets"
    cleared["icon"] = "▸"
    cleared["content"] = []
    cleared["left_points"] = []
    cleared["right_points"] = []
    cleared["left"] = []
    cleared["right"] = []
    cleared["steps"] = []
    cleared["grid_items"] = []
    cleared["table_columns"] = []
    cleared["table_rows"] = []
    cleared["stat"] = ""
    cleared["stat_label"] = ""
    cleared["stat_source"] = ""
    cleared["chart_title"] = ""
    cleared["chart_data"] = []
    cleared["chart_source"] = ""
    cleared["company"] = ""
    cleared["result"] = ""
    cleared["metrics"] = []
    cleared["_user_modified"] = True
    return cleared

def _clear_slide_payload(slide: dict) -> dict:
    cleared = _blank_slide_payload(slide)
    return cleared

def switch_active_ppt(ppt_id: str) -> bool:
    item = get_ppt_by_id(ppt_id)
    if not item:
        return False
    current_id = st.session_state.get("current_ppt_id")
    outline = copy.deepcopy(item.get("outline_payload") or {})
    slides = copy.deepcopy(item.get("slides") or outline.get("slides", []))
    if outline and "slides" not in outline:
        outline["slides"] = copy.deepcopy(slides)
    if slides:
        slides = _dedupe_thank_you_slides(slides)
        outline["slides"] = copy.deepcopy(slides)
    elif outline.get("slides"):
        outline["slides"] = _dedupe_thank_you_slides(outline.get("slides", []))
    if current_id and current_id != ppt_id:
        _remember_ppt_history_switch(current_id, ppt_id)
        st.session_state.last_ppt = current_id
    st.session_state.current_ppt_id = ppt_id
    st.session_state.outline_payload = outline
    st.session_state.ppt_bytes = item.get("ppt_bytes")
    st.session_state.ppt_filename = item.get("ppt_filename")
    st.session_state.topic = item.get("topic", "")
    st.session_state.sections = item.get("sections", "")
    st.session_state.num_slides = item.get("slide_count", st.session_state.num_slides)
    st.session_state.slide_count = st.session_state.num_slides
    if outline.get("slides"):
        sync_all_editor_widgets(copy.deepcopy(outline.get("slides", [])))
    st.session_state.pending_intent = None
    st.session_state.pending_action = None
    st.session_state.last_ppt_id = ppt_id
    st.session_state.last_slide_index = None
    st.session_state.last_action_type = None
    st.session_state.active_slide_index = None
    _build_ppt_record(
        ppt_id,
        outline_payload=outline,
        slides=slides or outline.get("slides", []),
        topic=st.session_state.topic,
        sections=st.session_state.sections,
        slide_count=len(slides or outline.get("slides", [])),
        ppt_bytes=item.get("ppt_bytes"),
        ppt_filename=item.get("ppt_filename"),
        file_path=item.get("file_path"),
    )
    _set_conv_state(
        active_ppt_id=ppt_id,
        active_slide_index=None,
        last_topic=item.get("topic", ""),
        last_action_type=None,
        pending_create=None,
        pending_edit=None,
        turn_summary=f"Switched to {_ppt_display_label(ppt_id)}",
    )
    return True

def resolve_ppt_reference(user_input: str) -> Tuple[bool, Optional[str], Optional[str]]:
    explicit_ref = _extract_explicit_ppt_ref_from_text(user_input)
    if explicit_ref:
        item = _resolve_ppt_history_item(explicit_ref)
        if item and item.get("id"):
            new_id = item["id"]
            if new_id != st.session_state.get("current_ppt_id"):
                switch_active_ppt(new_id)
            return True, new_id, None

    if _is_previous_ppt_request(user_input):
        last_ppt = _get_previous_ppt_id()
        if last_ppt and get_ppt_by_id(last_ppt):
            switch_active_ppt(last_ppt)
            return True, last_ppt, None
        return False, None, "No previous presentation found."

    override_patterns = [
        r"(?:in|to|switch to|open)\s+(?:the\s+)?(?:ppt|presentation|deck)[_\s-]*['\"]?([^'\"]+)['\"]?",
        r"(?:do this in|apply to)\s+(\d+(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck))",
        r"(?:go to|select)\s+(?:ppt|presentation|deck)[_\s-]*(\d+)",
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
    if re.search(r"(?:which ppt|what about|show me|open|switch to)", user_input, re.IGNORECASE):
        best_id, clarification = find_ppt_semantically(user_input, disambiguate=True)
        if clarification:
            return False, None, clarification
        if best_id and best_id != st.session_state.get("current_ppt_id"):
            switch_active_ppt(best_id)
            return True, best_id, None
    return False, None, None

def _ppt_display_label(ppt_id: str) -> str:
    if not ppt_id:
        return "ppt"
    m = re.search(r"ppt_(\d+)", str(ppt_id), re.IGNORECASE)
    if m:
        return f"ppt {m.group(1)}"
    return str(ppt_id).replace("_", " ")

def _resolve_ppt_history_item(ppt_ref) -> Optional[dict]:
    """
    Resolve a deck reference like "ppt 4", "4th ppt", or a semantic title.
    Returns the matching history item or None.
    """
    if ppt_ref is None:
        return None
    ref = str(ppt_ref).strip()
    if not ref:
        return None

    if _is_deictic_ppt_reference(ref):
        return _resolve_deictic_ppt_ref()

    num_match = re.search(r"\b(\d+)\b", ref)
    if num_match:
        idx = int(num_match.group(1)) - 1
        history = st.session_state.get("ppt_history", [])
        if 0 <= idx < len(history):
            return history[idx]

    named = _find_named_ppt_in_text(ref)
    if named:
        return named

    topic_match = _match_ppt_by_topic_text(ref)
    if topic_match:
        return topic_match
    return None

def _find_named_ppt_in_text(user_input: str) -> Optional[dict]:
    """
    Resolve a deck by name/topic mentioned in free-form text.
    Prefers exact/substring matches against stored topics before semantic search.
    """
    text = _norm_text(user_input)
    if not text:
        return None
    text_words = [w for w in text.split() if w]
    if len(text_words) < 3:
        return None
    generic_words = {
        "add", "edit", "delete", "remove", "update", "change", "modify",
        "preview", "download", "insert", "slide", "ppt", "presentation",
        "deck", "thank", "thanks", "you", "at", "the", "end", "new",
    }
    text_keywords = [w for w in text_words if w not in generic_words]
    if len(text_keywords) < 2:
        return None

    history = st.session_state.get("ppt_history", [])
    exact_matches = []
    partial_matches = []
    for item in history:
        if not isinstance(item, dict):
            continue
        topic = _norm_text(item.get("topic", ""))
        if not topic:
            continue
        topic_words = [w for w in topic.split() if w and w not in generic_words]
        overlap = len(set(text_keywords) & set(topic_words))
        if topic in text:
            exact_matches.append((len(topic), item))
        elif text in topic and overlap >= 2:
            partial_matches.append((len(topic), item))
        elif overlap >= 2:
            partial_matches.append((overlap, item))
    if exact_matches:
        return sorted(exact_matches, key=lambda x: x[0], reverse=True)[0][1]
    if partial_matches:
        return sorted(partial_matches, key=lambda x: x[0], reverse=True)[0][1]
    return None

def resolve_named_ppt_context(user_input: str) -> Tuple[bool, Optional[str], Optional[str]]:
    """
    If the user names a deck in a modification request, switch active context
    to that deck before the action executes.
    """
    if not user_input:
        return False, None, None
    # Exact references like "ppt 2" should never be overridden by fuzzy deck matching.
    if _extract_explicit_ppt_ref_from_text(user_input):
        return False, None, None
    if not re.search(r"\b(add|edit|delete|remove|update|change|modify|preview|download|insert|switch|open|go to|go back|show)\b", user_input, re.IGNORECASE):
        return False, None, None
    item = _find_named_ppt_in_text(user_input)
    if item and item.get("id") != st.session_state.get("current_ppt_id"):
        switch_active_ppt(item["id"])
        return True, item["id"], None
    return False, None, None

def _is_thank_you_slide_request(user_text: str) -> bool:
    text = _norm_text(user_text)
    if not text:
        return False
    return bool(re.search(r"\b(thank you|thanks|thank you slide|q and a|q a|questions)\b", text))

def _is_slide_level_request(user_input: str) -> bool:
    return bool(re.search(r"\bslide\s*\d+\b", user_input or "", re.IGNORECASE))

def _is_deck_level_edit_request(user_input: str) -> bool:
    text = user_input or ""
    edit_verbs = r"(?:edit|change|update|modify|revise|adjust)"
    deck_ref = r"(?:ppt|presentation|deck)\s*\d+"
    return bool(
        re.search(rf"\b{edit_verbs}\b.*\b{deck_ref}\b", text, re.IGNORECASE)
        and not _is_slide_level_request(text)
    )

def _is_topic_ideas_request(user_input: str) -> bool:
    text = user_input or ""
    return bool(re.search(
        r"\b("
        r"topic ideas|topic idea|suggest topic|suggest topics|new topic|new topics|"
        r"brainstorm topic|brainstorm topics|trending topic|trending topics|"
        r"hot topic|hot topics|topic suggestions|presentation topics"
        r")\b",
        text,
        re.IGNORECASE,
    ))

def _is_ppt_topic_lookup_request(user_input: str) -> bool:
    text = user_input or ""
    if _is_topic_ideas_request(text):
        return False
    if not re.search(r"\b(topic|title|name)\b", text, re.IGNORECASE):
        return False
    return bool(re.search(
        r"\b(?:give|show|tell|what|which|get|share|list)\b.*\btopic\b|"
        r"\btopic of\b|"
        r"\bwhat is(?: the)? topic\b|"
        r"\bmy generated ppt\b|"
        r"\bour generated ppt\b|"
        r"\bcurrent ppt\b|"
        r"\bthis ppt\b|"
        r"\bthe ppt\b",
        text,
        re.IGNORECASE,
    ))

def _extract_explicit_ppt_ref_from_text(user_input: str) -> Optional[str]:
    """
    Extract the most likely PPT reference from free-form text.

    Handles forms like:
    - ppt 1
    - ppt_1
    - ppt-1
    - ppt1
    - 1st ppt
    """
    text = user_input or ""
    if not text:
        return None

    ref_patterns = [
        r"\b(?:ppt|presentation|deck)[_\s-]*#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck)\b",
    ]

    matches = []
    for pat in ref_patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            matches.append((m.start(), f"ppt {m.group(1)}"))

    if not matches:
        return None

    action_patterns = [
        r"\badd\b", r"\binsert\b", r"\bedit\b", r"\bchange\b", r"\bupdate\b",
        r"\bmodify\b", r"\brevise\b", r"\bpreview\b", r"\bdownload\b",
        r"\bswitch\b", r"\bopen\b", r"\bgo to\b", r"\bgo back\b", r"\bshow\b",
    ]
    action_matches = []
    for pat in action_patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            action_matches.append(m.start())

    if action_matches:
        pivot = max(action_matches)
        after = [item for item in matches if item[0] >= pivot]
        if after:
            return after[0][1]
        before = [item for item in matches if item[0] < pivot]
        if before:
            return before[-1][1]

    return matches[-1][1]

def _same_topic(a: str, b: str) -> bool:
    left = _norm_text(a)
    right = _norm_text(b)
    if not left or not right:
        return False
    if left == right:
        return True
    left_words = [w for w in left.split() if w]
    right_words = [w for w in right.split() if w]
    if not left_words or not right_words:
        return False
    # Short topics like "AI" should only match on exact equality.
    if len(left_words) <= 2 or len(right_words) <= 2:
        return False
    stopwords = {"the", "a", "an", "of", "and", "in", "on", "for", "to", "with", "new", "ppt", "presentation", "deck"}
    left_set = {w for w in left_words if w not in stopwords}
    right_set = {w for w in right_words if w not in stopwords}
    if not left_set or not right_set:
        return False
    overlap = len(left_set & right_set)
    smaller = min(len(left_set), len(right_set))
    return overlap >= max(2, int(smaller * 0.75))

def _extract_create_ppt_topic(user_input: str) -> Optional[str]:
    """
    Extract a topic from requests like:
    - make ppt on AI
    - create presentation about blockchain
    - generate deck for healthcare
    """
    text = (user_input or "").strip()
    if not text:
        return None

    if not re.search(r"\b(?:make|create|generate|build)\b", text, re.IGNORECASE):
        return None
    if re.search(r"\b(?:add|edit|update|change|modify|delete|remove|continue)\b", text, re.IGNORECASE):
        # Avoid hijacking edit flows like "add slide about AI"
        if not re.search(r"\b(?:ppt|presentation|deck)\s+(?:on|about|for)\b", text, re.IGNORECASE):
            return None

    patterns = [
        r"\b(?:make|create|generate|build)\s+(?:a\s+)?(?:new\s+)?(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slides?\s+(?:ppt|presentation|deck)\s*(?:on|about|for)\s+(.+)$",
        r"\b(?:make|create|generate|build)\s+(?:a\s+)?(?:new\s+)?(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slide\s+(?:ppt|presentation|deck)\s*(?:on|about|for)\s+(.+)$",
        r"\b(?:make|create|generate|build)\s+(?:a\s+)?(?:new\s+)?(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slides?\s+(?:on|about|for)\s+(.+)$",
        r"\b(?:make|create|generate|build)\s+(?:a\s+)?(?:new\s+)?(?:ppt|presentation|deck)\s*(?:on|about|for)\s+(.+)$",
        r"\b(?:make|create|generate|build)\s+(?:a\s+)?(?:new\s+)?(?:ppt|presentation|deck)\s+(.+)$",
        r"\b(?:ppt|presentation|deck)\s*(?:on|about|for)\s+(.+)$",
    ]
    topic = None
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            topic = m.group(1).strip()
            break
    if not topic:
        return None

    topic = re.sub(r"\bwith\s+\d+\s+slides?\b.*$", "", topic, flags=re.IGNORECASE).strip()
    topic = re.sub(r"\bfor\s+\d+\s+slides?\b.*$", "", topic, flags=re.IGNORECASE).strip()
    topic = re.sub(r"\b(?:including|containing)\s+\d+\s+slides?\b.*$", "", topic, flags=re.IGNORECASE).strip()
    topic = re.sub(
        r"^\s*(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slides?\s+",
        "",
        topic,
        flags=re.IGNORECASE,
    ).strip()
    topic = re.split(
        r"(?i)\b(?:add\s+below\s+content|include\s+below\s+content|content\s*:|with\s+content|add\s+content|include\s+content)\b",
        topic,
        maxsplit=1,
    )[0].strip()
    topic = re.sub(r"[\.\!\?]+$", "", topic).strip()
    topic = topic.strip(' "\'“”‘’')
    topic = re.sub(r"\s+", " ", topic).strip()
    topic = re.sub(r"\s+(?:for\s+me|please|pls)$", "", topic, flags=re.IGNORECASE).strip()
    topic = re.sub(r"^(?:for\s+me|me|myself|please|pls)$", "", topic, flags=re.IGNORECASE).strip()
    if orchestrator_is_generic_presentation_context(topic):
        return None
    return topic or None

def _is_generic_presentation_context(value: str) -> bool:
    return orchestrator_is_generic_presentation_context(value)

def _extract_conversational_create_topic(user_input: str) -> Optional[str]:
    text = _safe_str(user_input, "")
    if not text:
        return None
    explicit = _extract_create_ppt_topic(text)
    pending = st.session_state.get("pending_new_ppt") or {}
    return orchestrator_extract_conversational_create_topic(
        text,
        explicit_topic=explicit,
        pending_need_topic=bool(pending.get("need_topic")),
    )

def _is_implicit_create_request(user_input: str) -> bool:
    return orchestrator_is_implicit_create_request(user_input)

def _create_topic_followup(user_input: str = "") -> str:
    return orchestrator_create_topic_followup(user_input)

def _ask_for_create_topic(user_input: str, question_override: Optional[str] = None) -> None:
    pending = st.session_state.get("pending_new_ppt") or {}
    pending.update({
        "topic": "",
        "sections": pending.get("sections") or _extract_create_sections_from_prompt(user_input),
        "theme_colors": pending.get("theme_colors") or extract_theme_colors_from_messages(st.session_state.get("messages", [])),
        "need_topic": True,
    })
    st.session_state.pending_new_ppt = pending
    st.session_state.pending_intent = {
        "intent": "create_ppt",
        "slots": {"topic": None, "slide_count": None, "sections": pending.get("sections")},
        "missing_slots": ["topic"],
        "next_question": question_override or _create_topic_followup(user_input),
        "action": "ask",
    }
    st.session_state.pending_action = st.session_state.pending_intent
    _set_conv_state(pending_create=st.session_state.pending_new_ppt)
    question = question_override or _create_topic_followup(user_input)
    with st.chat_message("assistant"):
        st.markdown(question)
    add_message("assistant", question)
    st.rerun()

def _extract_create_sections_from_prompt(user_input: str) -> Optional[str]:
    text = (user_input or "").strip()
    if not text:
        return None
    markers = [
        r"(?i)\badd below content\s*:\s*(.+)$",
        r"(?i)\bcontent\s*:\s*(.+)$",
        r"(?i)\binclude below content\s*:\s*(.+)$",
    ]
    for pat in markers:
        m = re.search(pat, text, re.DOTALL)
        if m:
            content = m.group(1).strip()
            content = re.sub(r"\s+", " ", content).strip()
            return content or None
    return None

def _is_explicit_create_request(user_input: str) -> bool:
    text = user_input or ""
    if not re.search(r"\b(?:make|create|generate|build)\b", text, re.IGNORECASE):
        return False
    return bool(re.search(r"\b(?:ppt|presentation|deck)\b", text, re.IGNORECASE))

def _handle_create_request(prompt: str) -> bool:
    """
    Return True if the prompt was fully handled as a create-new-PPT request.
    """
    topic = _extract_create_ppt_topic(prompt)
    if not topic:
        _ask_for_create_topic(prompt)
        return True
    slide_count = _extract_slide_count_from_text(prompt)
    theme_colors = extract_theme_colors_from_messages(st.session_state.get("messages", [])) or extract_theme_colors(topic)
    sections = _extract_create_sections_from_prompt(prompt)

    st.session_state.pending_new_ppt = None
    st.session_state.pending_intent = None
    st.session_state.pending_action = None
    if slide_count:
        if int(slide_count) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
            question = _slide_limit_confirmation_question(int(slide_count))
            pending = {
                "intent": "confirm_large_slide_count",
                "target_intent": "create_ppt",
                "slots": {"topic": topic, "slide_count": int(slide_count), "sections": sections},
                "next_question": question,
                "action": "confirm",
            }
            _store_pending_action(pending)
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        generate_outline_and_reply(topic, slide_count, st.session_state.tone, sections, theme_colors)
    else:
        st.session_state.pending_new_ppt = {
            "topic": topic,
            "sections": sections,
            "theme_colors": theme_colors,
        }
        _set_conv_state(pending_create=st.session_state.pending_new_ppt)
        with st.chat_message("assistant"):
            st.markdown("How many slides would you like the presentation to have?")
        add_message("assistant", "How many slides would you like the presentation to have?")
        st.rerun()
    return True

def _extract_slide_count_from_text(user_input: str) -> Optional[int]:
    text = user_input or ""
    patterns = [
        r"\b(?:with|of|for|only)\s+(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slides?\b",
        r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)[\s-]+slides?\b",
        r"\bslide\s+(\d+)\b",
        r"\bslides?\s+of\s+(\d+)\b",
        r"\b(\d+)[\s-]+slide(?:s)?\b",
        r"\b(\d+)[\s-]+slide(?:s)?\s+(?:presentation|ppt|deck)\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return _parse_small_number_token(m.group(1))
    if re.fullmatch(r"\d+", text.strip()):
        try:
            return int(text.strip())
        except ValueError:
            return None
    return None

def _extract_explicit_slide_number_from_text(user_input: str) -> Optional[int]:
    """
    Extract the slide number only when the user explicitly names a slide.
    This is used to override stale slide context from earlier turns.
    """
    text = user_input or ""
    patterns = [
        r"\bslide\s+(\d+)\b",
        r"\bslide\s+#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                return None
    return None

def _has_explicit_slide_reference(user_input: str) -> bool:
    text = user_input or ""
    return bool(re.search(r"\bslide\s*\d+\b", text, re.IGNORECASE))

def _is_ambiguous_add_followup(user_input: str) -> bool:
    text = _norm_text(user_input)
    if not text:
        return False
    if _has_explicit_slide_reference(text):
        return False
    if re.search(r"\b(?:ppt|presentation|deck)\s*\d+\b", text, re.IGNORECASE):
        return False
    # Only rewrite true bullet/point follow-ups here.
    # Generic "add slide" requests should keep their own intent so they can ask
    # the user for slide content or position instead of reusing the last edit.
    return bool(re.search(
        r"\b(?:add|another|one more|more|extra|continue)\b.*\b(?:point|bullet|bullets|points?)\b",
        text,
        re.IGNORECASE,
    ))

def _is_bare_add_slide_request(user_input: str) -> bool:
    text = user_input or ""
    if not text:
        return False
    lower = text.lower()
    if not re.search(r"\b(?:add|insert|new)\b", lower):
        return False
    if not re.search(r"\bslide\b", lower):
        return False
    if re.search(r"\bslide\s+\d+\b", lower):
        return False
    if re.search(r"\b(?:point|points|bullet|bullets)\b", lower):
        return False
    return bool(re.search(
        r"\b(?:add|insert)\s+(?:a\s+)?(?:new\s+|another\s+|one more\s+)?slide\b|"
        r"\bnew\s+slide\b|"
        r"\banother\s+slide\b|"
        r"\bone more\s+slide\b",
        lower,
        re.IGNORECASE,
    ))

def _is_add_slide_request(user_input: str) -> bool:
    text = user_input or ""
    if not text:
        return False
    if _has_explicit_slide_reference(text) and not re.search(r"\b(?:after|before)\s+slide\s*#?\d+\b", text, re.IGNORECASE):
        return False
    if re.search(r"\b(?:point|points|bullet|bullets)\b", text, re.IGNORECASE) and not _extract_explicit_provided_content(text):
        return False
    return bool(re.search(
        r"\b(?:add|insert)\b.*\b(?:new\s+|another\s+|one more\s+)?slide\b"
        r"|\b(?:create|make|build|generate)\b.*\bnew\s+slide\b"
        r"|\b(?:new|another|one more)\s+slide\b",
        text,
        re.IGNORECASE,
    ))

def _is_new_structured_slide_request(user_input: str) -> bool:
    """
    Treat "add a comparison table/chart" as a request for a new slide unless
    the user explicitly targets an existing slide.
    """
    text = _safe_str(user_input, "")
    if not text:
        return False
    if re.search(r"\b(?:to|on|in|into|within|for)\s+(?:the\s+)?slide\s*#?\d+\b", text, re.IGNORECASE):
        return False
    if re.search(r"\bslide\s*#?\d+\b", text, re.IGNORECASE) and not re.search(r"\b(?:after|before)\s+slide\s*#?\d+\b", text, re.IGNORECASE):
        return False
    return bool(re.search(
        r"\b(?:add|insert|create|make|build|generate)\b.{0,40}\b(?:comparison\s+table|table|chart|graph|matrix)\b",
        text,
        re.IGNORECASE,
    ))

def _structured_slide_content_from_request(user_input: str) -> str:
    text = _safe_str(user_input, "").strip()
    left, right = _extract_comparison_subjects(text, None)
    if re.search(r"\bcomparison|compare|vs|versus\b", text, re.IGNORECASE):
        return f"comparison table between {left} and {right}"
    if re.search(r"\bchart|graph\b", text, re.IGNORECASE):
        return text
    return text or "structured table"

def _is_bare_edit_slide_request(user_input: str) -> bool:
    text = user_input or ""
    if not text:
        return False
    if _has_meaningful_edit_instruction(text):
        return False
    lower = text.lower()
    if not re.search(r"\b(?:edit|change|update|modify|revise)\b", lower):
        return False
    if not re.search(r"\bslide\b", lower):
        return False
    if re.search(r"\bslide\s+\d+\b", lower):
        return False
    return True

def _is_position_only_text(text: str) -> bool:
    lower = _norm_text(text)
    if not lower:
        return False
    # Accept broader set of position-only phrases and synonyms.
    if re.fullmatch(r"(?:start|first|top|beginning|end|last|final|bottom|finish|append|append\s+to\s+end)", lower):
        return True
    if re.fullmatch(r"(?:after|before)\s+slide\s+\d+", lower):
        return True
    if re.fullmatch(r"(?:after|before)\s+\d+", lower):
        return True
    if re.fullmatch(r"(?:at\s+)?(?:the\s+)?(?:start|end|top|bottom|final)", lower):
        return True
    if re.fullmatch(r"(?:after|after\s+the|after\s+last|after\s+all|at\s+the\s+end)", lower):
        return True
    return False

def _extract_add_slide_position(text: str, slides: list) -> Optional[object]:
    raw = str(text or "").strip()
    if not raw:
        return None
    lower = raw.lower()
    # Deck references like "in ppt 2" should not be treated as slide positions.
    lower = re.sub(r"\b(?:ppt|presentation|deck)\s*#?\d+\b", "", lower, flags=re.IGNORECASE).strip()
    lower = re.sub(r"\b\d+(?:st|nd|rd|th)?\s+(?:ppt|presentation|deck)\b", "", lower, flags=re.IGNORECASE).strip()
    # Normalise common synonyms for 'end' and 'start'
    if re.search(r"\b(?:end|last|bottom|final|append|append to end|at the end|to the end|after all|after the last)\b", lower):
        return "end"
    if re.search(r"\b(?:start|first|top|beginning)\b", lower):
        return "start"
    after_match = re.search(r"\bafter\s+slide\s+(\d+)\b", lower)
    if after_match:
        return f"after slide {after_match.group(1)}"
    before_match = re.search(r"\bbefore\s+slide\s+(\d+)\b", lower)
    if before_match:
        return f"before slide {before_match.group(1)}"
    pos_match = re.search(r"^\s*(\d+)\s*$|\b(?:at\s+)?position\s+(\d+)\b|\bat\s+(\d+)\b", lower)
    if pos_match:
        try:
            pos = int(next(group for group in pos_match.groups() if group))
            return max(1, min((len(slides or []) + 1), pos))
        except ValueError:
            return None
    return None

def _strip_add_slide_position_terms(text: str) -> str:
    cleaned = re.sub(r"(?i)\b(?:at\s+)?(?:the\s+)?(?:end|last|bottom|start|first|top|beginning)\b", "", text or "")
    cleaned = re.sub(r"(?i)\b(?:after|before)\s+slide\s+\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:after|before)\s+\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:position\s+)?\d+\b", "", cleaned)
    # Remove trailing/leading prepositions left after stripping position phrases,
    # e.g. "at the end of ppt add summary" -> remove "of ppt" or leading "of".
    cleaned = re.sub(r"(?i)\b(?:of|in|on|for)\b\s*(?:the\s+)?(?:ppt|presentation|deck)(?:\s*#?\d+)?\b", "", cleaned)
    cleaned = re.sub(r"(?i)^\s*(?:of|in|on|for)\b\s*", "", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip(" ,.;:-")
    return cleaned.strip()

def _extract_add_slide_content(user_input: str) -> Optional[str]:
    """
    Pull the actual slide topic/content out of an add-slide request.

    This keeps prompts like "add a new slide about AI at the end" from being
    treated as if they contained no content at all.
    """
    text = (user_input or "").strip()
    if not text:
        return None

    text = _strip_add_slide_position_terms(text)
    # 1. Handle "add [Topic] slide" pattern (e.g., "add summary slide")
    # This must be tried before stripping the word "slide"
    topic_slide_match = re.search(r"(?i)\b(?:add|insert|create|make|generate)\s+(?:a\s+|an\s+|the\s+)?(?:new\s+|another\s+|one more\s+)?(.+?)\s+slide\b", text)
    if topic_slide_match:
        text = topic_slide_match.group(1).strip()
    # 2. Handle "add slide about [Topic]" or just "add slide [Topic]"
    text = re.sub(
        r"(?i)^\s*(?:please\s+)?(?:add|insert|create|make|generate)\s+"
        r"(?:a\s+|an\s+|the\s+)?(?:new\s+|another\s+|one more\s+)?slide(?:\s+with\s+content|\s+content)?\s*",
        "",
        text,
    ).strip()
    
    text = re.sub(r"(?i)^\s*(?:with\s+content|content)\s*[:\-]?\s*", "", text).strip()
    text = re.sub(r"(?i)^\s*(?:about|on|for)\s+", "", text).strip()
    text = text.strip(" ,.;:-")
    return text or None

def _extract_explicit_provided_content(user_input: str) -> Optional[str]:
    text = (user_input or "").strip()
    if not text:
        return None
    patterns = [
        r"(?is)\b(?:add|insert|include|use)\s+(?:the\s+)?(?:below|following)\s+content\s*[:\-]?\s*(.+)$",
        r"(?is)\b(?:content|text)\s*(?:is|:|-)\s*(.+)$",
        r"(?is)\bwith\s+(?:this\s+)?(?:exact\s+)?content\s*[:\-]?\s*(.+)$",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            content = m.group(1).strip()
            return content or None
    return None

def _slide_from_explicit_content(content: str, fallback_title: str = "User Provided Content") -> dict:
    lines = [clean_icon_tokens(line) for line in re.split(r"[\r\n]+", content or "") if clean_icon_tokens(line)]
    if not lines:
        lines = flatten_slide_content(content)
    title = fallback_title
    bullets = lines
    if lines:
        first = lines[0].strip()
        if len(first.split()) <= 10 and not first.endswith("."):
            title = first.strip(" :-")
            bullets = lines[1:] or [first]
    if len(bullets) == 1:
        pieces = [clean_icon_tokens(p) for p in re.split(r"\s*(?:;|•|\u2022)\s*", bullets[0]) if clean_icon_tokens(p)]
        if len(pieces) > 1:
            bullets = pieces
    return normalize_slide({
        "title": title or fallback_title,
        "subtitle": "",
        "layout": "bullets",
        "icon": "▸",
        "content": bullets or [clean_icon_tokens(content)],
        "style": {},
        "_user_modified": True,
    })

def _extract_edit_change_content(user_input: str) -> Optional[str]:
    """
    Pull the actual edit instruction out of a follow-up message.

    This lets requests like "edit slide 3 to make it shorter" skip an extra
    clarification step when the user already included the change text.
    """
    text = (user_input or "").strip()
    if not text:
        return None
    cleaned = re.sub(r"(?i)\b(?:sorry|nope|actually|wait|correction|instead)\b", "", text)
    cleaned = re.sub(
        r"(?i)^\s*(?:i\s+want\s+to|i\s+would\s+like\s+to|i\s+need\s+to|can\s+you|could\s+you|please)\s+",
        "",
        cleaned,
    )
    cleaned = re.sub(r"(?i)\b(?:in|of|for)\s+(?:ppt|presentation|deck)\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:ppt|presentation|deck)\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\bslide\s+#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)^\s*(?:please\s+)?(?:edit|change|update|modify|revise|improve|make|fix)\s+", "", cleaned)
    cleaned = re.sub(r"(?i)^\s*(?:edit|change|update|modify|revise)\s+", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:in|of|for)\s*$", "", cleaned)
    cleaned = cleaned.strip(" ,.;:-")
    norm = _norm_text(cleaned)
    if not norm or norm in {"slide", "slides"}:
        return None
    if re.fullmatch(r"(?:i\s+want\s+to|i\s+would\s+like\s+to|i\s+need\s+to)?\s*(?:edit|change|update|modify|revise)?", norm, re.IGNORECASE):
        return None
    if len(norm.split()) < 2 and not re.search(
        r"\b(?:add|remove|replace|title|subtitle|content|bullet|point|layout|chart|table|timeline|grid|stat|company|result|metric|summary|conclusion|thank you|questions?)\b",
        norm,
        re.IGNORECASE,
    ):
        return None
    return cleaned

def _is_correction_message(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.search(r"\b(?:sorry|no|nope|actually|wait|correction|instead)\b", text, re.IGNORECASE))

def _is_transform_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    if not text:
        return False
    return bool(re.search(
        r"\b(?:improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|condense|make it better|make it shorter|make it professional)\b",
        text,
        re.IGNORECASE,
    ))

def _is_expand_existing_slide_request(user_input: str) -> bool:
    text = user_input or ""
    if not _has_explicit_slide_reference(text):
        return False
    return bool(re.search(
        r"\b(?:longer|more detailed|more detail|expand|expanded|elaborate|add detail|add details)\b",
        text,
        re.IGNORECASE,
    ))

def _is_shortening_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.search(r"\b(shorter|less text|concise|condense|trim|remove fluff|cut down|tighten)\b", text, re.IGNORECASE))

def _is_generic_edit_request(user_input: str) -> bool:
    text = _norm_text(user_input)
    return bool(re.fullmatch(
        r"(?:edit|change|update|modify|revise)\s+slide(?:\s+\d+)?"
        r"|(?:edit|change|update|modify|revise)\s+slide\s+#?\d+"
        r"|(?:edit|change|update|modify|revise)\s+it"
        r"|(?:edit|change|update|modify|revise)\s+this",
        text,
        re.IGNORECASE,
    ))

def _has_meaningful_edit_instruction(user_input: str) -> bool:
    text = _safe_str(user_input)
    if not text:
        return False
    if _extract_edit_change_content(text):
        return True
    if _is_transform_request(text) or _is_shortening_request(text):
        return True
    if re.search(r"\b(?:add|remove|replace|title|subtitle|content|bullet|point|layout|chart|table|timeline|grid|stat|summary|conclusion|thank you|questions?)\b", text, re.IGNORECASE):
        return True
    return False

def _slide_exists(slide_num: Optional[int], slides: list) -> bool:
    try:
        idx = int(slide_num)
    except (TypeError, ValueError):
        return False
    return 1 <= idx <= len(slides or [])

def _extract_view_slide_request(user_input: str) -> Optional[dict]:
    text = user_input or ""
    if not re.search(r"\b(show|view|preview|display)\b", text, re.IGNORECASE):
        return None
    slide_num = _extract_explicit_slide_number_from_text(text)
    if slide_num is None:
        m = re.search(r"\bslide\s+#?(\d+)\b", text, re.IGNORECASE)
        if m:
            try:
                slide_num = int(m.group(1))
            except ValueError:
                slide_num = None
    if slide_num is None:
        return None
    ppt_ref = _extract_explicit_ppt_ref_from_text(text)
    return {"slide_number": slide_num, "ppt_ref": ppt_ref}

def _reasoning_state_snapshot() -> dict:
    conv = _conv_state()
    current_slide_id = (
        st.session_state.get("active_slide_index")
        or st.session_state.get("last_slide_index")
        or conv.get("active_slide_index")
    )
    return {
        "current_ppt_id": st.session_state.get("current_ppt_id") or conv.get("active_ppt_id"),
        "current_slide_id": current_slide_id,
        "last_intent": st.session_state.get("last_action_type") or conv.get("last_action_type"),
        "last_action": conv.get("last_action_text") or "",
    }

def _recent_message_digest(limit: int = 8) -> list[dict]:
    digest = []
    for msg in (st.session_state.get("messages", []) or [])[-limit:]:
        if not isinstance(msg, dict):
            continue
        content = _safe_str(msg.get("content", ""), "")
        content = re.sub(r"\s+", " ", content).strip()
        if len(content) > 450:
            content = content[:450].rstrip() + "..."
        digest.append({"role": msg.get("role", ""), "content": content})
    return digest

def build_assistant_state(slides: Optional[list] = None) -> dict:
    """
    Compact, LangGraph-ready conversation state.
    This is intentionally small: enough for contextual routing without sending
    the whole chat or full slide payloads to the model.
    """
    conv = _conv_state()
    current_id = st.session_state.get("current_ppt_id") or conv.get("active_ppt_id")
    item = get_ppt_by_id(current_id) if current_id else None
    outline = (item or {}).get("outline_payload") or st.session_state.get("outline_payload") or {}
    active_slides = slides if slides is not None else (outline.get("slides", []) if isinstance(outline, dict) else [])
    active_slide = (
        st.session_state.get("active_slide_index")
        or st.session_state.get("last_slide_index")
        or conv.get("active_slide_index")
    )
    try:
        active_slide = int(active_slide) if active_slide is not None else None
    except (TypeError, ValueError):
        active_slide = None

    slide_summary = summarize_slides_for_llm(active_slides) if active_slides else []
    active_slide_summary = None
    if active_slide and 1 <= active_slide <= len(slide_summary):
        active_slide_summary = slide_summary[active_slide - 1]

    return {
        "active_ppt_id": current_id,
        "active_topic": (item or {}).get("topic") or st.session_state.get("topic", "") or conv.get("last_topic", ""),
        "active_slide_index": active_slide,
        "active_slide": active_slide_summary,
        "slide_count": len(active_slides or []),
        "slides": slide_summary[:12],
        "last_action_type": st.session_state.get("last_action_type") or conv.get("last_action_type"),
        "last_action_text": conv.get("last_action_text", ""),
        "pending_intent": st.session_state.get("pending_intent"),
        "pending_action": st.session_state.get("pending_action"),
        "conversation_objects": _conversation_objects(),
        "recent_messages": _recent_message_digest(),
    }

def _has_explicit_deterministic_command(user_input: str) -> bool:
    return orchestrator_has_explicit_deterministic_command(user_input)

def _is_contextual_followup_candidate(user_input: str) -> bool:
    return orchestrator_is_contextual_followup_candidate(user_input)

def _fallback_contextual_intent(user_input: str, assistant_state: dict) -> Optional[dict]:
    return orchestrator_resolve_contextual_intent(user_input, assistant_state)

def resolve_contextual_intent(user_input: str, assistant_state: dict) -> Optional[dict]:
    """
    LLM contextual resolver for natural follow-ups. Returns a reasoned-style
    dict or None when the deterministic router should keep control.
    """
    return orchestrator_resolve_contextual_intent(
        user_input,
        assistant_state,
        llm_client=_client,
        deployment=AZURE_DEPLOYMENT,
    )

def resolve_open_conversation(user_input: str, assistant_state: dict, *, pending_need_topic: bool = False):
    """
    Broad LLM router for natural chat that is not an explicit deck command.
    It only returns safe pre-actions: ask, answer, or let deterministic code continue.
    """
    return orchestrator_resolve_open_conversation(
        user_input,
        assistant_state,
        llm_client=_client,
        deployment=AZURE_DEPLOYMENT,
        explicit_topic=_extract_create_ppt_topic(user_input),
        pending_need_topic=pending_need_topic,
    )

def _is_readonly_conversation_request(user_input: str) -> bool:
    """
    Read-only advice/help questions must not mutate the deck. This catches
    natural presentation coaching requests before they hit edit/transform routes.
    """
    return orchestrator_is_readonly_conversation_request(user_input)

def _readonly_target_slide(user_input: str, assistant_state: dict, slides: list) -> Optional[int]:
    slide_num = _extract_explicit_slide_number_from_text(user_input)
    if slide_num and _slide_exists(slide_num, slides):
        return slide_num
    active = assistant_state.get("active_slide_index")
    try:
        active = int(active) if active is not None else None
    except (TypeError, ValueError):
        active = None
    if active and _slide_exists(active, slides) and re.search(r"\b(?:this|it|that|current|shown|active)\b", user_input, re.IGNORECASE):
        return active
    return None

def _fallback_readonly_answer(user_input: str, assistant_state: dict, slides: list) -> str:
    slide_num = _readonly_target_slide(user_input, assistant_state, slides)
    if slide_num:
        slide = slides[slide_num - 1] if 1 <= slide_num <= len(slides) and isinstance(slides[slide_num - 1], dict) else {}
        title = _safe_str(slide.get("title", f"slide {slide_num}"), f"slide {slide_num}")
        return (
            f"Here are a few ways to improve slide {slide_num} ({title}):\n\n"
            "- Make the main takeaway explicit in the title or subtitle.\n"
            "- Reduce any long bullets to shorter, presentation-friendly phrases.\n"
            "- Add one concrete example, metric, or visual cue to make the point easier to remember.\n"
            "- Keep only the strongest 3-5 ideas so the slide is easier to present.\n\n"
            "I can apply one of these if you want."
        )
    topic = _safe_str(assistant_state.get("active_topic", ""), "this presentation")
    return (
        f"For {topic}, I would improve the deck by tightening the flow, reducing repeated points, "
        "adding concrete examples, and making each slide’s takeaway clearer. "
        "Tell me a slide number if you want specific feedback."
    )

def generate_conversational_answer(user_input: str, assistant_state: dict, slides: list) -> str:
    slide_num = _readonly_target_slide(user_input, assistant_state, slides)
    selected_slide = None
    if slide_num and 1 <= slide_num <= len(slides) and isinstance(slides[slide_num - 1], dict):
        selected_slide = slides[slide_num - 1]
    slide_context = _format_slide_view(
        selected_slide,
        slide_num,
        _ppt_display_label(assistant_state.get("active_ppt_id")),
    ) if selected_slide else ""
    prompt = f"""
You are a helpful PowerPoint coach inside a PPT assistant.

User asked:
{json.dumps(user_input)}

Compact state:
{json.dumps(assistant_state, ensure_ascii=True)}

Selected slide context:
{slide_context or "No single slide selected."}

Rules:
- Read-only answer only. Do not claim you edited, changed, regenerated, added, deleted, or updated anything.
- Give practical, specific advice grounded in the selected slide/deck context.
- If the user asks how to improve, give 3-5 concrete suggestions as a numbered list.
- End by asking whether they want you to apply one suggestion only when appropriate.
- Keep it concise and conversational.
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.35,
            max_tokens=500,
        )
        answer = _safe_str(resp.choices[0].message.content, "")
        if answer:
            return answer
    except Exception:
        pass
    return _fallback_readonly_answer(user_input, assistant_state, slides)

def reasoning_layer(user_input: str, state: dict) -> dict:
    """
    Enhance intent understanding using context.
    Must correct intent, not just pass through.
    """
    state = state if isinstance(state, dict) else {}
    text = (user_input or "").strip()
    lower = _norm_text(text)
    base = classify_ppt_intent(
        text,
        ppt_id=state.get("current_ppt_id"),
        slide_id=state.get("current_slide_id"),
    )
    result = {
        "intent": base.get("intent", "general_request"),
        "slide_id": base.get("slide_id"),
        "slide_ids": base.get("slide_ids", []),
        "ppt_id": base.get("ppt_id"),
        "operation": base.get("operation"),
        "content": base.get("content"),
        "action_type": base.get("action_type"),
        "target": base.get("target"),
        "sub_intent": base.get("sub_intent"),
        "missing_entities": base.get("missing_entities", []),
        "clarification_question": base.get("clarification_question"),
        "confidence": float(base.get("confidence", 0.35) or 0.35),
    }

    def _set(intent: str, operation: Optional[str] = None, slide_id: Optional[int] = None,
             ppt_id: Optional[str] = None, content: Optional[str] = None, confidence: float = 0.9):
        result["intent"] = intent
        result["operation"] = operation
        result["slide_id"] = slide_id if slide_id is not None else result.get("slide_id")
        result["ppt_id"] = ppt_id if ppt_id is not None else result.get("ppt_id")
        result["content"] = content
        result["confidence"] = confidence

    explicit_slide = _extract_explicit_slide_number_from_text(text)
    explicit_ppt = _extract_explicit_ppt_ref_from_text(text)
    current_slide = state.get("current_slide_id")
    last_intent = _safe_str(state.get("last_intent", ""))

    if _is_cancel_command(text):
        _set("cancel", operation="cancel", content=text, confidence=0.99)
        return result

    if _is_blank_slide_request(text):
        target = explicit_slide if explicit_slide is not None else current_slide
        _set("blank_slide", operation="clear", slide_id=target, content=text, confidence=0.97)
        return result

    if _is_clear_slide_request(text):
        target = explicit_slide if explicit_slide is not None else current_slide
        _set("clear_slide", operation="clear", slide_id=target, content=text, confidence=0.97)
        return result

    if explicit_ppt:
        result["ppt_id"] = explicit_ppt
        if explicit_slide is None and (
            result["intent"] in {"general_request", "unknown", "view_slide", "switch_ppt"}
            or re.search(r"\b(?:go to|open|switch|show|view)\b", text, re.IGNORECASE)
        ):
            _set(
                "switch_ppt",
                operation="switch",
                ppt_id=explicit_ppt,
                content=text or user_input,
                confidence=max(result["confidence"], 0.95),
            )

    if re.search(r"\b(?:sorry|no|actually|instead|correction|wait)\b", text, re.IGNORECASE):
        if explicit_slide is not None:
            result["slide_id"] = explicit_slide
            result["confidence"] = max(result["confidence"], 0.96)

    if re.search(
        r"\b(?:improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|condense|make it better|make it shorter|make it professional)\b",
        text,
        re.IGNORECASE,
    ) or re.search(r"\bfix this\b", lower):
        _set(
            "transform_content",
            operation="transform",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.98,
        )

    elif re.search(r"\b(?:regenerate|re-generate|redo|rebuild)\s+(?:only\s+|just\s+)?slide\s*#?\d+\b", text, re.IGNORECASE):
        _set(
            "regenerate_slide",
            operation="regenerate",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.98,
        )

    elif re.search(
        r"\b(?:speaker notes?|speaker script|slide script|script for each|short script|slide by slide"
        r"audience question|questions? audience|audience.*questions?|q\s*&?\s*a|"
        r"2[\s-]minute speech|two[\s-]minute|summarize.*speech|summary.*speech|"
        r"interactive question|question per slide|like i.m present|presenting|"
        r"explain\s+(?:this\s+)?(?:ppt|presentation|deck)|walk\s+me\s+through\s+(?:this\s+)?(?:ppt|presentation|deck))\b",
        text,
        re.IGNORECASE,
    ):
        result["intent"] = "presentation_mode"
        result["operation"] = "presentation_mode"
        result["content"] = text or user_input
        if not result.get("sub_intent"):
            result["sub_intent"] = "presentation_explain"
        result["confidence"] = 0.98

    elif _is_expand_existing_slide_request(text):
        _set(
            "add_points",
            operation="append",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.97,
        )

    elif re.search(r"\b(show|view|preview|display)\b.*\bslide\b|\bslide\b.*\b(show|view|preview|display)\b", text, re.IGNORECASE):
        _set(
            "view_slide",
            operation="view",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.95,
        )

    elif re.fullmatch(r"(add one more|one more|more)", lower) or re.search(r"\badd one more\b", lower):
        if last_intent in {"add_points", "edit_slide", "transform_content", "add_slide"}:
            _set(
                last_intent,
                operation="append" if last_intent == "add_points" else "modify",
                slide_id=explicit_slide if explicit_slide is not None else current_slide,
                ppt_id=explicit_ppt,
                content=text or user_input,
                confidence=0.9,
            )

    elif re.search(r"\badd\b.*\bmore\b", lower) and last_intent in {"add_points", "edit_slide", "transform_content"}:
        _set(
            last_intent,
            operation="append" if last_intent == "add_points" else "modify",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.88,
        )

    # Prefer interpreting explicit point/bullet additions as `add_points`.
    elif (
        not (_is_add_slide_request(text) and _extract_explicit_provided_content(text))
        and (
            re.search(r"\b(?:add|insert|append)\b.*\b(point|points|bullet|bullets|item|items)\b", text, re.IGNORECASE)
            or re.search(r"\bboth\b.*\b(side|sides|column|columns)\b", text, re.IGNORECASE)
        )
    ):
        _set(
            "add_points",
            operation="append",
            slide_id=explicit_slide if explicit_slide is not None else current_slide,
            ppt_id=explicit_ppt,
            content=text or user_input,
            confidence=0.97,
        )

    elif _is_add_slide_request(text):
        content = _extract_add_slide_content(text)
        _set(
            "add_slide",
            operation="insert",
            slide_id=None,
            ppt_id=explicit_ppt,
            content=content or text or user_input,
            confidence=0.95,
        )

    elif result["intent"] in {"general_request", "unknown"} and last_intent in {"add_points", "edit_slide", "transform_content", "update_slide"} and current_slide:
        if re.fullmatch(r"(this|it|that|here|continue|carry on)", lower):
            _set(
                last_intent,
                operation="append" if last_intent == "add_points" else "modify",
                slide_id=current_slide,
                ppt_id=explicit_ppt,
                content=text or user_input,
                confidence=0.84,
            )

    if result["slide_id"] is None and current_slide and result["intent"] in {"transform_content", "view_slide", "add_points", "update_slide", "edit_slide", "explain_slide"}:
        result["slide_id"] = current_slide

    if result["ppt_id"] is None:
        result["ppt_id"] = state.get("current_ppt_id")

    if result["intent"] in {"transform_content", "view_slide", "add_points", "update_slide", "explain_slide"} and result["slide_id"] is not None:
        result["confidence"] = max(result["confidence"], 0.9)

    return result

def _canonicalize_reasoned_prompt(user_input: str, reasoned: dict) -> str:
    text = (user_input or "").strip()
    intent = _safe_str((reasoned or {}).get("intent", ""))
    slide_id = (reasoned or {}).get("slide_id")
    ppt_id = _safe_str((reasoned or {}).get("ppt_id", ""))
    content = _safe_str((reasoned or {}).get("content", ""))

    if intent == "cancel":
        return "cancel"
    if intent == "blank_slide" and slide_id is not None:
        return f"blank slide {slide_id}"
    if intent == "clear_slide" and slide_id is not None:
        return f"clear slide {slide_id}"
    if intent == "transform_content" and slide_id is not None:
        return f"improve slide {slide_id}: {content or text}"
    if intent == "update_slide" and slide_id is not None:
        return f"edit slide {slide_id}: {content or text}"
    if intent == "view_slide" and slide_id is not None:
        if ppt_id:
            return f"show slide {slide_id} of {ppt_id}"
        return f"show slide {slide_id}"
    if intent == "explain_slide" and slide_id is not None:
        return f"explain slide {slide_id}"
    if intent == "add_points" and slide_id is not None:
        return f"{content or text or 'add point'} in slide {slide_id}"
    if intent == "edit_slide" and slide_id is not None and content:
        return f"edit slide {slide_id}: {content}"
    if intent == "delete_slide" and slide_id is not None:
        return f"delete slide {slide_id}"
    return text

def _store_pending_action(state: Optional[dict]):
    st.session_state.pending_intent = copy.deepcopy(state) if state else None
    st.session_state.pending_action = copy.deepcopy(state) if state else None

def _clear_pending_turn_state():
    st.session_state.pending_intent = None
    st.session_state.pending_action = None
    if "agent" in st.session_state:
        st.session_state.agent.clear_state()

def _is_unrelated_chat_message(user_input: str) -> bool:
    text = _safe_str(user_input).lower().strip()
    if not text:
        return False
    return bool(re.fullmatch(
        r"(hi|hello|hey|thanks|thank you|ok|okay|cool|great|good|nice|bye|goodbye)",
        text,
    ))

def _is_pure_chat_message(user_input: str) -> bool:
    return _is_unrelated_chat_message(user_input)

def _remember_ppt_history_switch(prev_ppt_id: Optional[str], new_ppt_id: Optional[str]) -> None:
    if not prev_ppt_id or not new_ppt_id or prev_ppt_id == new_ppt_id:
        return
    recent = list(st.session_state.get("recent_ppt_ids", []) or [])
    recent = [ppt_id for ppt_id in recent if ppt_id and ppt_id != prev_ppt_id and ppt_id != new_ppt_id]
    recent.insert(0, prev_ppt_id)
    st.session_state["recent_ppt_ids"] = recent[:10]
    st.session_state["previous_ppt_id"] = prev_ppt_id

def _get_previous_ppt_id() -> Optional[str]:
    current_id = st.session_state.get("current_ppt_id")
    for ppt_id in st.session_state.get("recent_ppt_ids", []) or []:
        if ppt_id and ppt_id != current_id and get_ppt_by_id(ppt_id):
            return ppt_id
    previous = st.session_state.get("previous_ppt_id") or st.session_state.get("last_ppt")
    if previous and previous != current_id and get_ppt_by_id(previous):
        return previous
    history = st.session_state.get("ppt_history", []) or []
    for item in reversed(history):
        if not isinstance(item, dict):
            continue
        ppt_id = item.get("id")
        if ppt_id and ppt_id != current_id and get_ppt_by_id(ppt_id):
            return ppt_id
    return None

def _is_previous_ppt_request(user_input: str) -> bool:
    text = _safe_str(user_input)
    return bool(re.search(r"\b(?:go back|previous ppt|last presentation)\b", text, re.IGNORECASE))

def _extract_ppt_info_slots(user_input: str, reasoned: dict) -> dict:
    text = user_input or ""
    slots = {}
    ppt_ref = (reasoned or {}).get("ppt_id") or _extract_explicit_ppt_ref_from_text(text)
    if ppt_ref:
        slots["ppt_ref"] = ppt_ref
    if re.search(r"\b(count|how many|number of)\b", text, re.IGNORECASE):
        slots["info_type"] = "count"
    elif re.search(r"\b(list|which|show all|all ppt|all ppts|presentations)\b", text, re.IGNORECASE):
        slots["info_type"] = "list"
    elif re.search(r"\b(topic|title|name)\b", text, re.IGNORECASE):
        slots["info_type"] = "topic"
    return slots

def _extract_suggest_topic_slots(user_input: str) -> dict:
    text = user_input or ""
    scope = "new"
    if re.search(r"\b(current|this|existing|active)\b", text, re.IGNORECASE):
        scope = "current"
    return {"scope": scope, "topic": _extract_topic_constraint_from_text(text)}

def _topic_keywords(text: str) -> set[str]:
    norm = _norm_text(text)
    if not norm:
        return set()
    stopwords = {
        "the", "a", "an", "and", "or", "of", "for", "to", "in", "on", "with",
        "about", "around", "regarding", "topic", "topics", "presentation",
        "ppt", "deck", "ideas", "idea", "suggest", "brainstorm", "new",
        "current", "this", "existing", "active",
    }
    return {word for word in norm.split() if word and word not in stopwords}

def _get_active_topic_domain() -> str:
    explicit = _safe_str(st.session_state.get("active_topic_domain", ""))
    if explicit:
        return explicit
    current_topic = _safe_str(st.session_state.get("topic", ""))
    if current_topic:
        return current_topic
    return _safe_str(_conv_state().get("last_topic", ""))

def _set_active_topic_domain(topic: str):
    cleaned = _safe_str(topic).strip()
    if cleaned:
        st.session_state["active_topic_domain"] = cleaned
        _set_conv_state(last_topic=cleaned)

def _is_topic_fragment_request(user_input: str) -> bool:
    text = _safe_str(user_input)
    if not text:
        return False
    if _is_topic_ideas_request(text):
        return False
    if re.search(r"\b(?:create|make|generate|build)\b.*\b(?:ppt|presentation|deck)\b", text, re.IGNORECASE):
        return False
    return bool(re.fullmatch(r"\s*(?:about|on|of|for|around)\s+.{2,80}\s*", text, re.IGNORECASE))

def _is_ambiguous_topic_request(user_input: str, active_domain: str = "") -> bool:
    text = _safe_str(user_input)
    topic_hint = _extract_topic_constraint_from_text(text)
    if not topic_hint:
        return False
    if _is_topic_ideas_request(text):
        return False
    if active_domain:
        return False
    return len(_topic_keywords(topic_hint)) <= 3

def _topic_domains_conflict(new_topic: str, active_domain: str) -> bool:
    new_keywords = _topic_keywords(new_topic)
    active_keywords = _topic_keywords(active_domain)
    if not new_keywords or not active_keywords:
        return False
    return not bool(new_keywords & active_keywords)

def _merge_topic_refinement(active_domain: str, fragment: str) -> str:
    active = _safe_str(active_domain).strip()
    frag = _safe_str(fragment).strip()
    if not active:
        return frag
    if not frag:
        return active
    if _topic_keywords(active) & _topic_keywords(frag):
        return frag
    return f"{active}: {frag}"

def _extract_create_ppt_slots(user_input: str) -> dict:
    return {
        "topic": _extract_create_ppt_topic(user_input),
        "slide_count": _extract_slide_count_from_text(user_input),
        "sections": _extract_create_sections_from_prompt(user_input),
    }

def _extract_topic_constraint_from_text(user_input: str) -> str:
    text = _safe_str(user_input)
    if not text:
        return ""
    patterns = [
        r"(?i)\b(?:about|on|of|for|around)\s+([a-z0-9][a-z0-9\s&/-]{1,60})",
        r"(?i)\brelated\s+to\s+([a-z0-9][a-z0-9\s&/-]{1,60})",
        r"(?i)\bsuggest(?:\s+trending)?\s+topics?\s+(?:about|on|of|for)\s+([a-z0-9][a-z0-9\s&/-]{1,60})",
        r"(?i)\btrending\s+topics?\s+(?:about|on|of|for)\s+([a-z0-9][a-z0-9\s&/-]{1,60})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            topic = re.sub(r"\b(?:for|ppt|presentation|deck)\b.*$", "", m.group(1), flags=re.IGNORECASE).strip(" ,.;:-")
            topic = re.sub(r"(?i)\b(?:generation|ideas?)\b.*$", "", topic).strip(" ,.;:-")
            topic = re.sub(r"(?i)\bto\s+(?:create|generate|make|build)(?:\s+(?:a\s+)?(?:ppt|presentation|deck))?\b.*$", "", topic).strip(" ,.;:-")
            topic = re.sub(r"(?i)\bslide\s*#?\d+\b", "", topic).strip(" ,.;:-")
            topic = re.sub(r"(?i)\b(?:left|right|both)(?:\s+the)?\s+(?:side|sides|column|columns)\b", "", topic).strip(" ,.;:-")
            topic = re.sub(r"(?i)\bin\s*$", "", topic).strip(" ,.;:-")
            topic = re.sub(r"\s{2,}", " ", topic).strip()
            if re.fullmatch(r"(?i)(?:left|right|both|side|sides|column|columns)", topic):
                return ""
            if topic:
                return topic
    if re.search(r"\bAI\b", text, re.IGNORECASE):
        return "AI"
    return ""

def _point_matches_topic_hint(point: str, topic_hint: str) -> bool:
    point_norm = _norm_text(point)
    topic_norm = _norm_text(topic_hint)
    if not point_norm or not topic_norm:
        return True
    if topic_norm in point_norm:
        return True
    stopwords = {
        "the", "a", "an", "and", "or", "of", "for", "to", "in", "on", "with",
        "by", "about", "around", "into", "from", "daily", "life",
    }
    topic_words = {w for w in topic_norm.split() if w and w not in stopwords}
    point_words = {w for w in point_norm.split() if w and w not in stopwords}
    if not topic_words:
        return True
    return bool(topic_words & point_words)

def _looks_like_change_instruction(user_input: str) -> bool:
    text = _norm_text(user_input)
    if not text:
        return False
    if re.search(r"\b(?:edit|change|update|modify|revise)\s+slide\s+\d+\b", text, re.IGNORECASE):
        if not re.search(
            r"\b(?:add|insert|remove|delete|replace|title|subtitle|content|text|bullet|point|"
            r"layout|chart|table|timeline|grid|stat|company|result|metric|section|summary|conclusion|"
            r"thank you|questions?)\b",
            text,
            re.IGNORECASE,
        ):
            return False
    return bool(re.search(
        r"\b(?:add|insert|remove|delete|replace|change|update|modify|revise|rewrite|"
        r"shorten|expand|improve|fix|make|append|trim|reorder)\b",
        text,
        re.IGNORECASE,
    ))

def _format_slide_view(slide: dict, slide_number: int, ppt_label: str) -> str:
    if not isinstance(slide, dict):
        return f"Slide {slide_number} is unavailable in {ppt_label}."
    title = _safe_str(slide.get("title", f"Slide {slide_number}"))
    subtitle = _safe_str(slide.get("subtitle", ""))
    layout = _safe_str(slide.get("layout", "bullets"), "bullets")
    lines = [f"Slide {slide_number} — {title}"]
    if subtitle:
        lines.append(subtitle)
    lines.append(f"Layout: {layout}")

    if layout == "two_column":
        left_title = _safe_str(slide.get("left_title", "Left"))
        right_title = _safe_str(slide.get("right_title", "Right"))
        lines.append("")
        lines.append(f"{left_title}:")
        for item in flatten_slide_content(slide.get("left_points", [])):
            lines.append(f"- {item}")
        lines.append("")
        lines.append(f"{right_title}:")
        for item in flatten_slide_content(slide.get("right_points", [])):
            lines.append(f"- {item}")
        return "\n".join(lines)

    if layout == "timeline":
        lines.append("")
        for step in slide.get("steps", []) or []:
            if isinstance(step, dict):
                label = _safe_str(step.get("label", ""))
                detail = _safe_str(step.get("detail", ""))
                if label or detail:
                    lines.append(f"- {label}: {detail}".strip(": "))
        return "\n".join(lines)

    if layout == "icon_grid":
        lines.append("")
        for item in slide.get("grid_items", []) or []:
            if isinstance(item, dict):
                title_clean = _safe_str(item.get("title", ""))
                detail_clean = _safe_str(item.get("detail", ""))
                if title_clean or detail_clean:
                    lines.append(f"- {title_clean}: {detail_clean}".strip(": "))
        return "\n".join(lines)

    if layout == "case_study":
        company = _safe_str(slide.get("company", ""))
        result = _safe_str(slide.get("result", ""))
        if company:
            lines.append("")
            lines.append(f"Company: {company}")
        if result:
            lines.append(f"Result: {result}")

    if layout in ("big_stat", "hybrid_insight"):
        stat = _safe_str(slide.get("stat", ""))
        stat_label = _safe_str(slide.get("stat_label", ""))
        if stat:
            lines.append("")
            lines.append(f"Stat: {stat}")
        if stat_label:
            lines.append(f"Label: {stat_label}")

    content = flatten_slide_content(slide.get("content", []))
    if content:
        lines.append("")
        for item in content:
            lines.append(f"- {item}")
    return "\n".join(lines)

def _remember_action_context(action_type: str, slide_index: Optional[int] = None, ppt_id: Optional[str] = None, action_text: Optional[str] = None):
    if action_type:
        st.session_state.last_action_type = action_type
    if slide_index is not None:
        st.session_state.last_slide_index = slide_index
        st.session_state.active_slide_index = slide_index
    if ppt_id:
        st.session_state.last_ppt_id = ppt_id
    _set_conv_state(
        last_action_type=action_type or st.session_state.get("last_action_type"),
        active_slide_index=slide_index if slide_index is not None else st.session_state.get("active_slide_index"),
        active_ppt_id=ppt_id or st.session_state.get("last_ppt_id") or st.session_state.get("current_ppt_id"),
        last_action_text=action_text or _conv_state().get("last_action_text", ""),
    )

def _latest_slide_context() -> Optional[dict]:
    slide_index = st.session_state.get("last_slide_index")
    ppt_id = st.session_state.get("last_ppt_id") or st.session_state.get("current_ppt_id")
    if not slide_index or not ppt_id:
        return None
    return {
        "slide_index": slide_index,
        "ppt_id": ppt_id,
        "action_type": st.session_state.get("last_action_type"),
    }

def _should_start_new_ppt(user_input: str) -> Optional[str]:
    topic = _extract_create_ppt_topic(user_input)
    if not topic:
        return None
    current_topic = st.session_state.get("topic", "")
    if not st.session_state.get("current_ppt_id"):
        return topic
    if not _same_topic(topic, current_topic):
        return topic
    return None

def _resolve_vague_followup_prompt(prompt: str) -> str:
    text = prompt or ""
    if not _is_ambiguous_add_followup(text):
        return prompt
    ctx = _latest_slide_context()
    if not ctx or not ctx.get("slide_index"):
        return prompt
    if ctx.get("ppt_id") and ctx.get("ppt_id") != st.session_state.get("current_ppt_id"):
        switch_active_ppt(ctx["ppt_id"])
    requested = _parse_bullet_add_request(text)
    if requested is None:
        requested = 1
    label = "point" if int(requested) == 1 else "points"
    return f"add {int(requested)} {label} in slide {ctx['slide_index']}"

def _slot_missing(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip()) or (isinstance(value, (list, tuple, dict)) and not value)

def _agent_required_slots(intent: Optional[str]) -> list[str]:
    return {
        "create_ppt": ["topic", "slide_count"],
        "edit_slide": ["slide_number", "change_content"],
        "transform_content": ["slide_number", "change_content"],
        "update_slide": ["slide_number", "change_content"],
        "add_points": ["slide_number", "change_content"],
        "clear_slide": ["slide_number"],
        "blank_slide": ["slide_number"],
        "add_slide": ["slide_content", "position"],
        "move_slide": ["slide_number", "anchor_slide"],
        "swap_slides": ["slide_a", "slide_b"],
        "merge_slides": ["slide_a", "slide_b"],
        "delete_slide": ["slide_number"],
        "regenerate_slide": ["slide_number"],
        "presentation_mode": [],
        "explain_slide": ["slide_number"],
        "download_ppt": [],
        "preview_ppt": [],
        "view_slide": ["slide_number"],
        "ppt_info": ["info_type"],
        "suggest_topic": [],
        "refine_ppt": [],
        "cancel": [],
        "greeting": [],
        "smalltalk": [],
        "switch_ppt": [],
        "unknown": [],
    }.get(intent or "", [])

def _detect_explicit_view_slide_request(user_input: str) -> Optional[dict]:
    """
    EARLY DETECTION: Check for explicit "show slide 5" patterns BEFORE LLM processing.
    Returns None if not a view_slide request, else returns dict ready for execution.
    
    This bypasses LLM ambiguity when user explicitly requests slide retrieval.
    Called FIRST to ensure previous conversation context doesn't interfere.
    """
    text = (user_input or "").strip()
    if not text:
        return None
    
    # Explicit "show slide X" pattern — highest priority
    ppt_ref = _extract_explicit_ppt_ref_from_text(text)
    patterns = [
        r"\b(?:show|view|display|see)\s+slide\s+#?(\d+)\b",
        r"\bslide\s+#?(\d+)\s+(?:show|view|display|see)\b",
        r"\bgo\s+to\s+slide\s+#?(\d+)\b",
        r"\b(?:jump|skip|navigate)\s+to\s+slide\s+#?(\d+)\b",
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            slide_num = int(match.group(1))
            return {
                "intent": "view_slide",
                "slots": {"slide_number": slide_num, "ppt_ref": ppt_ref},
                "missing_slots": [],
                "next_question": None,
                "action": "execute",
                "confidence": 0.98,
            }
    return None


def _infer_agent_slots(intent: Optional[str], user_input: str, slots: dict, context: dict) -> dict:
    text = user_input or ""
    updated = dict(slots or {})
    slides = context.get("slides", []) if isinstance(context, dict) else []

    if intent in {"edit_slide", "transform_content", "update_slide", "view_slide", "delete_slide", "add_points", "explain_slide"}:
        slide_idx = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
        if slide_idx is not None and updated.get("slide_number") is None:
            updated["slide_number"] = slide_idx

    if intent == "regenerate_slide":
        slide_idx = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
        if slide_idx is not None and updated.get("slide_number") is None:
            updated["slide_number"] = slide_idx

    if intent in {"edit_slide", "transform_content", "update_slide", "add_points"} and updated.get("change_content") is None:
        change = _extract_edit_change_content(text) or (_resolve_vague_followup_prompt(text) if _is_ambiguous_add_followup(text) else None)
        if change:
            updated["change_content"] = change
    elif intent in {"edit_slide", "transform_content", "update_slide", "add_points"} and updated.get("change_content") is not None:
        if _is_generic_edit_request(text) and not _has_meaningful_edit_instruction(text):
            updated.pop("change_content", None)

    if intent == "add_slide":
        explicit_content = _extract_explicit_provided_content(text)
        content = _extract_add_slide_content(text)
        position = _extract_add_slide_position(text, slides)
        if explicit_content and updated.get("slide_content") is None:
            updated["slide_content"] = explicit_content
            updated["explicit_user_content"] = explicit_content
        elif content and updated.get("slide_content") is None:
            updated["slide_content"] = content
        if position is not None and updated.get("position") is None:
            updated["position"] = position

    if intent == "move_slide":
        if updated.get("slide_number") is None:
            slide_idx = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
            if slide_idx is not None:
                updated["slide_number"] = slide_idx
        if updated.get("position") is None:
            pos = _extract_slide_count_from_text(text)
            if pos is not None:
                updated["position"] = pos

    if intent in {"swap_slides", "merge_slides"}:
        slide_nums = re.findall(r"\bslide\s*#?(\d+)\b|\b(\d+)(?:st|nd|rd|th)?\s+slide\b|\b(\d+)\b", text, re.IGNORECASE)
        parsed_nums = []
        for groups in slide_nums:
            raw = next((g for g in groups if g), None)
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            if value not in parsed_nums:
                parsed_nums.append(value)
        if intent == "merge_slides" and len(parsed_nums) >= 2:
            updated.setdefault("slide_a", parsed_nums[0])
            updated.setdefault("slide_b", parsed_nums[1])
            updated.setdefault("slide_ids", parsed_nums[:2])
        if updated.get("slide_a") is None:
            slide_idx = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
            if slide_idx is not None:
                updated["slide_a"] = slide_idx
        elif updated.get("slide_b") is None:
            slide_idx = _extract_explicit_slide_number_from_text(text) or _resolve_slide_reference_text(text, slides)
            if slide_idx is not None:
                updated["slide_b"] = slide_idx

    if intent == "create_ppt":
        topic = _extract_create_ppt_topic(text)
        slide_count = _extract_slide_count_from_text(text)
        if topic and updated.get("topic") is None:
            updated["topic"] = topic
        if slide_count and updated.get("slide_count") is None:
            updated["slide_count"] = slide_count

    if intent in {"ppt_info", "download_ppt", "preview_ppt"} and updated.get("ppt_ref") is None:
        ppt_ref = _extract_explicit_ppt_ref_from_text(text)
        if ppt_ref:
            updated["ppt_ref"] = ppt_ref

    return updated

# ------------------------------------------------------------------------------
#  SLOT-FILLING CONVERSATIONAL AGENT
# ------------------------------------------------------------------------------
class ConversationalAgent:
    INTENT_SLOTS = {
        "create_ppt": ["topic", "slide_count", "sections"],
        "edit_slide": ["slide_number", "change_content"],
        "transform_content": ["slide_number", "change_content"],
        "update_slide": ["slide_number", "change_content"],
        "add_points": ["slide_number", "change_content"],
        "clear_slide": ["slide_number"],
        "blank_slide": ["slide_number"],
        "add_slide": ["slide_content", "position"],
        "move_slide": ["slide_number", "anchor_slide", "relation", "position"],
        "swap_slides": ["slide_a", "slide_b"],
        "merge_slides": ["slide_a", "slide_b"],
        "delete_slide": ["slide_number"],
        "regenerate_slide": ["slide_number"],
        "presentation_mode": [],
        "explain_slide": ["slide_number"],
        "download_ppt": ["ppt_ref"],
        "preview_ppt": ["ppt_ref"],
        "view_slide": ["ppt_ref", "slide_number"],
        "ppt_info": ["info_type", "ppt_ref"],
        "suggest_topic": ["scope"],
        "refine_ppt": ["tone", "style_hint"],
        "cancel": [],
        "greeting": [],
        "smalltalk": [],
        "unknown": [],
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
        
        # ── EARLY CHECK: Explicit view_slide patterns bypass LLM ─────────────────
        explicit_view = _detect_explicit_view_slide_request(user_input)
        if explicit_view:
            # Clear previous state and execute view_slide immediately
            self.clear_state()
            return None, {
                "intent": "view_slide",
                "slots": explicit_view["slots"],
            }, True
        
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
            
            # ── STRUCTURAL INTENT CHANGE: Force reset if switching categories ─────────
            _STRUCTURAL_INTENTS = {
                "view_slide", "explain_slide", "preview_ppt", "download_ppt", "switch_ppt",
                "ppt_info", "suggest_topic", "create_ppt", "cancel", "greeting",
                "unknown",
            }
            _EDIT_INTENTS = {
                "edit_slide", "transform_content", "update_slide", "add_points", "add_slide", "delete_slide",
                "move_slide", "swap_slides", "merge_slides", "clear_slide", "blank_slide",
            }
            
            prev_category = "structural" if self.state["intent"] in _STRUCTURAL_INTENTS else "edit" if self.state["intent"] in _EDIT_INTENTS else "other"
            new_category = "structural" if new_intent in _STRUCTURAL_INTENTS else "edit" if new_intent in _EDIT_INTENTS else "other"
            
            # If switching between structural and edit intents, clear state
            if prev_category != new_category and self.state["intent"] is not None:
                merged_slots = new_slots
                self.state["intent"] = new_intent
            elif self.state["intent"] == new_intent:
                # Same intent: merge slots (conversation continuation)
                cleaned_new_slots = {k: v for k, v in new_slots.items() if not _slot_missing(v)}
                merged_slots = {**self.state["slots"], **cleaned_new_slots}
            else:
                # Different intent, same category: start fresh for new intent
                merged_slots = new_slots
                self.state["intent"] = new_intent
            merged_slots = _infer_agent_slots(new_intent, user_input, merged_slots, context)
            required_slots = _agent_required_slots(new_intent)
            computed_missing = [slot for slot in required_slots if _slot_missing(merged_slots.get(slot))]
            if new_intent in {"add_slide", "move_slide", "swap_slides", "merge_slides", "create_ppt", "edit_slide", "transform_content", "update_slide", "add_points", "delete_slide", "view_slide", "explain_slide"}:
                missing = computed_missing
            if new_intent in {"edit_slide", "transform_content", "update_slide", "add_points"} and not _has_meaningful_edit_instruction(user_input):
                merged_slots.pop("change_content", None)
                if "change_content" not in missing:
                    missing = list(missing) + ["change_content"]
            self.state["slots"] = merged_slots
            self.state["missing_slots"] = missing
            next_q = _polish_next_question(user_input, new_intent, merged_slots, missing, next_q)
            self.state["next_question"] = next_q
            self.state["action"] = action
            if missing:
                action = "ask"
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
        active_edit_context = st.session_state.get("active_edit_context") or {}
        conv = _conv_state()
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
- Active edit context: {json.dumps(active_edit_context)}
- Conversation memory: {json.dumps(conv)}

Your task: Update the state based on the user's input. Return a JSON object with:

{{
            "intent": "one of: create_ppt, edit_slide, update_slide, transform_content, add_points, explain_slide, clear_slide, blank_slide, view_slide, add_slide, move_slide, swap_slides, merge_slides, delete_slide, regenerate_slide, presentation_mode, download_ppt, preview_ppt, ppt_info, suggest_topic, refine_ppt, cancel, greeting, smalltalk, unknown",
  "slots": {{ ... }},
  "missing_slots": [...],
  "next_question": "string or null",
  "action": "ask" or "execute"
}}

Required slots for each intent:
- create_ppt: topic (string), slide_count (integer), sections (string, optional)
- edit_slide: slide_number (integer), change_content (string)
- update_slide: slide_number (integer), change_content (string)
- transform_content: slide_number (integer), change_content (string)
- add_points: slide_number (integer), change_content (string)
- explain_slide: slide_number (integer)
- view_slide: ppt_ref (string or integer, optional), slide_number (integer)
- add_slide: slide_content (string), position (string or integer)
- move_slide: slide_number (integer), anchor_slide (integer or string), relation (before/after/position), position (integer, optional)
- swap_slides: slide_a (integer or string), slide_b (integer or string)
- merge_slides: slide_a (integer or string), slide_b (integer or string)
- delete_slide: slide_number (integer)
- regenerate_slide: slide_number (integer)
- presentation_mode: no required slots; use this for speaker notes, scripts, audience questions, speech summaries, and presenting explanations
- download_ppt: ppt_ref (string or integer, optional)
- preview_ppt: ppt_ref (string or integer, optional)
- ppt_info: info_type (string: "count", "list", "topic"), ppt_ref (optional)
- suggest_topic: scope (string: "current", "new")
- refine_ppt: tone (string, optional), style_hint (string, optional)
- greeting / smalltalk / unknown: no slots

Rules:
- If the user's input indicates a new intent, reset slots and set intent accordingly.
- If the user is continuing a previous intent, merge new information.
- Intent priority: create_ppt > switch_ppt/reference lookup > move_slide/swap_slides > edit_slide/add_slide > follow-up clarification.
- CRITICAL: If the user asks to show, view, preview, display, see, go to, jump to, or navigate to a SPECIFIC SLIDE NUMBER, classify it as view_slide. This ALWAYS overrides previous conversation context.
- CRITICAL: Explicit slide numbers (e.g., "show slide 5") signal a NEW independent request. Do NOT merge with previous intent slots or context.
- If the user says "show slide 4 in ppt 3", set intent=view_slide, slide_number=4, ppt_ref="ppt 3", missing_slots=[] and action="execute". Never use preview_ppt for a specific slide.
- If the user says "explain slide 4", classify it as explain_slide. Never classify this as add_slide.
- If the user says "remove slide 3" or "delete slide 3", classify it as delete_slide. Never classify this as add_slide.
- If the user says "merge slide 4 and 5", classify it as merge_slides with slide_a=4 and slide_b=5. Never classify this as add_slide.
- If the user says "regenerate slide 3" or "regenerate only slide 3", classify it as regenerate_slide. Never classify this as add_slide.
- If the user asks for scripts, speaker notes, audience questions, a 2-minute speech, or how to present the PPT, classify it as presentation_mode. Never classify this as add_slide.
- If the user says "make slide 5 longer", "expand slide 5", or "add more detail to slide 5", classify it as update_slide with slide_number=5 and change_content set to the user's request. Never classify this as add_slide.
- If the intent is unclear, classify it as unknown and ask a clarification. Never default to add_slide.
- If the user asks to improve, shorten, simplify, rewrite, refine, or fix slide content, classify it as transform_content.
- If the user says "remove everything", "clear slide", or "blank slide", classify it as clear_slide or blank_slide.
- If the user says "stop" or "cancel", classify it as cancel.
- If the user says "make/create/generate/build ppt/presentation/deck on <topic>", treat it as create_ppt even if extra content is included in the same message.
- If the user says "add one more point", "add another bullet", or similar without a slide number, reuse the last edited slide from conversation memory.
- If the user says "add slide" or "add a slide" without a slide number, classify it as add_slide and ask for the slide content or placement instead of reusing the last edited slide.
- If the user says "move slide 5 before slide 4", classify it as move_slide.
- If the user says "move conclusion slide after slide 4", classify it as move_slide and use the slide title as the source slide reference.
- If the user says "swap slide 2 and slide 4" or "change position of both", classify it as swap_slides.
- Only mark a slot as missing if it's required and not yet filled.
- For add_slide, position is required. If missing, set missing_slots = ["position"].
- For edit_slide, if user says "edit slide 3" without change content, set slide_number=3, change_content=null, missing_slots=["change_content"].
- For edit_slide, if user says "add 2 points about AI to slide 3", set slide_number=3, change_content="add 2 points about AI", missing_slots=[].
- If the user says "add one point" or "add a point" to a slide, set change_content="add 1 point" and let the editor generate the bullet automatically.
- If the user says "edit slide" or "change slide" without an explicit slide number, do NOT reuse the previous slide. Set missing_slots=["slide_number"] and ask which slide they want to edit.
- If active edit context is deck-level and the user does not explicitly mention a slide number, ask about the PPT deck, not a slide.
- Only use slide-level wording when the user explicitly says "slide N".
- For ppt_info:
  - "how many ppts" => info_type="count"
  - "which ppts" or "list ppts" => info_type="list"
  - "topic of ppt 4" or "what is ppt 4" => info_type="topic", ppt_ref="ppt 4"
  - If the user names a deck number, prefer ppt_ref over asking a clarification.
- If the user asks for the topic/name/title of an existing or generated PPT, use ppt_info, not suggest_topic.
- Use suggest_topic only when the user explicitly asks for new topic ideas, brainstorming, or suggestions.
- Use refine_ppt for requests like "make it shorter", "make it cleaner", "make it more professional", "change tone to casual", or "add important points only" when the user is refining an existing deck.
- If the user says "go to slide 3" or "open slide 3", use the slide number as context and do not treat it like a PPT lookup.
- For add_slide, if the user specifies a slide type (e.g., "add a summary slide" or "add a conclusion slide"), set slide_content to that type (e.g., "summary"). This satisfies the content requirement.
- If the user provides a position (e.g., "at the end", "after slide 5") while an add_slide intent is pending, preserve the existing slide_content.
- For add_slide, if user says "add a summary slide at the end", set slide_content="summary", position="end", missing_slots=[].
- If the user says "add slide" or "add a slide" without a slide number, classify it as add_slide, not edit_slide, and ask for the new slide content/position instead of reusing the last edited slide.
- For move_slide, "before slide N" means place the slide immediately before N and "after slide N" means place it immediately after N.
- For move_slide, if the user names only the source slide but not the destination, ask for the destination.
- For swap_slides, ask for the second slide if only one slide is named.
- For create_ppt, if user says "make a ppt on AI", set topic="AI", slide_count=null, sections=null, missing_slots=["slide_count"].
- For smalltalk like "thanks", "ok", "yes", "no", set intent="smalltalk".
- For "hi", "hello" set intent="greeting".

Be conversational: next_question should ask ONLY for the next missing slot (one at a time).

Return ONLY valid JSON. No extra text.
"""
        return prompt

def _polish_next_question(user_input: str, intent: Optional[str], slots: dict, missing: list, next_question: Optional[str]) -> str:
    """
    Make slot-filling follow-ups more specific and PPT-aware.
    This keeps the assistant from asking generic wording like "What point?"
    """
    user_text = (user_input or "").lower()
    q = (next_question or "").strip()
    slide_num = slots.get("slide_number")
    ppt_id = st.session_state.get("current_ppt_id") or slots.get("ppt_ref")
    ppt_label = _ppt_display_label(str(ppt_id)) if ppt_id else "the current presentation"

    if intent == "cancel":
        return "Okay, I cancelled that request."

    if intent == "create_ppt":
        if "topic" in (missing or []) or not slots.get("topic"):
            return _create_topic_followup(user_input)
        if "slide_count" in (missing or []) or not slots.get("slide_count"):
            topic = _safe_str(slots.get("topic", ""), "the presentation")
            return f"Great — how many slides should the presentation on {topic} have?"

    if intent in {"clear_slide", "blank_slide"}:
        if "slide_number" in (missing or []) or slide_num is None:
            return f"Which slide should I clear in {ppt_label}?"
        if intent == "clear_slide":
            return f"This will remove everything from slide {slide_num}. Do you want me to continue?"
        return f"Should I make slide {slide_num} blank in {ppt_label}?"

    if intent == "edit_slide":
        if "slide_number" in (missing or []) or slide_num is None:
            return f"Which slide would you like to edit in {ppt_label}?"
        if "change_content" in (missing or []):
            return f"What change would you like to make to slide {slide_num} in {ppt_label}?"

    if intent == "update_slide":
        if "slide_number" in (missing or []) or slide_num is None:
            return f"Which slide should I update in {ppt_label}?"
        if "change_content" in (missing or []):
            return f"What should I change on slide {slide_num} in {ppt_label}?"

    if intent == "explain_slide":
        if "slide_number" in (missing or []) or slide_num is None:
            return f"Which slide should I explain in {ppt_label}?"

    if intent == "unknown":
        return q or "What would you like me to do with the presentation?"

    if intent == "add_slide" and "slide_content" in (missing or []):
        # If slide content is missing, ask for the content. Position is a separate slot.
        return f"What content should the new slide cover in {ppt_label}?"
    if intent == "add_slide" and "position" in (missing or []):
        return f"Where should I place the new slide in {ppt_label}?"

    if intent == "add_points":
        raw_requested_points = _parse_bullet_add_request(user_input) or _parse_bullet_add_request(str(slots.get("change_content", "")))
        if "slide_number" in (missing or []) or slide_num is None:
            if raw_requested_points and raw_requested_points > _MAX_POINTS_PER_REQUEST:
                return f"I can add up to {_MAX_POINTS_PER_REQUEST} points total at a time. Which slide should I add them to in {ppt_label}?"
            return f"Which slide should I add points to in {ppt_label}?"
        if "change_content" in (missing or []):
            return f"What should I add or expand on slide {slide_num} in {ppt_label}?"

    if intent == "move_slide":
        if "slide_number" in (missing or []) or slots.get("slide_number") is None:
            return f"Which slide would you like to move in {ppt_label}?"
        if "anchor_slide" in (missing or []) and "position" not in (missing or []):
            return f"Where should slide {slots.get('slide_number')} go in {ppt_label}?"
        if "position" in (missing or []):
            return f"What position should slide {slots.get('slide_number')} move to in {ppt_label}?"

    if intent == "swap_slides":
        if "slide_a" in (missing or []) or slots.get("slide_a") is None:
            return f"Which first slide should I swap in {ppt_label}?"
        if "slide_b" in (missing or []) or slots.get("slide_b") is None:
            return f"Which second slide should I swap with slide {slots.get('slide_a')} in {ppt_label}?"

    if q:
        q = re.sub(r"(?i)\bwhat point\b", "What bullet point", q)
        q = re.sub(r"(?i)\bwhat points\b", "What bullet points", q)
        return q

    return "Could you please provide more information?"

def _point_limit_confirmation_question(slide_num: int, requested: int, allowed: int, ppt_id: Optional[str] = None) -> str:
    ppt_label = _ppt_display_label(str(ppt_id)) if ppt_id else "the current presentation"
    return (
        f"I can add up to {allowed} points at a time in {ppt_label}, not {requested}. "
        f"Do you want me to add {allowed} points to slide {slide_num} instead?"
    )

def _slide_limit_confirmation_question(requested: int) -> str:
    return f"You requested {requested} slides. Do you want to continue?"

def _topic_switch_confirmation_question(active_domain: str, new_domain: str) -> str:
    return (
        f"We are currently working within '{active_domain}'. "
        f"Do you want to switch to a new topic domain: '{new_domain}'?"
    )

def _parse_bullet_add_request(change_content: str) -> Optional[int]:
    """
    Detect requests like:
    - add 3 points
    - add one point
    - add a point
    Returns how many bullets should be generated, or None if this is not a bullet-add request.
    """
    text = (change_content or "").lower().strip()
    if not text or "add" not in text:
        return None
    # Normalize common shorthand used by the routing layer.
    text = text.replace("point(s)", "points").replace("bullet(s)", "bullets")
    if not re.search(r"\b(point|points|bullet|bullets|item|items)\b", text):
        return None

    num_match = re.search(r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty)\s*(?:bullet\s*)?(?:point|points|bullet|bullets|item|items)\b", text)
    if num_match:
        parsed = _parse_small_number_token(num_match.group(1))
        return max(1, parsed or 1)

    if re.search(r"\b(?:one|a|an)\s+(?:bullet\s*)?(?:point|points|bullet|bullets|item|items)\b", text):
        return 1

    if re.search(r"\b(?:point|points|bullet|bullets|item|items)\b", text):
        return 2

    return 1

def _parse_point_side_hint(change_content: str) -> Optional[str]:
    text = (change_content or "").lower().strip()
    if not text:
        return None
    if re.search(r"\b(both(?:\s+the)?\s+sides?|both side|left and right|right and left)\b", text):
        return "both"
    if re.search(r"\b(left side|left-side|left column|on the left)\b", text):
        return "left"
    if re.search(r"\b(right side|right-side|right column|on the right)\b", text):
        return "right"
    return None

def _resolve_add_points_side_hint(change_content: str, slide: Optional[dict] = None) -> Optional[str]:
    explicit = _parse_point_side_hint(change_content)
    if explicit:
        return explicit
    if not isinstance(slide, dict):
        return None
    if _safe_str(slide.get("layout", ""), "").lower() != "two_column":
        return None
    requested = _parse_bullet_add_request(change_content)
    text = _norm_text(change_content)
    # On two-column slides, a singular "add one point" request should touch both
    # columns unless the user explicitly targeted only left or right.
    if requested == 1 and re.search(r"\b(?:add|insert|append)\b", text, re.IGNORECASE):
        return "both"
    return None

def _resolved_add_point_count(num: Optional[int], side_hint: Optional[str]) -> int:
    try:
        resolved = max(1, int(num or 1))
    except (ValueError, TypeError):
        resolved = 1
    # "one point on both sides" needs one point per column.
    if side_hint == "both":
        resolved = max(2, resolved)
    return min(resolved, _MAX_POINTS_PER_REQUEST)

def _normalize_two_column_generated_points(points: list[str]) -> list[str]:
    normalized: list[str] = []
    for point in points or []:
        text = clean_icon_tokens(point)
        if not text:
            continue
        combined = re.match(
            r"^\s*left\s*[:\-]\s*(.*?)\s*(?:\||;|/)\s*right\s*[:\-]\s*(.*?)\s*$",
            text,
            re.IGNORECASE,
        )
        if combined:
            left_text = clean_icon_tokens(combined.group(1))
            right_text = clean_icon_tokens(combined.group(2))
            if left_text:
                normalized.append(left_text)
            if right_text:
                normalized.append(right_text)
            continue
        text = re.sub(r"^\s*(?:left|right)\s*[:\-]\s*", "", text, flags=re.IGNORECASE).strip()
        if text:
            normalized.append(text)
    return _dedupe_preserve_order(normalized)

def _dedupe_preserve_order(items: list) -> list:
    seen = set()
    cleaned = []
    for item in items or []:
        text = clean_icon_tokens(str(item)).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        cleaned.append(text)
    return cleaned

def _looks_like_duplicate_point(candidate: str, existing_points: list) -> bool:
    cand = _norm_text(candidate)
    if not cand:
        return True
    cand_words = {w for w in cand.split() if w not in {"the", "a", "an", "and", "of", "to", "for", "in", "on", "with"}}
    if not cand_words:
        return True
    for item in existing_points or []:
        ex = _norm_text(item)
        if not ex:
            continue
        if cand == ex or cand in ex or ex in cand:
            return True
        ex_words = {w for w in ex.split() if w not in {"the", "a", "an", "and", "of", "to", "for", "in", "on", "with"}}
        if not ex_words:
            continue
        overlap = len(cand_words & ex_words)
        union = len(cand_words | ex_words) or 1
        if overlap >= 3 or (overlap / union) >= 0.6:
            return True
    return False

def _compress_text_for_shortening(text: str) -> str:
    cleaned = _safe_str(text)
    if not cleaned:
        return ""
    cleaned = re.sub(r"\([^)]*\)", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    for sep in [";", " - ", " -- ", ",", ":"]:
        if sep in cleaned and len(cleaned) > 40:
            cleaned = cleaned.split(sep, 1)[0].strip()
    cleaned = re.sub(
        r"\b(very|highly|significantly|substantially|extremely|really|actually|basically|clearly|simply|just)\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.;:-")
    return cleaned

def _slide_text_size(slide: dict) -> int:
    if not isinstance(slide, dict):
        return 0
    parts = [slide.get("title", ""), slide.get("subtitle", "")]
    if _safe_str(slide.get("layout", "")).lower() == "two_column":
        parts.extend(flatten_slide_content(slide.get("left_points", [])))
        parts.extend(flatten_slide_content(slide.get("right_points", [])))
    elif _safe_str(slide.get("layout", "")).lower() == "timeline":
        parts.extend([f"{s.get('label', '')} {s.get('detail', '')}" for s in (slide.get("steps", []) or []) if isinstance(s, dict)])
    elif _safe_str(slide.get("layout", "")).lower() == "icon_grid":
        parts.extend([f"{g.get('title', '')} {g.get('detail', '')}" for g in (slide.get("grid_items", []) or []) if isinstance(g, dict)])
    else:
        parts.extend(flatten_slide_content(slide.get("content", [])))
    return sum(len(_safe_str(p)) for p in parts)

def _shorten_slide_structure(slide: dict) -> dict:
    if not isinstance(slide, dict):
        return slide
    shortened = copy.deepcopy(slide)
    layout = _safe_str(shortened.get("layout", "bullets"), "bullets").lower()
    if layout == "two_column":
        left = [_compress_text_for_shortening(p) for p in flatten_slide_content(shortened.get("left_points", []))]
        right = [_compress_text_for_shortening(p) for p in flatten_slide_content(shortened.get("right_points", []))]
        left = [p for p in left if p]
        right = [p for p in right if p]
        if len(left) > 2:
            left = left[:2]
        if len(right) > 2:
            right = right[:2]
        shortened["left_points"] = left
        shortened["right_points"] = right
        shortened["left"] = copy.deepcopy(left)
        shortened["right"] = copy.deepcopy(right)
        shortened["content"] = _dedupe_preserve_order(left + right)
    elif layout == "timeline":
        steps = []
        for step in shortened.get("steps", []) or []:
            if not isinstance(step, dict):
                continue
            steps.append({
                "label": _compress_text_for_shortening(step.get("label", "")) or _safe_str(step.get("label", "")),
                "detail": _compress_text_for_shortening(step.get("detail", "")),
            })
        shortened["steps"] = steps[:4]
        shortened["content"] = _dedupe_preserve_order([
            f"{s.get('label', '')}: {s.get('detail', '')}".strip(": ").strip()
            for s in steps[:4]
            if s.get("label") or s.get("detail")
        ])
    elif layout == "icon_grid":
        items = []
        for item in shortened.get("grid_items", []) or []:
            if not isinstance(item, dict):
                continue
            items.append({
                "icon": item.get("icon", "•"),
                "title": _compress_text_for_shortening(item.get("title", "")) or _safe_str(item.get("title", "")),
                "detail": _compress_text_for_shortening(item.get("detail", "")),
            })
        shortened["grid_items"] = items[:4]
        shortened["content"] = _dedupe_preserve_order([
            f"{item.get('title', '')}: {item.get('detail', '')}".rstrip(": ").strip()
            for item in items[:4]
        ])
    else:
        content = [_compress_text_for_shortening(p) for p in flatten_slide_content(shortened.get("content", []))]
        content = [p for p in content if p]
        if len(content) > 3:
            content = content[:3]
        shortened["content"] = content
    shortened["_user_modified"] = True
    return shortened

def _normalize_match_text(value: str) -> str:
    return _norm_text(value).strip()

def _extract_remove_instruction(change_content: str) -> dict:
    text = _safe_str(change_content)
    lower = text.lower()
    if not re.search(r"\b(remove|delete|drop|erase|take out)\b", lower):
        return {}
    side_hint = _parse_point_side_hint(text)
    if re.search(r"\b(last|final|latest)\b", lower):
        return {"mode": "last", "side_hint": side_hint}
    idx_match = re.search(r"\b(?:point|bullet)\s*(?:number\s*)?#?(\d+)\b", lower)
    if idx_match:
        try:
            return {"mode": "index", "index": int(idx_match.group(1)), "side_hint": side_hint}
        except ValueError:
            pass
    content = re.sub(r"(?i)\b(?:remove|delete|drop|erase|take out)\b", "", text)
    content = re.sub(r"(?i)\b(?:from|on|in)\s+(?:the\s+)?(?:left|right|both)\s+(?:side|sides|column|columns)\b", "", content)
    content = re.sub(r"(?i)\b(?:left|right|both)\s+(?:side|sides|column|columns)\b", "", content)
    content = re.sub(r"(?i)\b(?:last|final|latest)\s+(?:point|bullet)?\b", "", content)
    content = re.sub(r"\s+", " ", content).strip(" ,.;:-")
    if content:
        return {"mode": "content", "content": content, "side_hint": side_hint}
    return {"mode": "last", "side_hint": side_hint}

def _remove_point_from_list(points: list, request: dict) -> list:
    items = _dedupe_preserve_order(points)
    if not items:
        return items
    mode = request.get("mode")
    if mode == "last":
        if items:
            items.pop()
        return items
    if mode == "index":
        idx = request.get("index")
        try:
            idx = int(idx)
        except (TypeError, ValueError):
            return items
        if 1 <= idx <= len(items):
            items.pop(idx - 1)
        return items
    if mode == "content":
        target = _normalize_match_text(request.get("content", ""))
        if not target:
            return items
        for i, item in enumerate(items):
            cand = _normalize_match_text(item)
            if cand == target or target in cand or cand in target:
                items.pop(i)
                return items
        for i, item in enumerate(items):
            cand_words = set(_normalize_match_text(item).split())
            target_words = set(target.split())
            if cand_words and target_words and len(cand_words & target_words) >= max(2, min(3, len(target_words))):
                items.pop(i)
                return items
    return items

def _extract_two_column_side_from_text(text: str) -> Optional[str]:
    lower = _safe_str(text).lower()
    if re.search(r"\b(left|left side|left-side|left column|on the left)\b", lower):
        return "left"
    if re.search(r"\b(right|right side|right-side|right column|on the right)\b", lower):
        return "right"
    if re.search(r"\b(both|both side|both sides|left and right|right and left)\b", lower):
        return "both"
    return None

def _remove_points_from_slide(slide: dict, request: dict) -> bool:
    if not isinstance(slide, dict) or not request:
        return False
    layout = _safe_str(slide.get("layout", "bullets"), "bullets") or "bullets"
    side_hint = request.get("side_hint")
    before = _slide_point_state(slide)
    changed = False

    if layout == "two_column":
        left_points = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", [])))
        right_points = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", [])))
        if side_hint == "left":
            left_points = _remove_point_from_list(left_points, request)
            changed = left_points != _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", [])))
        elif side_hint == "right":
            right_points = _remove_point_from_list(right_points, request)
            changed = right_points != _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", [])))
        else:
            return False
        slide["left_points"] = left_points
        slide["right_points"] = right_points
        slide["left"] = copy.deepcopy(left_points)
        slide["right"] = copy.deepcopy(right_points)
        slide["content"] = _dedupe_preserve_order(flatten_slide_content(left_points) + flatten_slide_content(right_points))
        slide["layout"] = "two_column"
    else:
        content = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
        new_content = _remove_point_from_list(content, request)
        changed = new_content != content
        slide["content"] = new_content

    slide["_user_modified"] = bool(changed or before != _slide_point_state(slide))
    _sync_two_column_aliases(slide)
    return slide["_user_modified"]

def _merge_updated_slide(slides: list, slide_num: int, updated_slide: dict) -> list:
    merged = copy.deepcopy(slides or [])
    if not (1 <= int(slide_num) <= len(merged)):
        return merged
    merged[int(slide_num) - 1] = copy.deepcopy(updated_slide)
    return merged

def _sync_two_column_aliases(slide: dict) -> dict:
    if not isinstance(slide, dict):
        return slide
    if _safe_str(slide.get("layout", ""), "").lower() == "two_column":
        left = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", slide.get("left", []))))
        right = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", slide.get("right", []))))
        slide["left_points"] = left
        slide["right_points"] = right
        slide["left"] = copy.deepcopy(left)
        slide["right"] = copy.deepcopy(right)
    return slide

def _point_text_to_step(point: str, idx: int) -> dict:
    text = clean_icon_tokens(str(point)).strip()
    if not text:
        return {"label": seq_icon(idx), "detail": ""}
    if ":" in text:
        label, detail = [p.strip() for p in text.split(":", 1)]
    elif len(text) > 28:
        label, detail = text[:24].strip(), text[24:].strip()
    else:
        label, detail = seq_icon(idx), text
    return {"label": label or seq_icon(idx), "detail": detail}

def _point_text_to_grid_item(point: str, idx: int) -> dict:
    text = clean_icon_tokens(str(point)).strip()
    if not text:
        return {"icon": seq_icon(idx), "title": f"Point {idx + 1}", "detail": ""}
    if ":" in text:
        title, detail = [p.strip() for p in text.split(":", 1)]
    else:
        words = text.split()
        title = " ".join(words[:4]).strip() or f"Point {idx + 1}"
        detail = " ".join(words[4:]).strip()
    return {"icon": seq_icon(idx), "title": title, "detail": detail}

def _slide_point_state(slide: dict) -> dict:
    layout = _safe_str(slide.get("layout", "bullets"), "bullets") or "bullets"
    state = {"layout": layout}
    if layout == "two_column":
        state["left_points"] = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", slide.get("left", []))))
        state["right_points"] = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", slide.get("right", []))))
    elif layout == "timeline":
        state["steps"] = [
            {
                "label": _safe_str(s.get("label", "")),
                "detail": _safe_str(s.get("detail", "")),
            }
            for s in (slide.get("steps", []) or [])
            if isinstance(s, dict)
        ]
    elif layout == "icon_grid":
        state["grid_items"] = [
            {
                "icon": _safe_str(g.get("icon", "")),
                "title": _safe_str(g.get("title", "")),
                "detail": _safe_str(g.get("detail", "")),
            }
            for g in (slide.get("grid_items", []) or [])
            if isinstance(g, dict)
        ]
    elif layout == "section_index":
        state["sections"] = _dedupe_preserve_order(flatten_slide_content(slide.get("sections", slide.get("content", []))))
    elif layout == "table":
        state["table_columns"] = _dedupe_preserve_order(flatten_slide_content(slide.get("table_columns", [])))
        state["table_rows"] = [
            [_safe_str(cell, "") for cell in row]
            for row in (slide.get("table_rows", []) or [])
            if isinstance(row, list)
        ]
    else:
        state["content"] = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
    return state

def _needs_layout_aware_update(change_content: str, slide: Optional[dict] = None) -> bool:
    text = _safe_str(change_content).lower()
    if not text:
        return False
    layout = _safe_str((slide or {}).get("layout", ""), "").lower()
    if layout in {"table", "two_column"} and re.search(r"\b(add|insert|update|modify|change|replace)\b", text):
        return True
    return bool(re.search(
        r"\b(comparison|compare|vs|versus|advantages?\s+and\s+disadvantages?|pros?\s+and\s+cons?|"
        r"real[\s-]?world examples?|case studi(?:es|y)|table|rows?|columns?)\b",
        text,
        re.IGNORECASE,
    ))

def _append_unique_table_row(rows: list[list[str]], row: list[str]) -> list[list[str]]:
    cleaned = [_safe_str(cell, "") for cell in row]
    key = " | ".join(cleaned).casefold()
    existing = {" | ".join([_safe_str(c, "") for c in r]).casefold() for r in rows if isinstance(r, list)}
    if key and key not in existing:
        rows.append(cleaned)
    return rows

def _extract_comparison_subjects(text: str, slide: Optional[dict] = None) -> tuple[str, str]:
    source = _safe_str(text, "")
    patterns = [
        r"\bbetween\s+(.+?)\s+and\s+(.+?)(?:$|[,.?;])",
        r"\bcompare\s+(.+?)\s+(?:and|with|vs|versus)\s+(.+?)(?:$|[,.?;])",
        r"\b(.+?)\s+(?:vs|versus)\s+(.+?)(?:$|[,.?;])",
    ]
    for pat in patterns:
        m = re.search(pat, source, re.IGNORECASE)
        if not m:
            continue
        left = re.sub(r"(?i)\b(?:a|an|the|one|comparison|table|chart|slide|between|compare)\b", "", m.group(1)).strip(" ,.;:-")
        right = re.sub(r"(?i)\b(?:a|an|the|one|comparison|table|chart|slide)\b", "", m.group(2)).strip(" ,.;:-")
        if left and right:
            return left[:40], right[:40]
    title = _safe_str((slide or {}).get("title", ""), "")
    if re.search(r"\b(?:vs|versus)\b", title, re.IGNORECASE):
        left, right = re.split(r"\b(?:vs|versus)\b", title, maxsplit=1, flags=re.IGNORECASE)
        if left.strip() and right.strip():
            return left.strip()[:40], right.strip()[:40]
    return "Option A", "Option B"

def _build_comparison_table_rows(left: str, right: str, topic_hint: str = "") -> list[list[str]]:
    l_norm = _norm_text(left)
    r_norm = _norm_text(right)
    if {"ml", "machine learning"} & {l_norm, r_norm} or "machine learning" in f"{l_norm} {r_norm}" or "deep learning" in f"{l_norm} {r_norm}":
        return [
            ["Core idea", "Learns patterns from structured features", "Uses layered neural networks to learn representations"],
            ["Data needs", "Works well with moderate, prepared datasets", "Usually needs large datasets for strong results"],
            ["Feature work", "Often depends on manual feature engineering", "Learns features automatically from raw data"],
            ["Best fit", "Forecasting, scoring, classification, and tabular tasks", "Images, speech, language, and complex perception tasks"],
            ["Explainability", "Often easier to inspect and explain", "Can be harder to interpret without special tools"],
        ]
    return [
        ["Definition", f"How {left} is typically understood", f"How {right} is typically understood"],
        ["Strength", f"Where {left} performs best", f"Where {right} performs best"],
        ["Limitation", f"Main constraint for {left}", f"Main constraint for {right}"],
        ["Use case", f"Practical example for {left}", f"Practical example for {right}"],
    ]

def _llm_layout_aware_restructure(slide: dict, change_content: str) -> Optional[dict]:
    """
    Optional fallback: ask LLM for a structured rewrite of the CURRENT slide only.
    Returns normalized slide dict or None when not usable.
    """
    if not isinstance(slide, dict):
        return None
    layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()
    allowed_layout = layout if layout in {"table", "two_column", "bullets"} else "bullets"
    slide_json = json.dumps(slide, ensure_ascii=True)
    prompt = f"""
You are restructuring ONE slide payload for a PPT editor.

User request: "{change_content}"
Required layout: {allowed_layout}
Current slide JSON:
{slide_json}

Rules:
1. Return JSON for THIS slide only. Do not create/remove/reorder slides.
2. Preserve title and subtitle unless the request explicitly asks otherwise.
3. If layout=table, return table_columns and table_rows with consistent row lengths.
4. If layout=two_column, return left_title, right_title, left_points, right_points.
5. If request mentions "advantages and disadvantages" or "pros and cons", keep separated sections.
6. If request mentions "real-world examples", integrate examples into existing categories/rows.
7. Avoid generic single-point appends unless explicitly requested.
8. No markdown, no icons, no explanatory text. Output JSON only.
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=900,
        )
        raw = _safe_str(resp.choices[0].message.content, "")
        parsed = extract_first_json(raw) if raw else {}
        if not isinstance(parsed, dict):
            return None
        # Keep identity/layout deterministic; allow LLM to enrich fields.
        merged = copy.deepcopy(slide)
        merged.update(parsed)
        merged["layout"] = _safe_str(merged.get("layout", allowed_layout), allowed_layout)
        normalized = ensure_editor_id(merged)
        return normalized if isinstance(normalized, dict) else None
    except Exception:
        return None

def _apply_layout_aware_update(slide: dict, change_content: str) -> bool:
    if not isinstance(slide, dict):
        return False
    text = _safe_str(change_content, "")
    low = text.lower()
    before = _slide_point_state(slide)
    layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()

    wants_comparison = bool(re.search(r"\b(comparison|compare|vs|versus|table)\b", low))
    wants_adv_dis = bool(re.search(r"\b(advantages?\s+and\s+disadvantages?|pros?\s+and\s+cons?)\b", low))
    wants_examples = bool(re.search(r"\b(real[\s-]?world examples?|case studi(?:es|y)|examples?)\b", low))

    if layout == "table" or wants_comparison:
        left_subject, right_subject = _extract_comparison_subjects(text, slide)
        cols = _dedupe_preserve_order(flatten_slide_content(slide.get("table_columns", [])))
        rows = [list(r) for r in (slide.get("table_rows", []) or []) if isinstance(r, list)]
        if wants_adv_dis:
            cols = ["Category", "Advantages", "Disadvantages"]
            rows = _append_unique_table_row(rows, ["Performance", "Fast and efficient on standard tasks", "May underperform on very complex patterns"])
            rows = _append_unique_table_row(rows, ["Interpretability", "Usually easier to interpret", "Can oversimplify nuanced relationships"])
        elif wants_comparison:
            cols = ["Aspect", left_subject, right_subject]
            rows = _build_comparison_table_rows(left_subject, right_subject, text)
        if wants_examples:
            rows = _append_unique_table_row(rows, ["Real-world example", f"Common use of {left_subject} in business workflows", f"Common use of {right_subject} in complex perception or language tasks"])
        if not rows and slide.get("content"):
            for line in flatten_slide_content(slide.get("content", [])):
                if "|" in line:
                    parts = [_safe_str(p, "") for p in line.split("|")]
                    if len(parts) >= 3:
                        rows = _append_unique_table_row(rows, parts[:max(3, len(cols) or 3)])
        if not cols:
            cols = ["Aspect", "Option A", "Option B"]
        slide["layout"] = "table"
        if _is_placeholder_slide_title(slide.get("title", ""), None):
            slide["title"] = f"{left_subject} vs {right_subject}"
        slide["table_columns"] = cols
        slide["table_rows"] = rows[:8]
        slide["content"] = []
        slide["_user_modified"] = True
        return before != _slide_point_state(slide)

    if layout == "two_column" or wants_adv_dis:
        left_points = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", slide.get("left", []))))
        right_points = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", slide.get("right", []))))
        if not left_points and not right_points:
            base = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
            mid = max(1, len(base) // 2)
            left_points, right_points = base[:mid], base[mid:]
        if wants_adv_dis:
            slide["left_title"] = "Advantages"
            slide["right_title"] = "Disadvantages"
            left_points = _dedupe_preserve_order(left_points + ["Higher efficiency in suitable tasks", "Easier operational scaling"])
            right_points = _dedupe_preserve_order(right_points + ["May require high-quality input data", "Can become costly at larger scale"])
        if wants_examples:
            left_points = _dedupe_preserve_order(left_points + ["Real-world example: deployment in production operations"])
            right_points = _dedupe_preserve_order(right_points + ["Real-world challenge: model drift and maintenance overhead"])
        slide["layout"] = "two_column"
        slide["left_points"] = left_points[:8]
        slide["right_points"] = right_points[:8]
        slide["left"] = copy.deepcopy(slide["left_points"])
        slide["right"] = copy.deepcopy(slide["right_points"])
        slide["content"] = _dedupe_preserve_order(slide["left_points"] + slide["right_points"])
        slide["_user_modified"] = True
        _sync_two_column_aliases(slide)
        return before != _slide_point_state(slide)

    if wants_examples:
        content = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
        content = _dedupe_preserve_order(content + [
            "Real-world example: practical deployment in an industry workflow",
            "Real-world example: measurable business impact after adoption",
        ])
        slide["content"] = content[:12]
        slide["_user_modified"] = True
        return before != _slide_point_state(slide)

    # Optional LLM fallback for higher-quality structuring when rules did not fire.
    llm_slide = _llm_layout_aware_restructure(slide, change_content)
    if llm_slide and isinstance(llm_slide, dict):
        slide.clear()
        slide.update(llm_slide)
        slide["_user_modified"] = True
        _sync_two_column_aliases(slide)
        return before != _slide_point_state(slide)

    return False

def _chunk_list(items: list, chunk_size: int) -> list[list]:
    if not items:
        return []
    size = max(1, int(chunk_size))
    return [items[i:i + size] for i in range(0, len(items), size)]

def _clone_slide_with_chunk(
    slide: dict,
    chunk: dict,
    part_idx: int,
    total_parts: int,
    *,
    title_override: Optional[str] = None,
    add_part_suffix: bool = True,
) -> dict:
    cloned = copy.deepcopy(slide)
    title = _safe_str(cloned.get("title", "Slide"), "Slide")
    if title_override:
        cloned["title"] = title_override
    elif add_part_suffix and total_parts > 1:
        cloned["title"] = f"{title} (Part {part_idx + 1})"
    layout = _safe_str(cloned.get("layout", "bullets"), "bullets").lower()
    if layout == "table":
        cloned["table_columns"] = copy.deepcopy(chunk.get("table_columns", cloned.get("table_columns", [])))
        cloned["table_rows"] = copy.deepcopy(chunk.get("table_rows", []))
        cols = [_safe_str(c, "") for c in cloned.get("table_columns", [])]
        rows = cloned.get("table_rows", []) or []
        cloned["content"] = [" | ".join(cols)] if cols else []
        for row in rows:
            if isinstance(row, list):
                cloned["content"].append(" | ".join([_safe_str(c, "") for c in row]))
    elif layout == "two_column":
        cloned["left_points"] = copy.deepcopy(chunk.get("left_points", []))
        cloned["right_points"] = copy.deepcopy(chunk.get("right_points", []))
        cloned["left"] = copy.deepcopy(cloned["left_points"])
        cloned["right"] = copy.deepcopy(cloned["right_points"])
        cloned["content"] = _dedupe_preserve_order(
            flatten_slide_content(cloned.get("left_points", [])) +
            flatten_slide_content(cloned.get("right_points", []))
        )
    else:
        cloned["content"] = copy.deepcopy(chunk.get("content", []))
    cloned["_user_modified"] = True
    return normalize_slide(cloned)

def _derive_subtopic_title(base_title: str, points: list[str], idx: int) -> str:
    for raw in points or []:
        text = _safe_str(raw, "")
        if not text:
            continue
        if ":" in text:
            head = _safe_str(text.split(":", 1)[0], "")
        elif "|" in text:
            head = _safe_str(text.split("|", 1)[0], "")
        else:
            head = " ".join(text.split()[:5]).strip()
        if head:
            return f"{base_title}: {head}"
    return f"{base_title}: Subtopic {idx + 1}"

def _split_single_slide_to_target(slide: dict, target_parts: int) -> list[dict]:
    if not isinstance(slide, dict):
        return [slide]
    target = max(1, int(target_parts))
    layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()
    base_title = _safe_str(slide.get("title", "Slide"), "Slide")
    parts: list[dict] = []
    if layout == "table":
        cols = [_safe_str(c, "") for c in (slide.get("table_columns", []) or []) if _safe_str(c, "")]
        rows = [list(r) for r in (slide.get("table_rows", []) or []) if isinstance(r, list) and r]
        if not rows:
            rows = [[r] for r in flatten_slide_content(slide.get("content", [])) if r]
        chunk_size = max(1, (len(rows) + target - 1) // target)
        row_chunks = _chunk_list(rows, chunk_size)[:target]
        for idx, chunk in enumerate(row_chunks):
            row_points = [" | ".join([_safe_str(c, "") for c in row]) for row in chunk]
            title = _derive_subtopic_title(base_title, row_points, idx)
            parts.append(_clone_slide_with_chunk(
                slide,
                {"table_columns": cols or ["Aspect", "Option A", "Option B"], "table_rows": chunk},
                idx,
                len(row_chunks),
                title_override=title,
                add_part_suffix=False,
            ))
        return parts if parts else [slide]
    if layout == "two_column":
        left = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", slide.get("left", []))))
        right = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", slide.get("right", []))))
        combined = _dedupe_preserve_order(left + right)
        if not combined:
            combined = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
        chunk_size = max(1, (len(combined) + target - 1) // target)
        chunks = _chunk_list(combined, chunk_size)[:target]
        for idx, chunk in enumerate(chunks):
            half = max(1, len(chunk) // 2)
            lp = chunk[:half]
            rp = chunk[half:] if len(chunk) > 1 else []
            title = _derive_subtopic_title(base_title, chunk, idx)
            parts.append(_clone_slide_with_chunk(
                slide,
                {"left_points": lp, "right_points": rp},
                idx,
                len(chunks),
                title_override=title,
                add_part_suffix=False,
            ))
        return parts if parts else [slide]
    points = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
    chunk_size = max(1, (len(points) + target - 1) // target)
    chunks = _chunk_list(points, chunk_size)[:target] if points else [[]]
    for idx, chunk in enumerate(chunks):
        title = _derive_subtopic_title(base_title, chunk, idx)
        parts.append(_clone_slide_with_chunk(
            slide,
            {"content": chunk},
            idx,
            len(chunks),
            title_override=title,
            add_part_suffix=False,
        ))
    return parts if parts else [slide]

def _expand_target_slide_in_deck(slides: list, slide_num: int, target_parts: int) -> tuple[list, int]:
    base = copy.deepcopy(slides or [])
    if not (1 <= int(slide_num) <= len(base)):
        return base, 0
    target = max(1, int(target_parts))
    original = base[int(slide_num) - 1]
    parts = _split_single_slide_to_target(original, target)
    if not parts:
        return base, 0
    rebuilt = base[:int(slide_num) - 1] + parts + base[int(slide_num):]
    rebuilt = refresh_section_index_slide(rebuilt)
    return rebuilt, max(0, len(parts) - 1)

def _split_slide_into_parts(slide: dict) -> list[dict]:
    if not isinstance(slide, dict):
        return [slide]
    layout = _safe_str(slide.get("layout", "bullets"), "bullets").lower()
    if layout == "table":
        cols = [_safe_str(c, "") for c in (slide.get("table_columns", []) or []) if _safe_str(c, "")]
        rows = [list(r) for r in (slide.get("table_rows", []) or []) if isinstance(r, list) and r]
        if len(rows) <= 4:
            return [slide]
        row_chunks = _chunk_list(rows, 4)
        return [
            _clone_slide_with_chunk(slide, {"table_columns": cols, "table_rows": chunk}, idx, len(row_chunks))
            for idx, chunk in enumerate(row_chunks)
        ]
    if layout == "two_column":
        left = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", slide.get("left", []))))
        right = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", slide.get("right", []))))
        max_per_side = 4
        if max(len(left), len(right)) <= max_per_side:
            return [slide]
        left_chunks = _chunk_list(left, max_per_side) or [[]]
        right_chunks = _chunk_list(right, max_per_side) or [[]]
        count = max(len(left_chunks), len(right_chunks))
        parts: list[dict] = []
        for idx in range(count):
            chunk = {
                "left_points": left_chunks[idx] if idx < len(left_chunks) else [],
                "right_points": right_chunks[idx] if idx < len(right_chunks) else [],
            }
            parts.append(_clone_slide_with_chunk(slide, chunk, idx, count))
        return parts
    points = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
    if len(points) <= 5:
        return [slide]
    chunks = _chunk_list(points, 5)
    return [
        _clone_slide_with_chunk(slide, {"content": chunk}, idx, len(chunks))
        for idx, chunk in enumerate(chunks)
    ]

def _expand_existing_deck_slides(slides: list, target_count: int) -> tuple[list, int]:
    """
    Expand current deck to target count with new, meaningful subtopic slides.
    Returns (updated_slides, added_count).
    """
    base = copy.deepcopy(slides or [])
    if target_count <= len(base):
        return base, 0
    need_add = target_count - len(base)

    # Deck-level expansion should add new topics, not cloned "(Part 2)" slides.
    additions: list[dict] = _generate_semantic_expansion_slides(base, need_add)
    if len(additions) < need_add:
        additions.extend(_generate_deterministic_expansion_slides(base + additions, need_add - len(additions)))
    updated = base + additions[:need_add]
    updated = refresh_section_index_slide(updated)
    return updated, min(len(additions), need_add)

def _is_structural_slide(slide: dict) -> bool:
    layout = _safe_str((slide or {}).get("layout", ""), "").lower()
    return layout in {"title_cover", "section_index"} or _slide_looks_like_thank_you(slide)

def _condense_existing_deck_slides(slides: list, target_count: int) -> tuple[list, int]:
    """
    Reduce the deck to an exact target count by preserving bookend slides and
    merging/removing the least essential body slides.
    """
    base = [copy.deepcopy(s) for s in (slides or []) if isinstance(s, dict)]
    target = max(1, int(target_count or 1))
    if target >= len(base):
        return base, 0
    if target == 1:
        body_points = []
        for slide in base:
            if _slide_looks_like_thank_you(slide):
                continue
            title = _safe_str(slide.get("title", ""), "")
            for point in flatten_slide_content(slide.get("content", []))[:2]:
                if point:
                    body_points.append(f"{title}: {point}".strip(": "))
        summary = {
            "title": _safe_str(st.session_state.get("topic", ""), "") or "Executive Summary",
            "subtitle": "Condensed presentation",
            "layout": "bullets",
            "icon": "▸",
            "content": _dedupe_preserve_order(body_points)[:6],
            "style": {},
            "_user_modified": True,
        }
        return [normalize_slide(summary)], max(0, len(base) - 1)

    title_slide = base[0] if base else None
    thank_slide = base[-1] if base and _slide_looks_like_thank_you(base[-1]) else None
    body_start = 1 if title_slide and _safe_str(title_slide.get("layout", "")).lower() == "title_cover" else 0
    body_end = len(base) - (1 if thank_slide else 0)
    body = [
        s for i, s in enumerate(base[body_start:body_end], start=body_start)
        if _safe_str(s.get("layout", "")).lower() != "section_index"
    ]

    reserved = []
    if title_slide:
        reserved.append(title_slide)
    if thank_slide and target > len(reserved) + 1:
        tail = [thank_slide]
    else:
        tail = []
    body_slots = max(0, target - len(reserved) - len(tail))
    if body_slots <= 0:
        condensed = (reserved + tail)[:target]
        return refresh_section_index_slide(condensed), len(base) - len(condensed)

    work = copy.deepcopy(body)
    topic_hint = _safe_str(st.session_state.get("topic", ""), "Presentation")
    while len(work) > body_slots:
        # Merge neighboring body slides to preserve flow instead of dropping
        # content abruptly.
        merge_idx = max(0, len(work) - 2)
        if merge_idx + 1 < len(work):
            work[merge_idx] = merge_slide_payloads(work[merge_idx], work[merge_idx + 1], topic=topic_hint)
            del work[merge_idx + 1]
        else:
            work.pop()
    work = [_shorten_slide_structure(s) for s in work[:body_slots]]
    condensed = reserved + work + tail
    condensed = condensed[:target]
    condensed = refresh_section_index_slide(condensed)
    return condensed, len(base) - len(condensed)

def _design_style_for_request(request: str, sub_intent: str = "") -> dict:
    text = f"{request} {sub_intent}".lower()
    style = {
        "pattern_name": "minimal-lines",
        "surface": "light",
        "header_variant": "banded",
        "card_variant": "outline",
        "footer_variant": "line",
        "badge_shape": "rect",
        "accent_rotation": "auto",
    }
    if re.search(r"\bdark\b", text):
        style.update({"pattern_name": "dark-contrast", "surface": "tint", "header_variant": "solid"})
    elif re.search(r"\bmodern|sleek|minimal|minimalist|clean\b", text):
        style.update({"pattern_name": "modern-minimal", "surface": "light", "header_variant": "split", "card_variant": "soft"})
    elif re.search(r"\bbusiness|corporate|professional|investor\b", text):
        style.update({"pattern_name": "executive-grid", "header_variant": "solid", "card_variant": "outline"})
    return style

def _apply_design_to_existing_slides(slides: list, request: str = "", sub_intent: str = "") -> list:
    style = _design_style_for_request(request, sub_intent)
    updated = []
    for idx, slide in enumerate(slides or [], start=1):
        if not isinstance(slide, dict):
            continue
        s = copy.deepcopy(slide)
        existing_style = s.get("style") if isinstance(s.get("style"), dict) else {}
        s["style"] = {**existing_style, **style}
        if _is_placeholder_slide_title(s.get("title", ""), idx):
            s["title"] = _derive_meaningful_slide_title(s, idx)
        s["_user_modified"] = True
        updated.append(s)
    return updated

def _is_bad_expansion_title(title: str, existing_titles: Optional[set[str]] = None) -> bool:
    norm = _norm_text(title)
    if not norm:
        return True
    if existing_titles and norm in existing_titles:
        return True
    return bool(re.search(r"\b(?:part|continuation|continued|summary\s+part)\s*\d*\b", norm, re.IGNORECASE))

def _generate_semantic_expansion_slides(slides: list, needed: int) -> list[dict]:
    needed = max(0, int(needed or 0))
    if needed <= 0:
        return []
    existing_titles = [_safe_str(s.get("title", ""), "") for s in slides if isinstance(s, dict)]
    compact = summarize_slides_for_llm(slides)
    prompt = f"""
Create {needed} NEW PowerPoint slide objects that expand this deck with meaningful missing subtopics.

Existing slides:
{json.dumps(compact, indent=2)}

Rules:
- Generate genuinely new slide topics, not split copies of existing slides.
- Never use title suffixes like "(Part 2)", "Continuation", or "Continued".
- Do not duplicate existing titles or produce summary/review duplicates.
- Each new slide must cover a distinct subtopic, application, implication, example, risk, opportunity, or future direction that logically expands the deck.
- Use concise, useful bullets grounded in the deck's topic, with new wording and new examples.
- Return ONLY JSON: {{"slides": [{{"title": "...", "subtitle": "", "layout": "bullets", "icon": "▸", "content": ["...", "..."], "style": {{}}}}]}}
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.55,
            max_tokens=900,
        )
        raw = (resp.choices[0].message.content or "").strip()
        if not raw.startswith("{"):
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            raw = m.group() if m else raw
        data = json.loads(raw)
        proposed = data.get("slides", []) if isinstance(data, dict) else []
    except Exception:
        proposed = []
    normalized: list[dict] = []
    seen = {_norm_text(t) for t in existing_titles if t}
    for slide in proposed:
        if not isinstance(slide, dict):
            continue
        title_norm = _norm_text(slide.get("title", ""))
        if _is_bad_expansion_title(slide.get("title", ""), seen):
            continue
        seen.add(title_norm)
        slide["_user_modified"] = True
        normalized.append(normalize_slide(slide))
        if len(normalized) >= needed:
            break
    return normalized

def _generate_deterministic_expansion_slides(slides: list, needed: int) -> list[dict]:
    needed = max(0, int(needed or 0))
    if needed <= 0:
        return []
    existing = {_norm_text(s.get("title", "")) for s in slides if isinstance(s, dict)}
    candidates: list[tuple[str, list[str]]] = []
    for slide in slides or []:
        if not isinstance(slide, dict):
            continue
        base_title = _safe_str(slide.get("title", ""), "Topic")
        points = _dedupe_preserve_order(
            flatten_slide_content(slide.get("content", []))
            + flatten_slide_content(slide.get("left_points", []))
            + flatten_slide_content(slide.get("right_points", []))
        )
        for point in points[:4]:
            words = [w for w in re.findall(r"[A-Za-z][A-Za-z0-9-]+", point) if len(w) > 3]
            focus = " ".join(words[:4]).strip() or base_title
            title = f"Deep Dive: {focus}"
            bullets = [
                f"Why {focus.lower()} matters in this topic area.",
                f"Practical implications connected to {base_title}.",
                "Key considerations for applying this idea responsibly.",
                "Audience takeaway: how this subtopic changes decisions or next steps.",
            ]
            candidates.append((title, bullets))
    if not candidates:
        candidates = [
            ("Real-World Applications", ["Where the topic is used today.", "Common implementation patterns.", "Benefits for users and organizations.", "Important constraints to manage."]),
            ("Risks and Considerations", ["Potential limitations to watch.", "Operational and ethical concerns.", "Ways to reduce unintended outcomes.", "Decision criteria for responsible use."]),
            ("Future Opportunities", ["Emerging directions in the field.", "New capabilities likely to matter next.", "Signals teams should monitor.", "Long-term impact on strategy and work."]),
        ]
    generated: list[dict] = []
    for title, bullets in candidates:
        if _is_bad_expansion_title(title, existing):
            continue
        existing.add(_norm_text(title))
        generated.append(normalize_slide({
            "title": title,
            "subtitle": "",
            "layout": "bullets",
            "icon": "▸",
            "content": bullets,
            "style": {},
            "_user_modified": True,
        }))
        if len(generated) >= needed:
            break
    return generated

def llm_expand_split_prompt_template() -> str:
    """
    Optional prompt template for semantic slide splitting.
    Not used by default; available for higher-quality future fallback.
    """
    return (
        "You are splitting one PPT slide into multiple focused sub-slides.\n"
        "Input: one slide JSON and a target_parts integer.\n\n"
        "Rules:\n"
        "1. Preserve title theme and subtitle continuity.\n"
        "2. Keep the original layout type when possible: table/two_column/bullets.\n"
        "3. Distribute content into logical subtopics; do not invent unrelated themes.\n"
        "4. Preserve important metrics/statements.\n"
        "5. Give each part a unique title that reflects its subtopic.\n"
        "6. Return JSON only: {\"parts\": [slide_json_1, slide_json_2, ...]}.\n"
        "7. Do not remove existing facts; only reorganize/split.\n"
    )

def _generate_add_points(slide: dict, num: int, side_hint: Optional[str] = None, topic_hint: str = "") -> list[str]:
    title = _safe_str(slide.get("title", ""), "this slide") or "this slide"
    subtitle = _safe_str(slide.get("subtitle", ""), "")
    layout = _safe_str(slide.get("layout", "bullets"), "bullets") or "bullets"
    existing_points = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
    if layout == "two_column":
        existing_points = _dedupe_preserve_order(
            existing_points
            + flatten_slide_content(slide.get("left_points", []))
            + flatten_slide_content(slide.get("right_points", []))
        )
    existing_bits = []
    if title:
        existing_bits.append(f"Title: {title}")
    if subtitle:
        existing_bits.append(f"Subtitle: {subtitle}")
    if layout == "two_column":
        left_title = _safe_str(slide.get("left_title", "Left"), "Left")
        right_title = _safe_str(slide.get("right_title", "Right"), "Right")
        existing_bits.append(f"Left column: {left_title}")
        existing_bits.append(f"Right column: {right_title}")
    if topic_hint:
        existing_bits.append(f"Topic focus: {topic_hint}")
    if existing_points:
        existing_bits.append("Existing points: " + " | ".join(existing_points[:8]))
    existing_text = " | ".join(existing_bits)
    topic_instruction = ""
    if topic_hint:
        topic_instruction = (
            f"Every point must be specifically about: {topic_hint}. "
            "Do not generate generic points outside that topic.\n"
        )
    side_text = ""
    if side_hint == "both":
        side_text = " Split the points across both sides of the slide."
    elif side_hint == "left":
        side_text = " Write the points for the left side only."
    elif side_hint == "right":
        side_text = " Write the points for the right side only."
    if layout == "two_column" and side_hint == "both":
        prompt = (
            f"Generate exactly {num} short, distinct point(s) split across the left and right sides of the slide.\n"
            f"{existing_text}\n"
            "Do not repeat existing points. Keep left and right content separate with at least one point for each side.\n"
            "Do not prefix any item with Left: or Right:.\n"
            f"{topic_instruction}"
            f"Return only a JSON list of strings with exactly {num} items."
        )
    else:
        prompt = (
            f"Generate exactly {num} short, distinct point(s) for the slide.\n"
            f"{existing_text}\n"
            f"{side_text}\n"
            f"{topic_instruction}"
            "Return only a JSON list of strings."
        )
    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4,
        max_tokens=300,
    )
    raw = resp.choices[0].message.content.strip()
    if raw.startswith('['):
        points = json.loads(raw)
    else:
        match = re.search(r'\[.*\]', raw, re.DOTALL)
        points = json.loads(match.group()) if match else []
    if not isinstance(points, list):
        return []
    filtered = []
    for pt in points:
        cleaned = clean_icon_tokens(pt)
        if not cleaned:
            continue
        if topic_hint and not _point_matches_topic_hint(cleaned, topic_hint):
            continue
        if _looks_like_duplicate_point(cleaned, existing_points + filtered):
            continue
        filtered.append(cleaned)
    if layout == "two_column" and side_hint == "both":
        filtered = _normalize_two_column_generated_points(filtered)
    target_count = num
    if len(filtered) < target_count and topic_hint:
        extra_prompt = (
            f"Generate {target_count - len(filtered)} additional short bullet(s) about {topic_hint}. "
            f"Avoid repeating these ideas: {json.dumps(existing_points + filtered)}. "
            "Return only a JSON list of strings."
        )
        try:
            extra_resp = _client.chat.completions.create(
                model=AZURE_DEPLOYMENT,
                messages=[{"role": "user", "content": extra_prompt}],
                temperature=0.4,
                max_tokens=220,
            )
            extra_raw = extra_resp.choices[0].message.content.strip()
            extra_points = json.loads(extra_raw) if extra_raw.startswith("[") else json.loads(re.search(r"\[.*\]", extra_raw, re.DOTALL).group()) if re.search(r"\[.*\]", extra_raw, re.DOTALL) else []
            if isinstance(extra_points, list):
                for pt in extra_points:
                    cleaned = clean_icon_tokens(pt)
                    if (
                        cleaned
                        and (not topic_hint or _point_matches_topic_hint(cleaned, topic_hint))
                        and not _looks_like_duplicate_point(cleaned, existing_points + filtered)
                    ):
                        filtered.append(cleaned)
        except Exception:
            pass
    attempts = 0
    while len(filtered) < target_count and attempts < 2:
        remaining = target_count - len(filtered)
        retry_topic_instruction = f"Every bullet must stay specifically about {topic_hint}.\n" if topic_hint else ""
        retry_prompt = (
            f"Generate exactly {remaining} NEW short bullet(s) for this slide.\n"
            f"{existing_text}\n"
            f"Avoid repeating these ideas: {json.dumps(existing_points + filtered)}.\n"
            f"{retry_topic_instruction}"
            "Return only a JSON list of strings."
        )
        try:
            retry_resp = _client.chat.completions.create(
                model=AZURE_DEPLOYMENT,
                messages=[{"role": "user", "content": retry_prompt}],
                temperature=0.6,
                max_tokens=180,
            )
            retry_raw = retry_resp.choices[0].message.content.strip()
            retry_points = json.loads(retry_raw) if retry_raw.startswith("[") else json.loads(re.search(r"\[.*\]", retry_raw, re.DOTALL).group()) if re.search(r"\[.*\]", retry_raw, re.DOTALL) else []
            if isinstance(retry_points, list):
                for pt in retry_points:
                    cleaned = clean_icon_tokens(pt)
                    if (
                        cleaned
                        and (not topic_hint or _point_matches_topic_hint(cleaned, topic_hint))
                        and not _looks_like_duplicate_point(cleaned, existing_points + filtered)
                    ):
                        filtered.append(cleaned)
                        if len(filtered) >= target_count:
                            break
        except Exception:
            pass
        attempts += 1
    if len(filtered) < target_count:
        fallback_base = _safe_str(topic_hint or title or "this slide")
        for i in range(target_count - len(filtered)):
            if layout == "two_column" and side_hint == "both":
                side_label = "left side" if (len(filtered) + i) % 2 == 0 else "right side"
                fallback = _compress_text_for_shortening(
                    f"Additional {side_label} point about {fallback_base} #{i + 1}"
                ) or f"Additional {side_label} point #{i + 1}"
            else:
                fallback = _compress_text_for_shortening(f"Additional point about {fallback_base} #{i + 1}") or f"Additional point {i + 1}"
            if not _looks_like_duplicate_point(fallback, existing_points + filtered):
                filtered.append(fallback)
    if layout == "two_column" and side_hint == "both" and len(filtered) < target_count:
        while len(filtered) < target_count:
            side_label = "left side" if len(filtered) % 2 == 0 else "right side"
            filtered.append(f"Additional {side_label} point {len(filtered) + 1}")
    return filtered[:target_count]

def _append_points_to_slide(slide: dict, points: list[str], side_hint: Optional[str] = None) -> bool:
    if not isinstance(slide, dict):
        return False
    slide.setdefault("_editor_id", uuid4().hex)
    slide.setdefault("layout", "bullets")
    slide.setdefault("content", [])
    layout = _safe_str(slide.get("layout", "bullets"), "bullets") or "bullets"
    cleaned_points = _dedupe_preserve_order(points)
    if not cleaned_points:
        return False
    before = _slide_point_state(slide)

    if layout == "two_column":
        left_points = _dedupe_preserve_order(flatten_slide_content(slide.get("left_points", [])))
        right_points = _dedupe_preserve_order(flatten_slide_content(slide.get("right_points", [])))

        if side_hint == "both":
            for idx, point in enumerate(cleaned_points):
                if idx % 2 == 0:
                    left_points.append(point)
                else:
                    right_points.append(point)
        elif side_hint == "left":
            left_points.extend(cleaned_points)
        elif side_hint == "right":
            right_points.extend(cleaned_points)
        else:
            for idx, point in enumerate(cleaned_points):
                if idx % 2 == 0:
                    left_points.append(point)
                else:
                    right_points.append(point)

        slide["left_points"] = _dedupe_preserve_order(left_points)
        slide["right_points"] = _dedupe_preserve_order(right_points)
        slide["left"] = copy.deepcopy(slide["left_points"])
        slide["right"] = copy.deepcopy(slide["right_points"])
        slide["content"] = _dedupe_preserve_order(flatten_slide_content(slide["left_points"]) + flatten_slide_content(slide["right_points"]))
        slide["layout"] = "two_column"
    elif layout == "timeline":
        steps = [s for s in (slide.get("steps", []) or []) if isinstance(s, dict)]
        start_idx = len(steps)
        for i, point in enumerate(cleaned_points):
            step = _point_text_to_step(point, start_idx + i)
            if not step.get("label") and not step.get("detail"):
                continue
            if step not in steps:
                steps.append(step)
        slide["steps"] = steps
        slide["content"] = _dedupe_preserve_order(flatten_slide_content([f"{s.get('label', '')}: {s.get('detail', '')}".strip(": ").strip() for s in steps]))
    elif layout == "icon_grid":
        items = [g for g in (slide.get("grid_items", []) or []) if isinstance(g, dict)]
        start_idx = len(items)
        for i, point in enumerate(cleaned_points):
            item = _point_text_to_grid_item(point, start_idx + i)
            if item not in items:
                items.append(item)
        slide["grid_items"] = items
        slide["content"] = _dedupe_preserve_order([
            f"{item.get('title', '')}: {item.get('detail', '')}".rstrip(": ").strip()
            for item in items
        ])
    elif layout == "section_index":
        sections = _dedupe_preserve_order(flatten_slide_content(slide.get("sections", slide.get("content", []))))
        sections.extend(cleaned_points)
        slide["sections"] = _dedupe_preserve_order(sections)
        slide["content"] = list(slide["sections"])
    elif layout == "table":
        cols = _dedupe_preserve_order(flatten_slide_content(slide.get("table_columns", [])))
        rows = [list(row) for row in (slide.get("table_rows", []) or []) if isinstance(row, list)]
        if not cols:
            cols = ["Point"]
        for point in cleaned_points:
            row = [point] + [""] * max(0, len(cols) - 1)
            if row not in rows:
                rows.append(row)
        slide["table_columns"] = cols
        slide["table_rows"] = rows
        slide["content"] = _dedupe_preserve_order(flatten_slide_content([cell for row in rows for cell in row]))
    else:
        existing = _dedupe_preserve_order(flatten_slide_content(slide.get("content", [])))
        existing.extend(cleaned_points)
        slide["content"] = _dedupe_preserve_order(existing)

    after = _slide_point_state(slide)
    changed = before != after
    slide["_user_modified"] = bool(changed)
    return changed

def _generate_topic_suggestions(scope: str = "new", current_topic: str = "", history: Optional[list] = None, topic_hint: str = "") -> list[str]:
    history = history or st.session_state.get("ppt_history", [])
    recent_topics = [str(item.get("topic", "")).strip() for item in history[:6] if isinstance(item, dict) and str(item.get("topic", "")).strip()]
    if scope == "current" and current_topic:
        base = current_topic
    elif topic_hint:
        base = topic_hint
    else:
        base = "presentation topics"
    prompt = f"""
You are an expert presentation coach who creates natural, practical, audience-specific presentation topics.

Generate exactly 6 specific, engaging presentation topics.

{"IMPORTANT: ALL topics MUST be strictly within this domain: " + topic_hint if topic_hint else ""}
{"Focus on angles related to: " + base if base != "presentation topics" else ""}

Requirements:
- Every single topic must be directly about: {topic_hint or base}
- Be specific (e.g., "Farm-to-Table Movement: Rebuilding Local Food Ecosystems" NOT "Food Trends")
- Prefer grounded, present-day, real-world topics
- Do NOT introduce futuristic, AI-driven, high-tech, or speculative angles unless the domain itself explicitly asks for them
- Avoid generic words like "overview", "introduction", "roadmap"
- Real-world applicable
- Vary the phrasing style across the 6 topics
- Avoid repetitive openings like "Harnessing", "Leveraging", "Integrating", or "Transforming" unless truly natural

{"Scope: Suggest follow-up/adjacent ideas to: " + current_topic if scope == "current" and current_topic else "Scope: Fresh, standalone topics"}

Avoid repeating: {json.dumps(recent_topics) if recent_topics else "none"}

Return ONLY a raw JSON array of 6 strings. No markdown, no numbering, no explanation.
Example format: ["Topic 1", "Topic 2", "Topic 3", "Topic 4", "Topic 5", "Topic 6"]
"""
    def _parse_topic_items(raw_text: str) -> list[str]:
        raw_text = (raw_text or "").strip()
        if not raw_text:
            return []

        candidates: list[str] = []

        def _add_candidate(line: str) -> None:
            line = clean_icon_tokens(line)
            line = re.sub(r"^\s*(?:[-*•▪‣▶►]+\s*|\d+[.)]\s*)", "", line).strip()
            line = line.strip(' "\'`')
            line = re.sub(r"\s+", " ", line).strip()
            if line and line not in candidates:
                candidates.append(line)

        # Try strict JSON first.
        try:
            parsed = json.loads(raw_text)
            if isinstance(parsed, list):
                for item in parsed:
                    if isinstance(item, str):
                        _add_candidate(item)
                if candidates:
                    return candidates
        except Exception:
            pass

        # Try to recover a JSON-ish array embedded in free text.
        match = re.search(r"\[[\s\S]*\]", raw_text)
        if match:
            try:
                parsed = json.loads(match.group(0))
                if isinstance(parsed, list):
                    for item in parsed:
                        if isinstance(item, str):
                            _add_candidate(item)
                    if candidates:
                        return candidates
            except Exception:
                pass

        # Fall back to one-item-per-line parsing.
        for line in raw_text.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            if re.match(r"^\s*(?:Here are|Topics|Suggestions|Trending|1\.)", stripped, re.IGNORECASE) and ":" in stripped:
                stripped = stripped.split(":", 1)[1].strip()
            _add_candidate(stripped)

        return candidates

    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.8,
            max_tokens=250,
        )
        raw = (resp.choices[0].message.content or "").strip()
        cleaned = _parse_topic_items(raw)
        if cleaned:
            return cleaned[:5]
    except Exception as e:
        print("Topic generation error:", e)
    fallback = [
        "AI Agents in Enterprise Automation",
        "Generative AI in Healthcare Transformation",
        "AI Governance and Regulation Trends",
        "Multimodal AI Applications",
        "AI in Cybersecurity"
    ]
    if scope == "current" and current_topic:
        return [f"{current_topic}: {x}" for x in fallback[:3]]
    return fallback[:5]

# ------------------------------------------------------------------------------
#  ACTION EXECUTORS
# ------------------------------------------------------------------------------
def commit_changes(updated_slides, success_msg):
    normalized_slides = []
    for idx, slide in enumerate(updated_slides or [], start=1):
        if isinstance(slide, dict):
            s = ensure_editor_id(slide)
            if _is_placeholder_slide_title(s.get("title", ""), idx):
                s["title"] = _derive_meaningful_slide_title(s, idx)
            normalized_slides.append(s)
    frozen_slides = _dedupe_thank_you_slides(normalized_slides)
    target_id = st.session_state.pop("action_target_ppt_id", None)
    current_id = target_id or st.session_state.get("current_ppt_id")
    if not current_id:
        return
    current_item = get_ppt_by_id(current_id) or {}
    outline = copy.deepcopy(current_item.get("outline_payload") or st.session_state.get("outline_payload") or {})
    outline["slides"] = copy.deepcopy(frozen_slides)
    st.session_state.outline_payload = copy.deepcopy(outline)
    sync_all_editor_widgets(copy.deepcopy(frozen_slides))
    ppt_bytes = rebuild_ppt_from_outline(outline)
    item = get_ppt_by_id(current_id) if current_id else None
    topic = (item.get("topic") if item else st.session_state.get("topic", "Presentation")) or "Presentation"
    label = f"Updated Preview — {topic} ({current_id})"
    _build_ppt_record(
        current_id,
        outline_payload=outline,
        slides=frozen_slides,
        topic=topic,
        sections=st.session_state.get("sections", ""),
        slide_count=len(frozen_slides),
        ppt_bytes=ppt_bytes,
        ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
    )
    add_message_with_preview(
        role="assistant",
        text=success_msg,
        preview_slides=frozen_slides,
        ppt_label=label,
        ppt_bytes=ppt_bytes,
        ppt_filename=st.session_state.get("ppt_filename", "presentation.pptx"),
    )
    st.rerun()

def _presentation_mode_instruction(sub_intent: Optional[str]) -> str:
    labels = {
        "speaker_notes": "Create concise speaker notes for each slide.",
        "slide_script": "Create a short speaking script for each slide.",
        "audience_questions": "List likely audience questions and brief suggested answers.",
        "speech_summary": "Write a natural 2-minute speech summarizing the whole PPT.",
        "presentation_explain": "Explain the PPT as if the user is presenting it live.",
        "interactive_questions": "Create one interactive audience question per slide.",
    }
    return labels.get(sub_intent or "", "Create practical presenter support for this PPT.")

def _generate_presentation_mode_response(slides: list, sub_intent: Optional[str], user_request: str) -> str:
    if not slides:
        return "No presentation is currently active. Please create or switch to a presentation first."
    compact_slides = summarize_slides_for_llm(slides)
    prompt = f"""
You are a presentation coach.

User request: {user_request}
Task: {_presentation_mode_instruction(sub_intent)}

Slides:
{json.dumps(compact_slides, indent=2)}

Rules:
- Do not add or move slides.
- Do not ask where to place a slide.
- Keep the answer directly usable by a presenter.
- Be concise, specific, and grounded only in the slide content.
- Avoid repetitive openings like "This slide shows" on every slide.
- Use natural transitions between slides and vary sentence structure.
- For whole-PPT explanations, create a flowing presenter talk track with a beginning, bridge phrases, and a closing.
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.35,
            max_tokens=900,
        )
        return (resp.choices[0].message.content or "").strip()
    except Exception as e:
        return f"Presentation coaching failed: {e}"

def _slide_signature(slide: dict) -> str:
    if not isinstance(slide, dict):
        return ""
    bits = [
        _safe_str(slide.get("title", "")),
        _safe_str(slide.get("subtitle", "")),
        " ".join(flatten_slide_content(slide.get("content", []))),
        " ".join(flatten_slide_content(slide.get("left_points", []))),
        " ".join(flatten_slide_content(slide.get("right_points", []))),
    ]
    return _norm_text(" ".join(bits))

def _slide_signature_overlap(a: dict, b: dict) -> float:
    aw = set(_slide_signature(a).split())
    bw = set(_slide_signature(b).split())
    if not aw or not bw:
        return 0.0
    return len(aw & bw) / max(1, len(aw | bw))

def regenerate_slide_variant(slides: list, slide_num: int) -> dict:
    old_slide = slides[slide_num - 1]
    nearby_titles = [_safe_str(s.get("title", ""), "") for s in slides if isinstance(s, dict)]
    old_layout = _safe_str(old_slide.get("layout", "bullets"), "bullets")
    prompt = f"""
Regenerate slide {slide_num} as a genuinely new version while keeping it relevant to the same deck.

Old slide:
{json.dumps(old_slide, indent=2)}

Deck titles:
{json.dumps(nearby_titles)}

Return ONLY one JSON slide object with title, subtitle, layout, icon, content, style.
Rules:
- Do not reuse the same bullet wording.
- Do not reuse the old title unless it is the only accurate title.
- Prefer a different layout than "{old_layout}" when the content still fits.
- Keep the same broad purpose, but change the framing, title, and examples where useful.
- Use 4-6 concise, plain-English bullets.
- No markdown, no nested objects in content.
"""
    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.8,
        max_tokens=650,
    )
    raw = (resp.choices[0].message.content or "").strip()
    if not raw.startswith("{"):
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        raw = m.group() if m else raw
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Invalid regenerated slide JSON")
    data["_user_modified"] = True
    regenerated = normalize_slide(data)
    if _slide_signature(regenerated) == _slide_signature(old_slide) or _slide_signature_overlap(regenerated, old_slide) > 0.72:
        regenerated["title"] = f"Fresh Perspective: {_safe_str(old_slide.get('title', 'Slide'), 'Slide')}"
        regenerated["layout"] = "two_column" if old_layout != "two_column" else "bullets"
        if regenerated["layout"] == "two_column":
            regenerated["left_title"] = "New Angle"
            regenerated["right_title"] = "Audience Takeaway"
            regenerated["left_points"] = [
                f"Reframe the topic around {_safe_str(old_slide.get('title', 'the main idea'), 'the main idea')}",
                "Use a different example or scenario to make the point feel fresh.",
            ]
            regenerated["right_points"] = [
                "Highlight what the audience should notice first.",
                "End with a clearer decision, implication, or takeaway.",
            ]
            regenerated["content"] = []
        else:
            regenerated["content"] = [
                f"Reframed focus: {_safe_str(old_slide.get('title', 'the topic'), 'the topic')}",
                "Updated angle with clearer emphasis for the audience.",
                "New supporting detail to avoid repeating the previous version.",
                "Sharper takeaway that connects this slide to the deck narrative.",
            ]
    if _is_placeholder_slide_title(regenerated.get("title", ""), slide_num):
        regenerated["title"] = _derive_meaningful_slide_title(regenerated, slide_num)
    return regenerated

def execute_action(intent: str, slots: dict, slides: list) -> Tuple[str, bool]:
    if intent == "update_slide":
        intent = "edit_slide"
    action_ppt_id = st.session_state.get("action_target_ppt_id") or st.session_state.get("current_ppt_id")
    if action_ppt_id and action_ppt_id != st.session_state.get("current_ppt_id"):
        switch_active_ppt(action_ppt_id)
    current_outline = _current_outline_payload()
    if current_outline and current_outline.get("slides") is not None:
        slides = copy.deepcopy(current_outline.get("slides", []))

    if intent == "view_slide":
        ppt_ref = slots.get("ppt_ref")
        if ppt_ref:
            explicit_item = _resolve_ppt_history_item(str(ppt_ref))
            best_id = explicit_item.get("id") if explicit_item else None
            if not best_id:
                best_id, clarification = find_ppt_semantically(str(ppt_ref), disambiguate=True)
                if clarification:
                    return clarification, False
            if best_id and best_id != st.session_state.get("current_ppt_id"):
                switch_active_ppt(best_id)
            elif ppt_ref and not best_id:
                return f"I couldn't find {_ppt_display_label(str(ppt_ref))}. Try `ppt 1`, `ppt 2`, or the deck title.", False
        current_outline = _current_outline_payload()
        if not current_outline or not current_outline.get("slides"):
            return "No presentation is currently active. Please create or switch to a presentation first.", False
        slide_num = slots.get("slide_number")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number to view.", False
        slides_local = current_outline.get("slides", [])
        ppt_label = _ppt_display_label(st.session_state.get("current_ppt_id"))
        if not (1 <= slide_num <= len(slides_local)):
            return f"Slide {slide_num} does not exist in {ppt_label}. This presentation only has {len(slides_local)} slide(s).", False
        slide = slides_local[slide_num - 1] if isinstance(slides_local[slide_num - 1], dict) else {}
        return _format_slide_view(slide, slide_num, ppt_label), True

    if intent == "explain_slide":
        slide_num = slots.get("slide_number")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Which slide should I explain?", False
        current_outline = _current_outline_payload()
        slides_local = current_outline.get("slides", []) if current_outline else slides
        if not (1 <= slide_num <= len(slides_local)):
            return f"Slide {slide_num} does not exist. The deck has {len(slides_local)} slide(s).", False
        slide = slides_local[slide_num - 1] if isinstance(slides_local[slide_num - 1], dict) else {}
        compact = _format_slide_view(slide, slide_num, _ppt_display_label(st.session_state.get("current_ppt_id")))
        try:
            resp = _client.chat.completions.create(
                model=AZURE_DEPLOYMENT,
                messages=[{
                    "role": "user",
                    "content": (
                        "Explain this PowerPoint slide in plain language. "
                        "Keep it concise and do not invent facts.\n\n"
                        f"{compact}"
                    ),
                }],
                temperature=0.2,
                max_tokens=260,
            )
            explanation = (resp.choices[0].message.content or "").strip()
        except Exception:
            explanation = compact
        _remember_action_context("explain_slide", slide_num, st.session_state.get("current_ppt_id"), f"Explained slide {slide_num}")
        return explanation, True

    if intent in ("edit_slide", "transform_content", "add_slide", "move_slide", "swap_slides", "merge_slides", "delete_slide", "regenerate_slide", "presentation_mode", "preview_ppt", "download_ppt", "refine_ppt", "design_change"):
        if not st.session_state.get("current_ppt_id"):
            return "No presentation is currently active. Please create or switch to a presentation first.", False
        if not current_outline:
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

        # Extract colors from the full conversation so we don't lose
        # "blue and white theme" after the slide-count follow-up.
        all_messages = st.session_state.get("messages", [])
        theme_colors = extract_theme_colors_from_messages(all_messages) or extract_theme_colors(topic)

        # Remove color phrases from topic to avoid duplication in slide titles
        color_pattern = r'\b(?:in\s+)?(?:dark\s+|light\s+)?(?:' + '|'.join(_COLOR_NAMES) + r')\s*(?:and\s+(?:dark\s+|light\s+)?(?:' + '|'.join(_COLOR_NAMES) + r'))?\s*(?:theme|color|colors|palette|scheme)?\b'
        topic_clean = re.sub(color_pattern, '', topic, flags=re.IGNORECASE).strip().strip(',').strip()
        topic_final = topic_clean if len(topic_clean) > 3 else topic

        generate_outline_and_reply(topic_final, slide_count, st.session_state.tone, sections or None, theme_colors)
        return "Generating presentation...", True

    elif intent == "presentation_mode":
        reply = _generate_presentation_mode_response(
            slides,
            slots.get("sub_intent"),
            slots.get("user_request") or slots.get("change_content") or "",
        )
        _remember_action_context("presentation_mode", None, st.session_state.get("current_ppt_id"), "Generated presenter support")
        return reply, not reply.startswith("Presentation coaching failed:")

    elif intent == "design_change":
        request = slots.get("user_request") or slots.get("change_content") or slots.get("style_hint") or ""
        source_slides = [_shorten_slide_structure(s) for s in slides] if re.search(
            r"\b(use less text|less text|shorter|concise|minimal text)\b",
            request,
            re.IGNORECASE,
        ) else slides
        updated_slides = _apply_design_to_existing_slides(source_slides, request, slots.get("sub_intent") or "")
        if not updated_slides:
            return "No active deck found to restyle.", False
        _remember_action_context("design_change", None, st.session_state.get("current_ppt_id"), "Updated deck design only")
        commit_changes(updated_slides, "Updated the deck design while preserving the existing slide content and structure.")
        _clear_transient_conversation_state()
        return "Updated the deck design while preserving the existing slide content and structure.", True

    elif intent == "regenerate_slide":
        slide_num = slots.get("slide_number")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Which slide should I regenerate?", False
        if not (1 <= slide_num <= len(slides)):
            return f"Slide {slide_num} does not exist. The deck has {len(slides)} slide(s).", False
        try:
            regenerated = regenerate_slide_variant(slides, slide_num)
        except Exception as e:
            return f"Regenerate failed: {e}", False
        updated_slides = copy.deepcopy(slides)
        updated_slides[slide_num - 1] = ensure_editor_id(regenerated)
        updated_slides = refresh_section_index_slide(updated_slides)
        _remember_action_context("regenerate_slide", slide_num, st.session_state.get("current_ppt_id"), f"Regenerated slide {slide_num}")
        msg = _natural_success_message("regenerate_slide", {"slide_number": slide_num, "change_type": "regenerated"})
        commit_changes(updated_slides, msg)
        _clear_transient_conversation_state()
        return msg, True

    elif intent == "move_slide":
        slide_num = slots.get("slide_number")
        anchor_slide = slots.get("anchor_slide")
        relation = str(slots.get("relation") or "").strip().lower()
        position = slots.get("position")
        if slide_num is None and slots.get("source_ref"):
            slide_num = _resolve_slide_reference_text(str(slots.get("source_ref")), slides)
        if slide_num is None:
            return "Which slide would you like to move?", False
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number to move.", False
        if not (1 <= slide_num <= len(slides)):
            return f"Slide {slide_num} does not exist. The deck has {len(slides)} slide(s).", False

        pos_lower = str(position).strip().lower() if position is not None else ""
        if position is not None and (relation == "position" or pos_lower in ("start", "first", "top", "end", "last", "bottom") or pos_lower.isdigit()):
            try:
                if pos_lower in ("start", "first", "top"):
                    position = 1
                elif pos_lower in ("end", "last", "bottom"):
                    position = len(slides)
                else:
                    position = int(position)
            except (ValueError, TypeError):
                return "Please provide a valid destination position.", False
            updated_slides = _move_slide_in_list(slides, slide_num, "position", position=position)
            updated_slides = _dedupe_thank_you_slides(updated_slides)
            updated_slides = refresh_section_index_slide(updated_slides)
            _remember_action_context("move_slide", position, st.session_state.get("current_ppt_id"), f"Moved slide {slide_num} to position {position}")
            msg = _natural_success_message("move_slide", {"slide_number": slide_num})
            commit_changes(updated_slides, msg)
            _clear_transient_conversation_state()
            return msg, True

        if anchor_slide is None and slots.get("anchor_ref"):
            anchor_slide = _resolve_slide_reference_text(str(slots.get("anchor_ref")), slides)
        if anchor_slide is None:
            return f"Where should I move slide {slide_num} in {_ppt_display_label(st.session_state.get('current_ppt_id'))}?", False
        try:
            anchor_slide = int(anchor_slide)
        except (ValueError, TypeError):
            return "Please provide a valid destination slide number.", False
        if not (1 <= anchor_slide <= len(slides)):
            return f"Destination slide {anchor_slide} does not exist. The deck has {len(slides)} slide(s).", False
        if relation not in ("before", "after"):
            relation = "before"
        updated_slides = _move_slide_in_list(slides, slide_num, relation, anchor_idx=anchor_slide)
        updated_slides = _dedupe_thank_you_slides(updated_slides)
        updated_slides = refresh_section_index_slide(updated_slides)
        _remember_action_context("move_slide", slide_num, st.session_state.get("current_ppt_id"), f"Moved slide {slide_num} {relation} slide {anchor_slide}")
        msg = _natural_success_message("move_slide", {"slide_number": slide_num})
        commit_changes(updated_slides, msg)
        _clear_transient_conversation_state()
        return msg, True

    elif intent == "swap_slides":
        slide_a = slots.get("slide_a")
        slide_b = slots.get("slide_b")
        if slide_a is None and slots.get("slide_a_ref"):
            slide_a = _resolve_slide_reference_text(str(slots.get("slide_a_ref")), slides)
        if slide_b is None and slots.get("slide_b_ref"):
            slide_b = _resolve_slide_reference_text(str(slots.get("slide_b_ref")), slides)
        try:
            slide_a = int(slide_a)
            slide_b = int(slide_b)
        except (ValueError, TypeError):
            return "Please provide two valid slide numbers to swap.", False
        if not (1 <= slide_a <= len(slides)) or not (1 <= slide_b <= len(slides)):
            return f"Both slides must exist in the deck. The deck has {len(slides)} slide(s).", False
        if slide_a == slide_b:
            return "Those two slides are already in the same position.", False
        updated_slides = list(copy.deepcopy(slides))
        updated_slides[slide_a - 1], updated_slides[slide_b - 1] = updated_slides[slide_b - 1], updated_slides[slide_a - 1]
        updated_slides = _dedupe_thank_you_slides(updated_slides)
        updated_slides = refresh_section_index_slide(updated_slides)
        _remember_action_context("swap_slides", slide_a, st.session_state.get("current_ppt_id"), f"Swapped slide {slide_a} with slide {slide_b}")
        msg = _natural_success_message("swap_slides", {"slide_number": slide_a})
        commit_changes(updated_slides, msg)
        _clear_transient_conversation_state()
        return msg, True

    elif intent == "merge_slides":
        slide_a = slots.get("slide_a")
        slide_b = slots.get("slide_b")
        slide_ids = slots.get("slide_ids") or []
        if (slide_a is None or slide_b is None) and len(slide_ids) >= 2:
            slide_a, slide_b = slide_ids[:2]
        try:
            slide_a = int(slide_a)
            slide_b = int(slide_b)
        except (ValueError, TypeError):
            return "Which two slides should I merge?", False
        if slide_a == slide_b:
            return "Please choose two different slides to merge.", False
        if not (1 <= slide_a <= len(slides)) or not (1 <= slide_b <= len(slides)):
            return f"Both slides must exist in the deck. The deck has {len(slides)} slide(s).", False
        first_idx, second_idx = sorted([slide_a, slide_b])
        updated_slides = copy.deepcopy(slides)
        primary = updated_slides[first_idx - 1]
        secondary = updated_slides[second_idx - 1]
        if not isinstance(primary, dict) or not isinstance(secondary, dict):
            return "I could not merge those slides because one slide is unavailable.", False
        topic_hint = _safe_str(st.session_state.get("topic", ""), "Presentation")
        updated_slides[first_idx - 1] = merge_slide_payloads(primary, secondary, topic=topic_hint)
        del updated_slides[second_idx - 1]
        updated_slides = refresh_section_index_slide(updated_slides)
        _remember_action_context("merge_slides", first_idx, st.session_state.get("current_ppt_id"), f"Merged slide {second_idx} into slide {first_idx}")
        msg = _natural_success_message("merge_slides", {"slide_number": first_idx})
        commit_changes(updated_slides, msg)
        _clear_transient_conversation_state()
        return msg, True

    elif intent in ("clear_slide", "blank_slide"):
        slide_num = slots.get("slide_number")
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number.", False
        if not (1 <= slide_num <= len(slides)):
            return f"Slide {slide_num} doesn't exist. Deck has {len(slides)} slides.", False
        updated_slide = _clear_slide_payload(slides[slide_num - 1])
        updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
        _remember_action_context(intent, slide_num, st.session_state.get("current_ppt_id"), f"Cleared slide {slide_num}")
        msg = _natural_success_message(intent, {"slide_number": slide_num})
        commit_changes(updated_slides, msg)
        _clear_transient_conversation_state()
        return msg, True

    elif intent == "edit_slide":
        slide_num = slots.get("slide_number")
        change_content = slots.get("change_content")
        if slide_num is None:
            slide_num = st.session_state.get("active_slide_index") or st.session_state.get("last_slide_index")
        # Always let an explicit slide reference in the latest message override
        # any stale slide number that may have been carried over from context.
        if change_content:
            explicit_slide_num = _extract_explicit_slide_number_from_text(change_content)
            if explicit_slide_num is not None:
                slide_num = explicit_slide_num
                slots["slide_number"] = slide_num
        # If slide_number is still missing, try a final recovery from change_content.
        if slide_num is None and change_content:
            fallback_slide_num = _extract_explicit_slide_number_from_text(change_content)
            if fallback_slide_num is not None:
                slide_num = fallback_slide_num
                slots["slide_number"] = slide_num
        try:
            slide_num = int(slide_num)
        except (ValueError, TypeError):
            return "Please provide a valid slide number.", False
        if not change_content:
            return "Missing change content for editing.", False
        if slide_num < 1 or slide_num > len(slides):
            return f"Slide {slide_num} doesn't exist. Deck has {len(slides)} slides.", False
        target_slide = slides[slide_num - 1]

        explicit_content = _extract_explicit_provided_content(change_content)
        if explicit_content:
            updated_slide = copy.deepcopy(target_slide)
            replacement = _slide_from_explicit_content(explicit_content, _safe_str(updated_slide.get("title", "User Provided Content"), "User Provided Content"))
            updated_slide["layout"] = replacement.get("layout", "bullets")
            updated_slide["content"] = replacement.get("content", [])
            updated_slide["subtitle"] = replacement.get("subtitle", updated_slide.get("subtitle", ""))
            if replacement.get("title") and replacement.get("title") != "User Provided Content":
                updated_slide["title"] = replacement.get("title")
            for key in ("left_points", "right_points", "left", "right", "steps", "grid_items", "table_columns", "table_rows"):
                updated_slide.pop(key, None)
            updated_slide["_user_modified"] = True
            updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
            _remember_action_context("edit_slide", slide_num, st.session_state.get("current_ppt_id"), f"Inserted provided content on slide {slide_num}")
            msg = _natural_success_message("edit_slide", {"slide_number": slide_num, "change_type": "inserted provided content", "change_content": change_content})
            commit_changes(updated_slides, msg)
            _clear_transient_conversation_state()
            return msg, True

        layout_change_match = re.search(
            r"\b(?:change|convert|switch|set)\b.*\blayout\b.*\b(?:to|as)\b.*\b(?:bullet|bullets|bullet points?)\b"
            r"|\b(?:convert|change)\b.*\bslide\b.*\b(?:to|into)\b.*\bbullet points?\b",
            change_content or "",
            re.IGNORECASE,
        )
        if layout_change_match:
            updated_slide = copy.deepcopy(target_slide)
            previous_layout = _safe_str(updated_slide.get("layout", "")).lower()
            combined_points = flatten_slide_content(updated_slide.get("content", []))
            if previous_layout == "two_column":
                combined_points = (
                    flatten_slide_content(updated_slide.get("left_points", []))
                    + flatten_slide_content(updated_slide.get("right_points", []))
                )
            if not combined_points:
                combined_points = (
                    flatten_slide_content(updated_slide.get("left_points", []))
                    + flatten_slide_content(updated_slide.get("right_points", []))
                )
            updated_slide["layout"] = "bullets"
            updated_slide["content"] = flatten_slide_content(combined_points)
            if not updated_slide["content"]:
                updated_slide["content"] = flatten_slide_content(target_slide.get("content", []))
            for key in ("left_points", "right_points", "left_title", "right_title", "left", "right", "steps", "grid_items"):
                if key in updated_slide:
                    updated_slide.pop(key, None)
            updated_slide["_user_modified"] = True
            updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
            _remember_action_context("edit_slide", slide_num, st.session_state.get("current_ppt_id"), f"Changed slide {slide_num} layout to bullets")
            msg = _natural_success_message("edit_slide", {"slide_number": slide_num, "change_type": "layout_change", "change_content": change_content})
            commit_changes(updated_slides, msg)
            _clear_transient_conversation_state()
            return msg, True

        remove_request = _extract_remove_instruction(change_content)
        if remove_request:
            if target_slide.get("layout") == "two_column" and not remove_request.get("side_hint"):
                return "Do you want to remove from left, right, or both?", False
            if _remove_points_from_slide(target_slide, remove_request):
                _sync_two_column_aliases(target_slide)
                updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, target_slide))
                _remember_action_context("edit_slide", slide_num, st.session_state.get("current_ppt_id"), f"Updated slide {slide_num}")
                msg = _natural_success_message("edit_slide", {"slide_number": slide_num, "change_content": change_content})
                commit_changes(updated_slides, msg)
                _clear_transient_conversation_state()
                return msg, True
            return "I couldn't find a matching point to remove.", False

        if _needs_layout_aware_update(change_content, target_slide):
            updated_slide = copy.deepcopy(target_slide)
            if _apply_layout_aware_update(updated_slide, change_content):
                _sync_two_column_aliases(updated_slide)
                updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
                _remember_action_context("edit_slide", slide_num, st.session_state.get("current_ppt_id"), f"Applied structured update on slide {slide_num}")
                msg = _natural_success_message("edit_slide", {"slide_number": slide_num, "change_type": "structured_update", "change_content": change_content})
                commit_changes(updated_slides, msg)
                _clear_transient_conversation_state()
                return msg, True

        if _is_transform_request(change_content) or intent == "transform_content":
            original_slide = copy.deepcopy(target_slide)
            try:
                result, usage = transform_slide_with_llm(f"Transform slide {slide_num}: {change_content}", slides, slide_num)
                record_usage("transform", usage)
            except Exception as e:
                return f"Transform failed: {e}", False
            if result.get("action") == "edit":
                raw_slides = result.get("slides", [])
                for s in raw_slides:
                    if "content" in s and isinstance(s["content"], list):
                        s["content"] = [clean_icon_tokens(item) for item in s["content"] if clean_icon_tokens(item)]
                    if "left_points" in s:
                        s["left_points"] = [clean_icon_tokens(p) for p in s["left_points"] if clean_icon_tokens(p)]
                    if "right_points" in s:
                        s["right_points"] = [clean_icon_tokens(p) for p in s["right_points"] if clean_icon_tokens(p)]
                cleaned = [ensure_editor_id(s) for s in raw_slides]
                if 1 <= slide_num <= len(cleaned):
                    cleaned[slide_num - 1]["_user_modified"] = True
                updated_slide = cleaned[slide_num - 1] if 1 <= slide_num <= len(cleaned) else target_slide
                _sync_two_column_aliases(updated_slide)
                updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
                _remember_action_context("transform_content", slide_num, st.session_state.get("current_ppt_id"), f"Transformed slide {slide_num}")
                msg = _natural_success_message("transform_content", {"slide_number": slide_num, "change_content": change_content})
                commit_changes(updated_slides, msg)
                _clear_transient_conversation_state()
                return msg, True
            return "Transform failed. Please try again.", False

        # SPECIAL HANDLE: "add point(s)" - generate points directly while preserving the existing layout.
        num = slots.get("n_points")
        if num is None:
            num = _parse_bullet_add_request(change_content)
        side_hint = _resolve_add_points_side_hint(change_content, target_slide)
        topic_hint = _extract_topic_constraint_from_text(change_content)
        if num is not None:
            num = _resolved_add_point_count(num, side_hint)
            slide = target_slide
            try:
                new_points = _generate_add_points(slide, num, side_hint=side_hint, topic_hint=topic_hint)
            except Exception as e:
                return f"Failed to generate points: {e}", False
            if new_points:
                if not _append_points_to_slide(slide, new_points, side_hint=side_hint):
                    return "No new point was added because the slide content did not change.", False
                _sync_two_column_aliases(slide)
                updated_slides = refresh_section_index_slide(slides)
                _remember_action_context(
                    "add_points",
                    slide_num,
                    st.session_state.get("current_ppt_id"),
                    f"Added {len(new_points)} point(s) to slide {slide_num}",
                )
                slide["_user_modified"] = True
                msg = _natural_success_message("add_points", {"slide_number": slide_num, "n_points": len(new_points), "change_type": "added points", "change_content": change_content})
                commit_changes(updated_slides, msg)
                _clear_transient_conversation_state()
                return msg, True
            return "Failed to generate points. Please try again.", False

        # For all other edits (remove, change, etc.), use the LLM
        edit_prompt = f"Edit slide {slide_num}: {change_content}"
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
                if 1 <= slide_num <= len(cleaned):
                    cleaned[slide_num - 1]["_user_modified"] = True
                updated_slide = cleaned[slide_num - 1]
                if _is_shortening_request(change_content) and _slide_text_size(updated_slide) >= max(1, int(_slide_text_size(original_slide) * 0.9)):
                    updated_slide = _shorten_slide_structure(updated_slide)
                    cleaned[slide_num - 1] = updated_slide
                _sync_two_column_aliases(updated_slide)
                updated_slides = refresh_section_index_slide(_merge_updated_slide(slides, slide_num, updated_slide))
                _remember_action_context("edit_slide", slide_num, st.session_state.get("current_ppt_id"), f"Updated slide {slide_num}")
                msg = _natural_success_message("edit_slide", {"slide_number": slide_num, "change_content": change_content})
                commit_changes(updated_slides, msg)
                _clear_transient_conversation_state()
                return msg, True
            else:
                num = slots.get("n_points")
                if num is None:
                    num = _parse_bullet_add_request(change_content)
                side_hint = _resolve_add_points_side_hint(change_content, target_slide)
                if num is not None:
                    num = _resolved_add_point_count(num, side_hint)
                    slide = target_slide
                    try:
                        new_points = _generate_add_points(slide, num, side_hint=side_hint, topic_hint=topic_hint)
                    except Exception as e:
                        return f"Failed to generate points: {e}", False
                    if new_points:
                        if not _append_points_to_slide(slide, new_points, side_hint=side_hint):
                            return "No new point was added because the slide content did not change.", False
                        _sync_two_column_aliases(slide)
                        updated_slides = refresh_section_index_slide(slides)
                        _remember_action_context(
                            "add_points",
                            slide_num,
                            st.session_state.get("current_ppt_id"),
                            f"Added {len(new_points)} point(s) to slide {slide_num}",
                        )
                        slide["_user_modified"] = True
                        msg = _natural_success_message("add_points", {"slide_number": slide_num, "n_points": len(new_points), "change_type": "added points", "change_content": change_content})
                        commit_changes(updated_slides, msg)
                        _clear_transient_conversation_state()
                        return msg, True
                return result.get("question", "Edit failed. Please rephrase."), False
        except Exception as e:
            return f"Edit failed: {e}", False

    elif intent == "add_slide":
        slide_content = slots.get("slide_content")
        position = slots.get("position")
        if not slide_content:
            return "Missing slide content.", False
        is_thank_you = _is_thank_you_slide_request(slide_content)
        if is_thank_you and not position:
            position = "end"
        if not position:
            return f"Where should I add the slide about '{slide_content}' in {_ppt_display_label(st.session_state.get('current_ppt_id'))}?", False
        try:
            before_count = len(slides)
            explicit_content = slots.get("explicit_user_content") or _extract_explicit_provided_content(slide_content)
            if explicit_content:
                new_slide_data = _slide_from_explicit_content(explicit_content, "User Provided Content")
            elif is_thank_you:
                new_slide_data = make_thank_you_slide({})
            else:
                new_slide_data = draft_slide_from_request(slide_content, slides)
            new_slide_data = normalize_slide(new_slide_data)
            new_title = _norm_text(new_slide_data.get("title", ""))
            last_title = _norm_text((slides[-1] if slides else {}).get("title", "")) if slides else ""
            duplicate_markers = {"thank you", "thanks", "conclusion", "q a", "q and a", "questions"}
            if new_title and last_title and new_title == last_title and new_title in duplicate_markers:
                return f"'{new_slide_data.get('title', 'This slide')}' already appears at the end of the deck.", False
            updated_slides = list(slides)
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
            if is_thank_you:
                updated_slides = [s for s in updated_slides if not _slide_looks_like_thank_you(s)]
                idx = len(updated_slides) + 1
            updated_slides.insert(idx - 1, new_slide_data)
            updated_slides = _dedupe_thank_you_slides(updated_slides)
            updated_slides = refresh_section_index_slide(updated_slides)
            if len(updated_slides) <= before_count:
                return "New slide was not added. Please try again.", False
            _remember_action_context("add_slide", idx, st.session_state.get("current_ppt_id"), f"Added new slide at position {idx}")
            msg = _natural_success_message("add_slide", {"slide_number": idx, "content": slide_content})
            commit_changes(updated_slides, msg)
            _clear_transient_conversation_state()
            return msg, True
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
            _remember_action_context("delete_slide", slide_num, st.session_state.get("current_ppt_id"), f"Removed slide {slide_num}")
            msg = _natural_success_message("delete_slide", {"slide_number": slide_num})
            commit_changes(updated_slides, msg)
            _clear_transient_conversation_state()
            return msg, True
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
                    preview_slides = copy.deepcopy(item["outline_payload"].get("slides", []))
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
            current_item = get_ppt_by_id(st.session_state.get("current_ppt_id")) or {}
            outline = copy.deepcopy(current_item.get("outline_payload") or st.session_state.get("outline_payload"))
            if outline and outline.get("slides"):
                preview_slides = copy.deepcopy(outline["slides"])
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
            lines = ["Here are your generated presentations:"]
            for i, p in enumerate(history, 1):
                topic = p.get("topic", f"PPT {i}")
                ppt_id = p.get("id", f"ppt_{i}")
                lines.append(f"ppt {i}: {topic}")
            return "\n".join(lines), True
        elif info_type == "topic":
            ppt_ref = slots.get("ppt_ref")
            if ppt_ref:
                item = _resolve_ppt_history_item(ppt_ref)
                if item:
                    topic = item.get("topic", "").strip() or "Untitled presentation"
                    ppt_id = item.get("id", "ppt")
                    label = _ppt_display_label(ppt_id)
                    return f"The topic for {label} is: {topic}.", True
                return "I couldn't find that presentation. Try 'ppt 1', 'ppt 2', or the presentation title.", False
            current_id = st.session_state.get("current_ppt_id")
            if current_id:
                item = get_ppt_by_id(current_id)
                if item:
                    topic = item.get("topic", "").strip() or "Untitled presentation"
                    label = _ppt_display_label(current_id)
                    return f"The topic for {label} is: {topic}.", True
            return "Please specify which PPT you want the topic for.", False
        else:
            return "I can tell you how many PPTs you have, list them, or give the topic of a specific PPT. What would you like?", False

    elif intent == "suggest_topic":
        scope = slots.get("scope")
        current_topic = st.session_state.get("topic", "")
        topic_hint = _safe_str(slots.get("topic", "")) or _extract_topic_constraint_from_text(current_topic)
        suggestions = _generate_topic_suggestions(scope or "new", current_topic=current_topic, topic_hint=topic_hint)
        _set_active_topic_domain(topic_hint or current_topic)
        _remember_action_context("suggest_topic", None, st.session_state.get("current_ppt_id"), f"Suggested topics about {topic_hint or current_topic}")
        return "Here are a few topic ideas:\n\n" + "\n".join(f"- {s}" for s in suggestions), True

    elif intent == "refine_ppt":
        current_id = st.session_state.get("current_ppt_id")
        if not current_id or not _current_outline_payload():
            return "No presentation is currently active. Please open a PPT first.", False
        request = " ".join([
            _safe_str(slots.get("style_hint", ""), ""),
            _safe_str(slots.get("change_content", ""), ""),
            _safe_str(slots.get("user_request", ""), ""),
        ])
        if re.search(r"\b(shorter|less text|concise|condense|simpler|remove fluff|key points only|important points only)\b", request, re.IGNORECASE):
            updated_slides = [_shorten_slide_structure(s) for s in slides]
            commit_changes(updated_slides, "Tightened the deck content while preserving the existing slide structure.")
            _clear_transient_conversation_state()
            return "Tightened the deck content while preserving the existing slide structure.", True
        updated_slides = _apply_design_to_existing_slides(slides, request, slots.get("sub_intent") or "")
        commit_changes(updated_slides, "Refined the deck styling while preserving existing slide content and structure.")
        _clear_transient_conversation_state()
        return "Refined the deck styling while preserving existing slide content and structure.", True

    elif intent == "greeting":
        return "👋 Hello! I'm your PPT assistant. You can ask me to create a new presentation, edit slides, add content, or switch between decks. What would you like to do?", True
    elif intent == "smalltalk":
        return "I'm here when you're ready to work on your presentation. Just tell me what you'd like to do.", True
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
    shorter_request = _is_shortening_request(user_text)
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
7. Content fields must be plain string arrays - no nested dicts.
8. Preserve each slide's existing layout and structured fields. Do not convert a structured slide into bullets unless the user explicitly asks for a layout change.
9. For two_column slides, keep left_points and right_points intact and update them directly instead of flattening into content.
10. If the request asks for a shorter or tighter version, reduce the amount of text materially, remove redundancy, and prefer fewer bullets or shorter sentences.
{ "11. Aim for roughly 30-40% less text on shorten requests." if shorter_request else "" }

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

def transform_slide_with_llm(user_text, slides, slide_num: int):
    shorter_request = _is_shortening_request(user_text)
    prompt = f"""
You are an AI presentation editor. Rewrite ONLY slide {slide_num} to improve clarity, professionalism, and wording.

User request: {user_text}

Current slides (FULL DECK):
{json.dumps(slides, indent=2)}

RULES:
1. Modify only slide {slide_num}.
2. Rewrite text to be clearer and more professional.
3. Keep the same slide structure and layout unless absolutely necessary.
4. Do NOT append duplicate content.
5. Do NOT leave the content unchanged.
6. Preserve meaning while improving phrasing.
7. Return the ENTIRE updated slides array with all other slides unchanged.
8. If the request asks for a shorter version, materially reduce the length by trimming redundant wording, cutting low-value bullets, and aiming for noticeably fewer words.
{ "9. Target about 30-40% less text than the original slide content." if shorter_request else "" }

Return ONLY valid JSON:
{{
  "action": "edit",
  "slides": [ ... full updated slides array ... ]
}}
"""
    response = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.5,
    )
    usage = usage_to_dict(response.usage)
    raw = response.choices[0].message.content
    try:
        data = extract_first_json(raw)
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
    if re.search(r"\b(comparison|compare|vs|versus|comparison\s+table)\b", user_text or "", re.IGNORECASE):
        left, right = _extract_comparison_subjects(user_text, None)
        return {
            "title": f"{left} vs {right}",
            "subtitle": "Comparison overview",
            "layout": "table",
            "icon": "▸",
            "content": [],
            "table_columns": ["Aspect", left, right],
            "table_rows": _build_comparison_table_rows(left, right, user_text),
            "style": {},
        }
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
        lower = str(sections).lower()
        if re.search(r"\b(shorter|less text|concise|clean|clearer|simpler|important points only|key points only|remove fluff)\b", lower):
            return (
                f"{base}. Refine the existing deck to be shorter and tighter. "
                f"Use fewer bullets, remove redundancy, and keep only the core ideas. "
                f"Refinement notes: {sections}."
            )
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
    for idx, slide in enumerate(slides, start=1):
        if isinstance(slide, dict) and _is_placeholder_slide_title(slide.get("title", ""), idx):
            slide["title"] = _derive_meaningful_slide_title(slide, idx)
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
        slides.extend([ensure_editor_id(s) for s in extra_slides])
        attempts += 1
    trimmed = slides[:required_count]
    for idx, slide in enumerate(trimmed, start=1):
        if isinstance(slide, dict) and _is_placeholder_slide_title(slide.get("title", ""), idx):
            slide["title"] = _derive_meaningful_slide_title(slide, idx)
    return trimmed

def generate_outline_and_reply(topic: str, count: int, tone: str, sections=None, theme_colors=None, ppt_id: Optional[str] = None):
    with st.spinner(f"✍️ Building {count}-slide deck on '{topic}'…"):
        try:
            payload = request_outline(topic, count, tone, sections, theme_colors)
            payload["slides"] = payload.get("slides", [])[:count]
            payload["slides"] = _dedupe_thank_you_slides(payload.get("slides", []))
            st.session_state.outline_payload = payload
            st.session_state.topic = topic
            existing_id = ppt_id
            is_update = bool(existing_id and get_ppt_by_id(existing_id))
            deck_id = existing_id if is_update else next_ppt_id()
            _set_conv_state(
                last_topic=topic,
                active_slide_index=None,
                last_action_type="refine_ppt" if is_update else "create_ppt",
                pending_create=None,
                pending_edit=None,
                turn_summary=f"{'Updated' if is_update else 'Created'} {count}-slide deck on {topic}",
            )
            st.session_state.sections = sections or ""
            st.session_state.num_slides = count
            st.session_state.slide_count = count
            st.session_state.current_ppt_id = deck_id
            st.session_state.active_topic_domain = topic
            _remember_action_context("refine_ppt" if is_update else "create_ppt", None, deck_id, f"{'Updated' if is_update else 'Created'} {count}-slide deck on {topic}")
            _set_conv_state(active_ppt_id=deck_id, last_topic=topic)
            _build_ppt_record(
                deck_id,
                topic=topic,
                slide_count=count,
                sections=sections or "",
                outline_payload=payload,
                slides=payload.get("slides", []),
                ppt_bytes=None,
                ppt_filename=None,
                file_path=None,
            )
            sync_all_editor_widgets(payload.get("slides", []))
            ppt_bytes = rebuild_ppt_from_outline()
            reply_text = summarize_outline(payload.get("slides", []))
            if is_update:
                reply_text = f"✅ Updated the deck on {topic}.\n\n{reply_text}"
            slides_for_preview = payload.get("slides", [])
            label = f"Preview — {topic} ({deck_id})"
            _clear_transient_conversation_state()
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

def rebuild_ppt_from_outline(outline_payload: Optional[dict] = None) -> bytes | None:
    current_id = st.session_state.get("current_ppt_id")
    current_item = get_ppt_by_id(current_id) if current_id else None
    outline = copy.deepcopy(
        outline_payload
        or st.session_state.get("outline_payload")
        or (current_item or {}).get("outline_payload")
        or {}
    )
    if not outline or not outline.get("slides"):
        return None
    outline["slides"] = _dedupe_thank_you_slides(outline.get("slides", []))
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
                slides = outline.get("slides", [])
                _build_ppt_record(
                    ppt_id,
                    ppt_bytes=ppt_bytes,
                    ppt_filename=filename,
                    file_path=None,
                    outline_payload=outline,
                    topic=st.session_state.topic,
                    sections=st.session_state.sections,
                    slides=slides,
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
    slides_source = payload.get("slides", []) if isinstance(payload, dict) else []
    safe = {
        "design_system": payload.get("design_system", {}) if isinstance(payload, dict) else {},
        "slides": [],
    }
    for slide in _dedupe_thank_you_slides(slides_source):
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
        if slide.get("_user_modified"):
            s["_user_modified"] = True
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
st.title("💬 AI PPT Generator")
st.caption("Describe your deck, then edit it in chat and download.")
st.caption(f"Active context: {_build_context_badge()}")

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
            logo_notice = f"Logo uploaded: **{uploaded_logo.name}**"
            with st.chat_message("assistant"):
                st.markdown(logo_notice)
            add_message("assistant", logo_notice)
        if _current_outline_payload():
            rebuild_ppt_from_outline()

if prompt is not None:
    prompt = prompt.strip()
    if not prompt:
        st.stop()

    early_pending_state = st.session_state.get("pending_action") or st.session_state.get("pending_intent") or {}

    # If this is a create_ppt request and no file was uploaded, clear any leftover logo
    if (files is None or len(files) == 0) and re.search(r'\b(?:make|create|generate|new)\s+(?:a\s+)?(?:ppt|presentation|deck)', prompt, re.IGNORECASE):
        st.session_state.logo_bytes = None
        st.session_state.logo_name = None

    # Render user message immediately
    with st.chat_message("user"):
        st.markdown(prompt)
    add_message("user", prompt)
    with st.chat_message("assistant"):
        st.markdown(
            """
            <div style="display:flex;align-items:center;gap:.55rem;color:#6b7280;font-size:.95rem;margin:.2rem 0 .6rem;">
              <span style="width:14px;height:14px;border:2px solid #d1d5db;border-top-color:#4f46e5;border-radius:50%;display:inline-block;animation:pptSpin .8s linear infinite;"></span>
              <span>Thinking...</span>
            </div>
            <style>@keyframes pptSpin{to{transform:rotate(360deg)}}</style>
            """,
            unsafe_allow_html=True,
        )
    st.session_state.action_target_ppt_id = None

    if _is_cancel_command(prompt):
        _clear_pending_turn_state()
        st.session_state.active_edit_context = None
        reply = "Okay, I cancelled the current request."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    if _is_null_add_request(prompt):
        reply = "Nothing to add."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    reference_outline = _current_outline_payload() or {}
    reference_slides = reference_outline.get("slides", []) if isinstance(reference_outline, dict) else []
    if _handle_logo_insertion_request(prompt, reference_slides):
        st.stop()
    if _handle_reference_execution_request(prompt, reference_slides):
        st.stop()

    open_pending_create = st.session_state.get("pending_new_ppt") or {}
    skip_open_router = bool(
        open_pending_create.get("topic")
        and not open_pending_create.get("need_topic")
        and _extract_slide_count_from_text(prompt)
    )
    if not skip_open_router:
        open_outline = _current_outline_payload() or {}
        open_slides = open_outline.get("slides", []) if isinstance(open_outline, dict) else []
        open_state = build_assistant_state(open_slides)
        open_decision = resolve_open_conversation(
            prompt,
            open_state,
            pending_need_topic=bool(open_pending_create.get("need_topic") or (open_pending_create and not open_pending_create.get("topic"))),
        )
        if open_decision and open_decision.intent == "create_ppt" and open_decision.mode == "ask":
            open_topic = _safe_str((open_decision.slots or {}).get("topic"), "")
            open_slide_count = (open_decision.slots or {}).get("slide_count") or _extract_slide_count_from_text(prompt)
            if not open_topic:
                _ask_for_create_topic(prompt, question_override=open_decision.message)
            else:
                sections = _extract_create_sections_from_prompt(prompt)
                theme_colors = extract_theme_colors_from_messages(st.session_state.get("messages", [])) or extract_theme_colors(open_topic)
                if open_slide_count:
                    generate_outline_and_reply(open_topic, int(open_slide_count), st.session_state.tone, sections, theme_colors)
                    st.stop()
                st.session_state.pending_new_ppt = {
                    "topic": open_topic,
                    "sections": sections,
                    "theme_colors": theme_colors,
                    "need_topic": False,
                }
                question = open_decision.message or f"Great - how many slides should the presentation on {open_topic} have?"
                st.session_state.pending_intent = {
                    "intent": "create_ppt",
                    "slots": {"topic": open_topic, "slide_count": None, "sections": sections},
                    "missing_slots": ["slide_count"],
                    "next_question": question,
                    "action": "ask",
                }
                st.session_state.pending_action = st.session_state.pending_intent
                _set_conv_state(pending_create=st.session_state.pending_new_ppt)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
        if open_decision and open_decision.mode == "answer":
            if open_decision.intent == "readonly_advice":
                reply = generate_conversational_answer(prompt, open_state, open_slides)
                _remember_answer_objects(prompt, reply, _readonly_target_slide(prompt, open_state, open_slides))
            else:
                reply = open_decision.message or generate_conversational_answer(prompt, open_state, open_slides)
                _remember_answer_objects(prompt, reply, None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()

    pre_route = pre_route_message(
        prompt,
        explicit_topic=_extract_create_ppt_topic(prompt),
        pending_need_topic=bool((st.session_state.get("pending_new_ppt") or {}).get("need_topic")),
    )
    if pre_route and pre_route.intent == "create_ppt":
        implicit_topic = pre_route.slots.get("topic")
        if not implicit_topic:
            _ask_for_create_topic(prompt)
        else:
            slide_count = _extract_slide_count_from_text(prompt)
            sections = _extract_create_sections_from_prompt(prompt)
            theme_colors = extract_theme_colors_from_messages(st.session_state.get("messages", [])) or extract_theme_colors(implicit_topic)
            if slide_count:
                generate_outline_and_reply(implicit_topic, slide_count, st.session_state.tone, sections, theme_colors)
                st.stop()
            st.session_state.pending_new_ppt = {
                "topic": implicit_topic,
                "sections": sections,
                "theme_colors": theme_colors,
                "need_topic": False,
            }
            st.session_state.pending_intent = {
                "intent": "create_ppt",
                "slots": {"topic": implicit_topic, "slide_count": None, "sections": sections},
                "missing_slots": ["slide_count"],
                "next_question": pre_route.message or f"Great - how many slides should the presentation on {implicit_topic} have?",
                "action": "ask",
            }
            st.session_state.pending_action = st.session_state.pending_intent
            _set_conv_state(pending_create=st.session_state.pending_new_ppt)
            question = pre_route.message or f"Great - how many slides should the presentation on {implicit_topic} have?"
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()

    if (
        early_pending_state.get("intent") in {"confirm_topic_switch", "confirm_add_points_limit", "clear_slide"}
        and (_is_affirmative_command(prompt) or _is_negative_command(prompt))
    ):
        routing_prompt = prompt
    elif _is_pure_chat_message(prompt):
        if st.session_state.get("pending_new_ppt") or st.session_state.get("pending_intent") or st.session_state.get("pending_action"):
            st.session_state.pending_new_ppt = None
            st.session_state.active_edit_context = None
            _clear_pending_turn_state()
        reply = (
            "👋 Hello! I'm your PPT assistant. You can ask me to create a new presentation, edit slides, add content, or switch between decks. What would you like to do?"
            if re.search(r"\b(hi|hello|hey)\b", prompt, re.IGNORECASE)
            else "I'm here when you're ready to work on your presentation."
        )
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    # Highest-priority route: any explicit create request should start a new deck,
    # even if the user included content in the same message.
    if _is_explicit_create_request(prompt):
        st.session_state.pending_new_ppt = None
        handled = _handle_create_request(prompt)
        if handled:
            st.stop()

    cancelled_create = st.session_state.get("last_cancelled_create") or {}
    recovery_slide_count = _extract_slide_count_from_text(prompt)
    if (
        cancelled_create.get("topic")
        and recovery_slide_count
        and re.search(r"\b(?:generate|create|make|build)\b", prompt, re.IGNORECASE)
    ):
        topic = cancelled_create.get("topic")
        sections = cancelled_create.get("sections")
        theme_colors = cancelled_create.get("theme_colors") or extract_theme_colors_from_messages(st.session_state.get("messages", [])) or extract_theme_colors(topic)
        st.session_state["last_cancelled_create"] = None
        if int(recovery_slide_count) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
            question = _slide_limit_confirmation_question(int(recovery_slide_count))
            pending = {
                "intent": "confirm_large_slide_count",
                "target_intent": "create_ppt",
                "slots": {"topic": topic, "slide_count": int(recovery_slide_count), "sections": sections, "theme_colors": theme_colors},
                "next_question": question,
                "action": "confirm",
            }
            _store_pending_action(pending)
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        generate_outline_and_reply(topic, int(recovery_slide_count), st.session_state.tone, sections, theme_colors)
        st.stop()

    # If we're waiting for the slide count of a newly requested deck, keep
    # that flow pinned until the user answers or cancels.
    pending_new_ppt = st.session_state.get("pending_new_ppt")
    if pending_new_ppt:
        pending_topic = pending_new_ppt.get("topic", "")
        if pending_new_ppt.get("need_topic") or not pending_topic:
            topic_candidate = _extract_conversational_create_topic(prompt)
            if not topic_candidate:
                question = _create_topic_followup(prompt)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            pending_topic = topic_candidate
            pending_new_ppt["topic"] = pending_topic
            pending_new_ppt["need_topic"] = False
            st.session_state.pending_new_ppt = pending_new_ppt
        slide_count = _extract_slide_count_from_text(prompt)
        if slide_count:
            all_messages = st.session_state.get("messages", [])
            theme_colors = (
                pending_new_ppt.get("theme_colors")
                or extract_theme_colors_from_messages(all_messages)
                or extract_theme_colors(pending_topic)
            )
            st.session_state.pending_new_ppt = None
            st.session_state.pending_intent = None
            st.session_state.pending_action = None
            if int(slide_count) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
                question = _slide_limit_confirmation_question(int(slide_count))
                pending = {
                    "intent": "confirm_large_slide_count",
                    "target_intent": "create_ppt",
                    "slots": {
                        "topic": pending_topic,
                        "slide_count": int(slide_count),
                        "sections": pending_new_ppt.get("sections"),
                    },
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            generate_outline_and_reply(
                pending_topic,
                slide_count,
                st.session_state.tone,
                pending_new_ppt.get("sections"),
                theme_colors,
            )
            st.stop()
        if re.search(r"\b(?:cancel|stop|never mind|nevermind)\b", prompt, re.IGNORECASE):
            st.session_state.pending_new_ppt = None
            with st.chat_message("assistant"):
                st.markdown("Okay, I cancelled the new presentation request.")
            add_message("assistant", "Okay, I cancelled the new presentation request.")
            st.rerun()
        with st.chat_message("assistant"):
            st.markdown("How many slides would you like the presentation to have?")
        add_message("assistant", "How many slides would you like the presentation to have?")
        st.rerun()

    routing_prompt = prompt
    if _is_explicit_create_request(prompt):
        routing_prompt = prompt
    else:
        routing_prompt = _resolve_vague_followup_prompt(prompt)
    previous_ppt_request = _is_previous_ppt_request(routing_prompt)

    # If there's an active pending slot-filling intent, avoid re-classifying
    # the user's follow-up as a new intent unless they explicitly signal a
    # different action. This preserves slots across turns and enables
    # true multi-turn slot filling.
    pending_state = st.session_state.get("pending_action") or st.session_state.get("pending_intent") or {}
    reasoned = None
    if pending_state and pending_state.get("action") in {"ask", "confirm"}:
        # If the user explicitly uses verbs like 'switch', 'create', or
        # 'cancel', allow re-classification. Avoid overly-permissive words
        # like 'new' or 'make' which often appear inside clarifying
        # questions (e.g. "Where should I place the new slide?").
        if re.search(
            r"\b(?:switch|change|create|generate|cancel|stop|nevermind|never mind|"
            r"show|view|display|see|explain|describe|summarize|remove|delete|merge|combine|"
            r"add|insert|include|update|edit|improve|shorten|rewrite|expand|regenerate|redo|rebuild|"
            r"script|notes?|audience|questions?|speech|presenting|present)\b",
            prompt,
            re.IGNORECASE,
        ):
            reasoning_state = _reasoning_state_snapshot()
            reasoned = reasoning_layer(routing_prompt, reasoning_state)
            routing_prompt = _canonicalize_reasoned_prompt(routing_prompt, reasoned)
        else:
            # Preserve existing intent; create minimal reasoned dict so
            # downstream code can read `intent` as before.
            reasoned = {"intent": pending_state.get("intent"), "confidence": 0.99}
    else:
        reasoning_state = _reasoning_state_snapshot()
        reasoned = reasoning_layer(routing_prompt, reasoning_state)
        routing_prompt = _canonicalize_reasoned_prompt(routing_prompt, reasoned)

    context_outline_payload = _current_outline_payload()
    context_slides = context_outline_payload.get("slides", []) if context_outline_payload else []
    assistant_state = build_assistant_state(context_slides)
    contextual_reasoned = resolve_contextual_intent(prompt, assistant_state)
    if contextual_reasoned and float(contextual_reasoned.get("confidence") or 0.0) >= 0.55:
        reasoned = {**reasoned, **contextual_reasoned}
        routing_prompt = _canonicalize_reasoned_prompt(routing_prompt, reasoned)

    direct_intent = reasoned.get("intent")

    if _is_deck_level_edit_request(prompt):
        target_ppt_ref = reasoned.get("ppt_id") or _extract_explicit_ppt_ref_from_text(routing_prompt) or routing_prompt
        changed, new_ppt_id, clarification = resolve_ppt_reference(str(target_ppt_ref))
        if clarification:
            with st.chat_message("assistant"):
                st.markdown(clarification)
            add_message("assistant", clarification)
            st.rerun()
        if changed and new_ppt_id:
            st.session_state.action_target_ppt_id = new_ppt_id
            st.session_state.active_edit_context = {
                "level": "ppt",
                "ppt_id": new_ppt_id,
            }
            ppt_label = _ppt_display_label(new_ppt_id)
            followup = f"What changes would you like to make in {ppt_label}?"
            with st.chat_message("assistant"):
                st.markdown(followup)
            add_message("assistant", followup)
            st.rerun()

    reasoned_ppt_ref = reasoned.get("ppt_id")
    if reasoned_ppt_ref:
        resolved_reasoned_ppt = _resolve_ppt_history_item(str(reasoned_ppt_ref))
        if resolved_reasoned_ppt and resolved_reasoned_ppt.get("id"):
            reasoned_ppt_id = resolved_reasoned_ppt["id"]
            if reasoned_ppt_id != st.session_state.get("current_ppt_id"):
                switch_active_ppt(reasoned_ppt_id)
            st.session_state.action_target_ppt_id = reasoned_ppt_id
            if direct_intent == "switch_ppt":
                st.session_state.current_ppt_id = reasoned_ppt_id

    # Resolve explicit deck references before any slide-level follow-up logic.
    # This prevents "add slide in ppt 1" from asking a generic follow-up while
    # still operating on whatever deck happened to be active previously.
    if previous_ppt_request:
        early_changed, early_ppt_id, early_clarification = resolve_ppt_reference(routing_prompt)
        if early_clarification:
            with st.chat_message("assistant"):
                st.markdown(early_clarification)
            add_message("assistant", early_clarification)
            st.rerun()
        if early_changed and early_ppt_id:
            st.session_state.action_target_ppt_id = early_ppt_id
            active_item = get_ppt_by_id(early_ppt_id)
            if active_item and active_item.get("outline_payload"):
                preview_slides = active_item["outline_payload"].get("slides", [])
                topic = active_item.get("topic", "Presentation")
                label = f"Preview — {topic} ({early_ppt_id})"
                add_message_with_preview(
                    role="assistant",
                    text=f"Here is your previous presentation: **{early_ppt_id}** — {topic}.",
                    preview_slides=preview_slides,
                    ppt_label=label,
                    ppt_bytes=active_item.get("ppt_bytes"),
                    ppt_filename=active_item.get("ppt_filename", "presentation.pptx"),
                )
            else:
                with st.chat_message("assistant"):
                    st.markdown(f"Switched back to **{_ppt_display_label(early_ppt_id)}**.")
                add_message("assistant", f"Switched back to {_ppt_display_label(early_ppt_id)}.")
            st.rerun()
    if re.search(r"(ppt|presentation|deck)\s*\d+|ppt[_-]\d+|switch|open|go to", routing_prompt, re.IGNORECASE):
        early_changed, early_ppt_id, early_clarification = resolve_ppt_reference(routing_prompt)
        if early_clarification:
            with st.chat_message("assistant"):
                st.markdown(early_clarification)
            add_message("assistant", early_clarification)
            st.rerun()
        if early_changed:
            st.session_state.action_target_ppt_id = early_ppt_id

    outline_payload = _current_outline_payload()
    slides = outline_payload.get("slides", []) if outline_payload else []
    clear_slide_request = _is_clear_slide_request(routing_prompt)
    blank_slide_request = _is_blank_slide_request(routing_prompt)

    pending_state = st.session_state.get("pending_action") or st.session_state.get("pending_intent") or {}
    protected_pending_intents = {"confirm_topic_switch", "confirm_add_points_limit", "confirm_large_slide_count", "confirm_add_slide", "clear_slide"}
    if pending_state and direct_intent in {"greeting", "smalltalk"} and pending_state.get("intent") not in protected_pending_intents:
        _clear_pending_turn_state()
        pending_state = {}
    elif pending_state and direct_intent == "general_request" and _is_unrelated_chat_message(prompt) and pending_state.get("intent") not in protected_pending_intents:
        _clear_pending_turn_state()
        pending_state = {}

    if pending_state.get("intent") == "clear_slide":
        if _is_cancel_command(routing_prompt) or _is_negative_command(routing_prompt):
            _clear_pending_turn_state()
            reply = "Okay, I cancelled that request."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if _is_affirmative_command(routing_prompt):
            slots = dict(pending_state.get("slots", {}))
            result_intent = "blank_slide" if slots.get("clear_mode") == "blank" else "clear_slide"
            result_msg, success = execute_action(result_intent, slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                st.rerun()
        else:
            slots = dict(pending_state.get("slots", {}))
            if slots.get("slide_number") is None:
                slide_idx = _extract_explicit_slide_number_from_text(routing_prompt) or _resolve_slide_reference_text(routing_prompt, slides)
                if slide_idx is not None:
                    slots["slide_number"] = slide_idx
                    st.session_state.pending_intent["slots"] = slots
                    st.session_state.pending_action = st.session_state.pending_intent
            if slots.get("slide_number") is None:
                question = _polish_next_question(prompt, "clear_slide", slots, ["slide_number"], None)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()

    if pending_state.get("intent") == "confirm_add_points_limit":
        if _is_cancel_command(routing_prompt) or _is_negative_command(routing_prompt):
            _clear_pending_turn_state()
            reply = "Okay, I cancelled that request."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if _is_affirmative_command(routing_prompt):
            slots = dict(pending_state.get("slots", {}))
            result_msg, success = execute_action("edit_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                _clear_pending_turn_state()
                st.rerun()
        question = pending_state.get("next_question") or _point_limit_confirmation_question(
            int((pending_state.get("slots") or {}).get("slide_number") or 0),
            int((pending_state.get("requested_points") or _MAX_POINTS_PER_REQUEST)),
            int((pending_state.get("allowed_points") or _MAX_POINTS_PER_REQUEST)),
            st.session_state.get("current_ppt_id"),
        )
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if pending_state.get("intent") == "confirm_large_slide_count":
        slots = dict(pending_state.get("slots", {}))
        target_intent = pending_state.get("target_intent") or "create_ppt"
        revised_count = _extract_slide_count_from_text(routing_prompt)
        original_count = slots.get("slide_count") or slots.get("target_count")
        if target_intent == "create_ppt" and revised_count and int(revised_count) != int(original_count or 0):
            slots["slide_count"] = int(revised_count)
            _clear_pending_turn_state()
            if int(revised_count) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
                question = _slide_limit_confirmation_question(int(revised_count))
                pending = {
                    "intent": "confirm_large_slide_count",
                    "target_intent": "create_ppt",
                    "slots": slots,
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            result_msg, success = execute_action("create_ppt", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                st.rerun()
        if _is_cancel_command(routing_prompt) or _is_negative_command(routing_prompt):
            if target_intent == "create_ppt" and slots.get("topic"):
                st.session_state["last_cancelled_create"] = {
                    "topic": slots.get("topic"),
                    "sections": slots.get("sections"),
                    "theme_colors": slots.get("theme_colors"),
                }
            _clear_pending_turn_state()
            reply = "Okay, I cancelled that request."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if _is_affirmative_command(routing_prompt):
            _clear_pending_turn_state()
            if target_intent == "bulk_resize":
                target_count = int(slots.get("target_count") or slots.get("slide_count") or 0)
                if target_count <= len(slides):
                    updated_slides, removed = _condense_existing_deck_slides(slides, target_count)
                    if removed > 0 and len(updated_slides) == target_count:
                        commit_changes(updated_slides, f"Condensed the deck to {len(updated_slides)} slides.")
                        st.stop()
                    result_msg = f"I could not condense this deck to {target_count} slides automatically."
                    success = False
                else:
                    updated_slides, added = _expand_existing_deck_slides(slides, target_count)
                    if added > 0:
                        commit_changes(updated_slides, f"Expanded the deck to {len(updated_slides)} slides by splitting existing content.")
                        st.stop()
                    result_msg = "I could not expand this deck automatically. Try a slightly smaller target or ask to expand specific slides."
                    success = False
            elif target_intent == "single_slide_resize":
                target_count = int(slots.get("target_count") or slots.get("slide_count") or 0)
                slide_num = slots.get("slide_number")
                try:
                    slide_num = int(slide_num)
                except (TypeError, ValueError):
                    slide_num = None
                if slide_num is None:
                    result_msg = "Which slide should I expand?"
                    success = False
                elif not (1 <= slide_num <= len(slides)):
                    result_msg = f"Slide {slide_num} does not exist. The deck has {len(slides)} slide(s)."
                    success = False
                else:
                    updated_slides, added = _expand_target_slide_in_deck(slides, slide_num, target_count)
                    if added > 0:
                        commit_changes(updated_slides, f"Expanded slide {slide_num} into {target_count} focused slides.")
                        st.stop()
                    result_msg = f"I could not split slide {slide_num} into {target_count} meaningful parts automatically."
                    success = False
            else:
                result_msg, success = execute_action("create_ppt", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                st.rerun()
        question = pending_state.get("next_question") or _slide_limit_confirmation_question(
            int((pending_state.get("slots") or {}).get("slide_count") or (pending_state.get("slots") or {}).get("target_count") or 0)
        )
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if pending_state.get("intent") == "confirm_add_slide":
        # Expect a clear yes/no. If affirmative, execute; if negative/cancel, clear.
        if _is_cancel_command(routing_prompt) or _is_negative_command(routing_prompt):
            _clear_pending_turn_state()
            reply = "Okay, I cancelled that request."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if _is_affirmative_command(routing_prompt):
            slots = dict(pending_state.get("slots", {}))
            result_msg, success = execute_action("add_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                _clear_pending_turn_state()
                st.rerun()
        # otherwise re-prompt the confirmation question
        question = pending_state.get("next_question") or "Do you want me to add that slide now?"
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if pending_state.get("intent") == "confirm_topic_switch":
        if _is_cancel_command(routing_prompt) or _is_negative_command(routing_prompt):
            _clear_pending_turn_state()
            reply = "Okay, I’ll stay with the current topic context."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if _is_affirmative_command(routing_prompt):
            slots = dict(pending_state.get("slots", {}))
            _clear_pending_turn_state()
            result_msg, success = execute_action("suggest_topic", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                st.rerun()
        question = pending_state.get("next_question") or _topic_switch_confirmation_question(
            _safe_str(pending_state.get("active_domain", "")),
            _safe_str((pending_state.get("slots") or {}).get("topic", "")),
        )
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if _is_readonly_conversation_request(routing_prompt):
        reply = generate_conversational_answer(routing_prompt, assistant_state, slides)
        target_slide = _readonly_target_slide(routing_prompt, assistant_state, slides)
        _remember_answer_objects(routing_prompt, reply, target_slide)
        _remember_action_context(
            "readonly_advice",
            target_slide,
            st.session_state.get("current_ppt_id"),
            f"Answered read-only advice request: {routing_prompt}",
        )
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    if _is_new_structured_slide_request(routing_prompt):
        _clear_pending_turn_state()
        slots = {
            "slide_content": _structured_slide_content_from_request(routing_prompt),
            "content": _structured_slide_content_from_request(routing_prompt),
        }
        position_candidate = _extract_add_slide_position(routing_prompt, slides)
        if position_candidate is not None:
            slots["position"] = position_candidate
        missing_slots = []
        if not slots.get("slide_content"):
            missing_slots.append("slide_content")
        if slots.get("position") is None:
            missing_slots.append("position")
        if not missing_slots:
            result_msg, success = execute_action("add_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if success:
                st.rerun()
        followup = _polish_next_question(routing_prompt, "add_slide", slots, missing_slots, None)
        st.session_state.pending_intent = {
            "intent": "add_slide",
            "slots": slots,
            "missing_slots": missing_slots,
            "next_question": followup,
            "action": "ask",
        }
        st.session_state.pending_action = st.session_state.pending_intent
        with st.chat_message("assistant"):
            st.markdown(followup)
        add_message("assistant", followup)
        st.rerun()

    if direct_intent in {"greeting", "smalltalk"}:
        _clear_pending_turn_state()
        result_msg, success = execute_action(direct_intent, {}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        st.rerun()

    if direct_intent == "design_change":
        _clear_pending_turn_state()
        slots = {
            "user_request": routing_prompt,
            "change_content": routing_prompt,
            "sub_intent": reasoned.get("sub_intent"),
        }
        result_msg, success = execute_action("design_change", slots, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    structured_target = None
    if direct_intent == "unknown" and _needs_layout_aware_update(routing_prompt, None):
        structured_target = (
            _extract_explicit_slide_number_from_text(routing_prompt)
        )
        if structured_target is not None:
            _clear_pending_turn_state()
            result_msg, success = execute_action(
                "edit_slide",
                {"slide_number": structured_target, "change_content": routing_prompt},
                slides,
            )
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()

    if direct_intent == "unknown":
        reply = reasoned.get("clarification_question") or "What would you like me to do with the presentation?"
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    if direct_intent == "bulk_update":
        _clear_pending_turn_state()
        sub_intent = _safe_str(reasoned.get("sub_intent", ""))
        if sub_intent == "resize":
            target_count = reasoned.get("slide_count_target") or _extract_slide_count_from_text(routing_prompt)
            try:
                target_count = int(target_count)
            except (TypeError, ValueError):
                target_count = None
            if not target_count or target_count < 1:
                reply = "How many slides should I expand the presentation to?"
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
            single_slide_expand = bool(re.search(r"\bexpand\b.*\bslide\b.*\b(?:to|into)\s+\d+\s+slides?\b", routing_prompt, re.IGNORECASE))
            if int(target_count) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
                target_slide_num = None
                if single_slide_expand:
                    target_slide_num = reasoned.get("slide_id") or _extract_explicit_slide_number_from_text(routing_prompt) or st.session_state.get("active_slide_index") or st.session_state.get("last_slide_index")
                question = _slide_limit_confirmation_question(int(target_count))
                pending = {
                    "intent": "confirm_large_slide_count",
                    "target_intent": "single_slide_resize" if single_slide_expand else "bulk_resize",
                    "slots": {"target_count": int(target_count), "slide_count": int(target_count), "slide_number": target_slide_num},
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            if single_slide_expand:
                target_slide_num = reasoned.get("slide_id") or _extract_explicit_slide_number_from_text(routing_prompt) or st.session_state.get("active_slide_index") or st.session_state.get("last_slide_index")
                try:
                    target_slide_num = int(target_slide_num)
                except (TypeError, ValueError):
                    target_slide_num = None
                if target_slide_num is None:
                    reply = "Which slide should I expand?"
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
                if not (1 <= target_slide_num <= len(slides)):
                    reply = f"Slide {target_slide_num} does not exist. The deck has {len(slides)} slide(s)."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
                if int(target_count) <= 1:
                    reply = "Please choose a number greater than 1 for slide expansion."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
                updated_slides, added = _expand_target_slide_in_deck(slides, target_slide_num, int(target_count))
                if added <= 0:
                    reply = f"I could not split slide {target_slide_num} into {target_count} meaningful parts automatically."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
                commit_changes(updated_slides, f"Expanded slide {target_slide_num} into {target_count} focused slides.")
                st.stop()
            if not slides:
                reply = "No active deck found to expand. Please open or create a presentation first."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
            if target_count < len(slides):
                updated_slides, removed = _condense_existing_deck_slides(slides, int(target_count))
                if removed <= 0 or len(updated_slides) != int(target_count):
                    reply = f"I could not condense this deck to {target_count} slides automatically."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
                commit_changes(updated_slides, f"Condensed the deck to {len(updated_slides)} slides.")
                st.stop()
            if target_count == len(slides):
                updated_slides = [_shorten_slide_structure(s) for s in slides]
                commit_changes(updated_slides, f"Kept the deck at {len(updated_slides)} slides and tightened the slide content.")
                st.stop()
            updated_slides, added = _expand_existing_deck_slides(slides, int(target_count))
            if added <= 0:
                reply = "I could not expand this deck automatically. Try a slightly larger target or ask to expand specific slides."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
            commit_changes(updated_slides, f"Expanded the deck to {len(updated_slides)} slides by splitting existing content.")
            st.stop()
        if sub_intent in {"shorten", "simplify"} or re.search(r"\b(shorten|condense|simplify|compress|less text|concise)\b", routing_prompt, re.IGNORECASE):
            updated_slides = [_shorten_slide_structure(s) for s in slides]
            commit_changes(updated_slides, "Tightened the deck content while preserving the slide count.")
            st.stop()
        reply = "I can handle deck-wide updates like resize. Please specify the target, for example: shorten this PPT to 5 slides."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    if direct_intent == "refine_ppt":
        _clear_pending_turn_state()
        slots = {
            "user_request": routing_prompt,
            "change_content": reasoned.get("content") or routing_prompt,
            "style_hint": reasoned.get("style_hint") or reasoned.get("content") or routing_prompt,
            "sub_intent": reasoned.get("sub_intent"),
        }
        result_msg, success = execute_action("refine_ppt", slots, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent in {"ppt_info", "suggest_topic", "create_ppt"}:
        _clear_pending_turn_state()
        if direct_intent == "ppt_info":
            slots = _extract_ppt_info_slots(routing_prompt, reasoned)
            result_msg, success = execute_action("ppt_info", slots, slides)
        elif direct_intent == "suggest_topic":
            slots = _extract_suggest_topic_slots(routing_prompt)
            active_domain = _get_active_topic_domain()
            topic_hint = _safe_str(slots.get("topic", ""))
            if _is_topic_fragment_request(routing_prompt):
                if active_domain:
                    slots["scope"] = "current"
                    slots["topic"] = _merge_topic_refinement(active_domain, topic_hint or routing_prompt)
                elif _is_ambiguous_topic_request(routing_prompt, active_domain):
                    reply = "Do you want topic ideas about that area, or are you refining a previous topic?"
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.rerun()
            if topic_hint and active_domain and _topic_domains_conflict(topic_hint, active_domain):
                question = _topic_switch_confirmation_question(active_domain, topic_hint)
                pending_switch = {
                    "intent": "confirm_topic_switch",
                    "slots": slots,
                    "active_domain": active_domain,
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending_switch)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            result_msg, success = execute_action("suggest_topic", slots, slides)
        else:
            slots = _extract_create_ppt_slots(routing_prompt)
            if not slots.get("topic") or not slots.get("slide_count"):
                missing = []
                if not slots.get("topic"):
                    missing.append("topic")
                if not slots.get("slide_count"):
                    missing.append("slide_count")
                reply = _polish_next_question(prompt, "create_ppt", slots, missing, None)
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
            if int(slots.get("slide_count")) > _MAX_SLIDES_WITHOUT_CONFIRMATION:
                question = _slide_limit_confirmation_question(int(slots.get("slide_count")))
                pending = {
                    "intent": "confirm_large_slide_count",
                    "target_intent": "create_ppt",
                    "slots": slots,
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            result_msg, success = execute_action("create_ppt", slots, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent == "add_slide" and pending_state.get("intent") != "add_slide":
        _clear_pending_turn_state()
        slots = {}
        explicit_content = _extract_explicit_provided_content(prompt)
        content_candidate = explicit_content or _extract_add_slide_content(prompt)
        position_candidate = _extract_add_slide_position(prompt, slides)
        if content_candidate:
            slots["slide_content"] = content_candidate
            slots["content"] = content_candidate
            if explicit_content:
                slots["explicit_user_content"] = explicit_content
        else:
            # Fall back to classifier/reasoned content if available
            if reasoned and reasoned.get("content"):
                slots["slide_content"] = reasoned.get("content")
                slots["content"] = reasoned.get("content")
        if position_candidate is not None:
            slots["position"] = position_candidate

        missing_slots = []
        if not slots.get("slide_content"):
            missing_slots.append("slide_content")
        if slots.get("position") is None:
            missing_slots.append("position")

        if not missing_slots:
            # Ask for explicit confirmation before creating the slide
            pos = slots.get("position")
            if isinstance(pos, int):
                pos_text = f"position {pos}"
            elif isinstance(pos, str):
                pos_text = ("the end" if pos == "end" else "the start" if pos == "start" else pos)
            else:
                pos_text = "the requested position"
            content_preview = (slots.get("slide_content") or "a new slide").strip()
            question = f"Do you want me to add a slide titled: \"{content_preview}\" at {pos_text}?"
            pending_confirm = {"intent": "confirm_add_slide", "slots": slots, "next_question": question, "action": "confirm"}
            _store_pending_action(pending_confirm)
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()

        followup = _polish_next_question(prompt, "add_slide", slots, missing_slots, None)
        st.session_state.pending_intent = {
            "intent": "add_slide",
            "slots": slots,
            "missing_slots": missing_slots,
            "next_question": followup,
            "action": "ask",
        }
        st.session_state.pending_action = st.session_state.pending_intent
        with st.chat_message("assistant"):
            st.markdown(followup)
        add_message("assistant", followup)
        st.rerun()

    if clear_slide_request or blank_slide_request:
        slide_num = reasoned.get("slide_id") or _extract_explicit_slide_number_from_text(routing_prompt) or _resolve_slide_reference_text(routing_prompt, slides)
        if slide_num is None:
            reply = _polish_next_question(prompt, "blank_slide" if blank_slide_request else "clear_slide", {}, ["slide_number"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if clear_slide_request:
            question = f"This will remove everything from slide {slide_num}. Do you want me to continue?"
            st.session_state.pending_intent = {
                "intent": "clear_slide",
                "slots": {"slide_number": slide_num, "clear_mode": "clear"},
                "missing_slots": [],
                "next_question": question,
                "action": "confirm",
            }
            st.session_state.pending_action = st.session_state.pending_intent
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        result_msg, success = execute_action("blank_slide", {"slide_number": slide_num, "clear_mode": "blank"}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent == "switch_ppt":
        _clear_pending_turn_state()
        ppt_ref = reasoned.get("ppt_id") or _extract_explicit_ppt_ref_from_text(routing_prompt)
        changed, new_ppt_id, clarification = resolve_ppt_reference(str(ppt_ref or routing_prompt))
        if clarification:
            with st.chat_message("assistant"):
                st.markdown(clarification)
            add_message("assistant", clarification)
            st.rerun()
        if changed and new_ppt_id:
            st.session_state.action_target_ppt_id = new_ppt_id
            active_item = get_ppt_by_id(new_ppt_id)
            if active_item and active_item.get("outline_payload") and not previous_ppt_request and re.search(r"\b(?:open|switch to|go to|show)\b", routing_prompt, re.IGNORECASE):
                preview_slides = active_item["outline_payload"].get("slides", [])
                topic = active_item.get("topic", "Presentation")
                label = f"Preview — {topic} ({new_ppt_id})"
                add_message_with_preview(
                    role="assistant",
                    text=f"Showing preview for **{new_ppt_id}** — {topic}.",
                    preview_slides=preview_slides,
                    ppt_label=label,
                    ppt_bytes=active_item.get("ppt_bytes"),
                    ppt_filename=active_item.get("ppt_filename", "presentation.pptx"),
                )
            elif not previous_ppt_request:
                active_label = _ppt_display_label(new_ppt_id)
                with st.chat_message("assistant"):
                    message = f"✅ Active deck set to **{active_label}**."
                    st.markdown(message)
                add_message("assistant", f"Active deck set to {active_label}.")
            else:
                active_label = _ppt_display_label(new_ppt_id)
                active_item = get_ppt_by_id(new_ppt_id)
                if active_item and active_item.get("outline_payload"):
                    preview_slides = active_item["outline_payload"].get("slides", [])
                    topic = active_item.get("topic", "Presentation")
                    label = f"Preview — {topic} ({new_ppt_id})"
                    add_message_with_preview(
                        role="assistant",
                        text=f"Here is your previous presentation: **{new_ppt_id}** — {topic}.",
                        preview_slides=preview_slides,
                        ppt_label=label,
                        ppt_bytes=active_item.get("ppt_bytes"),
                        ppt_filename=active_item.get("ppt_filename", "presentation.pptx"),
                    )
                else:
                    with st.chat_message("assistant"):
                        st.markdown(f"Switched back to **{active_label}**.")
                    add_message("assistant", f"Switched back to {active_label}.")
            st.rerun()
        with st.chat_message("assistant"):
            st.markdown("I couldn’t find that presentation. Try `ppt 1`, `ppt 2`, or the deck title.")
        add_message("assistant", "I couldn’t find that presentation. Try `ppt 1`, `ppt 2`, or the deck title.")
        st.rerun()

    if direct_intent == "view_slide":
        if reasoned.get("slide_id") is None:
            reply = _polish_next_question(routing_prompt, "view_slide", {}, ["slide_number"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        view_slots = {
            "slide_number": reasoned.get("slide_id"),
            "ppt_ref": reasoned.get("ppt_id"),
        }
        if not _should_execute_directly(reasoned):
            reply = _polish_next_question(routing_prompt, "view_slide", view_slots, [], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("view_slide", view_slots, slides)
        if success:
            _remember_action_context(
                "view_slide",
                reasoned.get("slide_id"),
                st.session_state.get("current_ppt_id"),
                f"Viewed slide {reasoned.get('slide_id')}",
            )
            preface = _confidence_preface(reasoned)
            full_msg = f"{preface}\n\n{result_msg}".strip() if preface else result_msg
            with st.chat_message("assistant"):
                st.markdown(full_msg)
            add_message("assistant", full_msg)
            st.rerun()
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        st.rerun()

    if direct_intent == "explain_slide":
        slide_num = reasoned.get("slide_id")
        if slide_num is None:
            reply = _polish_next_question(routing_prompt, "explain_slide", {}, ["slide_number"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("explain_slide", {"slide_number": slide_num}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        st.rerun()

    if direct_intent == "delete_slide":
        slide_num = reasoned.get("slide_id")
        if slide_num is None:
            reply = _polish_next_question(routing_prompt, "delete_slide", {}, ["slide_number"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("delete_slide", {"slide_number": slide_num}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent == "merge_slides":
        slide_ids = reasoned.get("slide_ids") or []
        if len(slide_ids) < 2:
            reply = "Which two slides should I merge?"
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("merge_slides", {"slide_ids": slide_ids, "slide_a": slide_ids[0], "slide_b": slide_ids[1]}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent == "regenerate_slide":
        slide_num = reasoned.get("slide_id") or _extract_explicit_slide_number_from_text(routing_prompt) or _resolve_slide_reference_text(routing_prompt, slides)
        if slide_num is None:
            reply = "Which slide should I regenerate?"
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("regenerate_slide", {"slide_number": slide_num}, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        if success:
            st.rerun()

    if direct_intent == "presentation_mode":
        _clear_pending_turn_state()
        result_msg, success = execute_action(
            "presentation_mode",
            {
                "sub_intent": reasoned.get("sub_intent"),
                "user_request": prompt,
            },
            slides,
        )
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        st.rerun()

    if direct_intent == "edit_slide" and reasoned.get("slide_id") is not None:
        if not _slide_exists(reasoned.get("slide_id"), slides):
            _clear_pending_turn_state()
            st.session_state.active_edit_context = None
            st.session_state.last_slide_index = None
            st.session_state.active_slide_index = None
            reply = f"Slide {reasoned.get('slide_id')} does not exist in {_ppt_display_label(st.session_state.get('current_ppt_id'))}. Which slide would you like to edit?"
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if not _has_meaningful_edit_instruction(routing_prompt):
            pending_edit = {
                "intent": "edit_slide",
                "slots": {
                    "slide_number": int(reasoned.get("slide_id")),
                    "ppt_ref": st.session_state.get("current_ppt_id"),
                },
                "missing_slots": ["change_content"],
                "next_question": None,
                "action": "ask",
            }
            _store_pending_action(pending_edit)
            reply = _polish_next_question(
                routing_prompt,
                "edit_slide",
                {"slide_number": reasoned.get("slide_id")},
                ["change_content"],
                None,
            )
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if not _should_execute_directly(reasoned):
            pending_edit = {
                "intent": "edit_slide",
                "slots": {
                    "slide_number": int(reasoned.get("slide_id")),
                    "ppt_ref": st.session_state.get("current_ppt_id"),
                },
                "missing_slots": ["change_content"],
                "next_question": None,
                "action": "ask",
            }
            _store_pending_action(pending_edit)
            reply = _polish_next_question(
                routing_prompt,
                "edit_slide",
                {"slide_number": reasoned.get("slide_id")},
                ["change_content"],
                None,
            )
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        slide_num = int(reasoned.get("slide_id"))
        change = reasoned.get("content") or routing_prompt
        slots = {"slide_number": slide_num, "change_content": change}
        result_msg, success = execute_action("edit_slide", slots, slides)
        if success:
            preface = _confidence_preface(reasoned)
            full_msg = f"{preface}\n\n{result_msg}".strip() if preface else result_msg
            with st.chat_message("assistant"):
                st.markdown(full_msg)
            add_message("assistant", full_msg)
            st.rerun()

    if direct_intent in {"transform_content", "add_points", "update_slide"}:
        slide_num = reasoned.get("slide_id")
        if slide_num is None:
            reply = _polish_next_question(routing_prompt, direct_intent, {}, ["slide_number"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if not _slide_exists(slide_num, slides):
            _clear_pending_turn_state()
            st.session_state.active_edit_context = None
            st.session_state.last_slide_index = None
            st.session_state.active_slide_index = None
            reply = f"Slide {slide_num} does not exist in {_ppt_display_label(st.session_state.get('current_ppt_id'))}. Which slide would you like to edit?"
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        if not _should_execute_directly(reasoned):
            reply = _polish_next_question(routing_prompt, direct_intent, {"slide_number": slide_num}, ["change_content"], None)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        _clear_pending_turn_state()
        slide_num = int(slide_num)
        change = reasoned.get("content") or routing_prompt
        target_slide = slides[slide_num - 1] if 1 <= slide_num <= len(slides) else None
        side_hint = _resolve_add_points_side_hint(change or routing_prompt, target_slide)
        requested_points = None
        raw_requested_points = None
        if direct_intent in {"add_points", "update_slide"}:
            parsed_points = _parse_bullet_add_request(change)
            if parsed_points is not None:
                raw_requested_points = int(parsed_points)
                requested_points = _resolved_add_point_count(raw_requested_points, side_hint)
            elif direct_intent == "add_points":
                raw_requested_points = 1
                requested_points = _resolved_add_point_count(raw_requested_points, side_hint)
            if raw_requested_points and raw_requested_points > _MAX_POINTS_PER_REQUEST:
                question = _point_limit_confirmation_question(
                    slide_num,
                    int(raw_requested_points),
                    int(requested_points),
                    st.session_state.get("current_ppt_id"),
                )
                pending_limit = {
                    "intent": "confirm_add_points_limit",
                    "slots": {
                        "slide_number": slide_num,
                        "change_content": change,
                        "topic_hint": _extract_topic_constraint_from_text(change),
                        "n_points": requested_points,
                    },
                    "requested_points": int(raw_requested_points),
                    "allowed_points": int(requested_points),
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending_limit)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
        topic_hint = _extract_topic_constraint_from_text(change)
        slots = {"slide_number": slide_num, "change_content": change, "topic_hint": topic_hint}
        if requested_points is not None:
            slots["n_points"] = requested_points
        result_msg, success = execute_action("edit_slide", slots, slides)
        if success:
            preface = _confidence_preface(reasoned)
            added_points = requested_points if requested_points is not None else 1
            cap_note = ""
            raw_requested_points = _parse_bullet_add_request(change or routing_prompt) if direct_intent in {"add_points", "update_slide"} else None
            if raw_requested_points and raw_requested_points > _MAX_POINTS_PER_REQUEST:
                cap_note = f"I can add up to {_MAX_POINTS_PER_REQUEST} points total at a time, so I added {added_points} point(s)."
            success_slots = {"slide_number": slide_num, "change_content": change}
            if requested_points is not None:
                success_slots["n_points"] = added_points
            nat_msg = _natural_success_message(
                direct_intent,
                success_slots,
            )
            if cap_note:
                nat_msg = f"{cap_note}\n\n{nat_msg}"
            full_msg = f"{preface}\n\n{nat_msg}".strip() if preface else nat_msg
            with st.chat_message("assistant"):
                st.markdown(full_msg)
            add_message("assistant", full_msg)
            st.rerun()

    pending_state = st.session_state.get("pending_action") or st.session_state.get("pending_intent") or {}

    if pending_state.get("intent") == "edit_slide" and "slide_number" in (pending_state.get("missing_slots") or []) and (pending_state.get("slots") or {}).get("change_content"):
        if re.search(r"\b(?:cancel|stop|never mind|nevermind)\b", routing_prompt, re.IGNORECASE):
            _clear_pending_turn_state()
            reply = "Okay, I cancelled that edit request."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        slots = dict(pending_state.get("slots", {}))
        target_answer = _extract_pending_slide_answer(prompt, slides)
        if target_answer == "all":
            change_content = _safe_str(slots.get("change_content"), "")
            if re.search(r"\blogo\b", change_content, re.IGNORECASE):
                _handle_logo_insertion_request("insert logo into all slides", slides)
            question = "Please choose one slide number for that suggestion, or ask me to apply a specific deck-wide change."
            st.session_state.pending_intent = {
                "intent": "edit_slide",
                "slots": slots,
                "missing_slots": ["slide_number"],
                "next_question": question,
                "action": "ask",
            }
            st.session_state.pending_action = st.session_state.pending_intent
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        if target_answer is None:
            question = pending_state.get("next_question") or "Which slide should I apply that suggestion to?"
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        slots["slide_number"] = int(target_answer)
        slots["change_content"] = _safe_str(slots.get("change_content"), "")
        if not slots["change_content"]:
            question = _polish_next_question(prompt, "edit_slide", slots, ["change_content"], None)
            st.session_state.pending_intent = {
                "intent": "edit_slide",
                "slots": slots,
                "missing_slots": ["change_content"],
                "next_question": question,
                "action": "ask",
            }
            st.session_state.pending_action = st.session_state.pending_intent
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        _clear_pending_turn_state()
        result_msg, success = execute_action("edit_slide", slots, slides)
        with st.chat_message("assistant"):
            st.markdown(result_msg)
        add_message("assistant", result_msg)
        st.rerun()

    if pending_state.get("intent") in ("move_slide", "swap_slides") and (pending_state.get("missing_slots") or []):
        if re.search(r"\b(?:cancel|stop|never mind|nevermind)\b", routing_prompt, re.IGNORECASE):
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            st.session_state.pending_intent = None
            st.session_state.pending_action = None
            with st.chat_message("assistant"):
                st.markdown("Okay, I cancelled that request.")
            add_message("assistant", "Okay, I cancelled that request.")
            st.rerun()

        slots = dict(pending_state.get("slots", {}))
        intent = pending_state.get("intent")

        if intent == "move_slide":
            if slots.get("slide_number") is None:
                slide_idx = _extract_explicit_slide_number_from_text(prompt)
                if slide_idx is not None:
                    slots["slide_number"] = slide_idx
                else:
                    slide_idx = _resolve_slide_reference_text(prompt, slides)
                    if slide_idx is not None:
                        slots["slide_number"] = slide_idx
            if slots.get("anchor_slide") is None and slots.get("position") is None:
                pos_text = _norm_text(prompt)
                if re.search(r"\b(end|last|bottom)\b", pos_text):
                    slots["position"] = "end"
                elif re.search(r"\b(start|first|top)\b", pos_text):
                    slots["position"] = "start"
                else:
                    anchor_idx = _resolve_slide_reference_text(prompt, slides)
                    if anchor_idx is not None:
                        slots["anchor_slide"] = anchor_idx
                    else:
                        pos_idx = _extract_slide_count_from_text(prompt)
                        if pos_idx is not None:
                            slots["position"] = pos_idx
            resolved_pending = _resolve_move_swap_slots({"intent": intent, "slots": slots, "missing_slots": []}, slides)
            if resolved_pending and not resolved_pending.get("missing_slots"):
                result_msg, success = execute_action(intent, resolved_pending.get("slots", {}), slides)
                with st.chat_message("assistant"):
                    st.markdown(result_msg)
                add_message("assistant", result_msg)
                if "agent" in st.session_state:
                    st.session_state.agent.clear_state()
                st.rerun()
            question = _polish_next_question(prompt, intent, slots, resolved_pending.get("missing_slots", []) if resolved_pending else ["slide_number", "anchor_slide"], None)
            st.session_state.pending_intent = {"intent": intent, "slots": slots, "missing_slots": resolved_pending.get("missing_slots", []) if resolved_pending else ["slide_number", "anchor_slide"], "next_question": question, "action": "ask"}
            st.session_state.pending_action = st.session_state.pending_intent
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()

        if intent == "swap_slides":
            if slots.get("slide_a") is None:
                idx_a = _extract_explicit_slide_number_from_text(prompt) or _resolve_slide_reference_text(prompt, slides)
                if idx_a is not None:
                    slots["slide_a"] = idx_a
            elif slots.get("slide_b") is None:
                idx_b = _extract_explicit_slide_number_from_text(prompt) or _resolve_slide_reference_text(prompt, slides)
                if idx_b is not None:
                    slots["slide_b"] = idx_b
            resolved_pending = _resolve_move_swap_slots({"intent": intent, "slots": slots, "missing_slots": []}, slides)
            if resolved_pending and not resolved_pending.get("missing_slots"):
                result_msg, success = execute_action(intent, resolved_pending.get("slots", {}), slides)
                with st.chat_message("assistant"):
                    st.markdown(result_msg)
                add_message("assistant", result_msg)
                if "agent" in st.session_state:
                    st.session_state.agent.clear_state()
                st.rerun()
            question = _polish_next_question(prompt, intent, slots, resolved_pending.get("missing_slots", []) if resolved_pending else ["slide_a", "slide_b"], None)
            st.session_state.pending_intent = {"intent": intent, "slots": slots, "missing_slots": resolved_pending.get("missing_slots", []) if resolved_pending else ["slide_a", "slide_b"], "next_question": question, "action": "ask"}
            st.session_state.pending_action = st.session_state.pending_intent
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()

    move_swap_request = _parse_move_swap_request(routing_prompt, slides)
    if move_swap_request:
        resolved_move_swap = _resolve_move_swap_slots(move_swap_request, slides)
        if resolved_move_swap and resolved_move_swap.get("missing_slots"):
            st.session_state.pending_intent = resolved_move_swap
            st.session_state.pending_action = resolved_move_swap
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            question = _polish_next_question(prompt, resolved_move_swap.get("intent"), resolved_move_swap.get("slots", {}), resolved_move_swap.get("missing_slots", []), None)
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()
        if resolved_move_swap:
            result_msg, success = execute_action(resolved_move_swap["intent"], resolved_move_swap.get("slots", {}), slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()

    if pending_state.get("intent") == "add_slide" and (pending_state.get("missing_slots") or []):
        if re.search(r"\b(?:cancel|stop|never mind|nevermind)\b", routing_prompt, re.IGNORECASE):
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            st.session_state.pending_intent = None
            st.session_state.pending_action = None
            with st.chat_message("assistant"):
                st.markdown("Okay, I cancelled that new slide request.")
            add_message("assistant", "Okay, I cancelled that new slide request.")
            st.rerun()

        slots = dict(pending_state.get("slots", {}))
        # Normalize content keys: prefer `slide_content`, but accept `content` from
        # classifier or earlier flows and copy it so downstream logic is consistent.
        if not slots.get("slide_content") and slots.get("content"):
            slots["slide_content"] = slots.get("content")
        if slots.get("position") is None:
            extracted_position = _extract_add_slide_position(prompt, slides)
            if extracted_position is not None:
                slots["position"] = extracted_position

        if not slots.get("slide_content"):
            content_candidate = _extract_add_slide_content(prompt)
            if content_candidate and not _is_position_only_text(prompt):
                slots["slide_content"] = content_candidate

        resolved_missing = []
        if not slots.get("slide_content"):
            resolved_missing.append("slide_content")
        if slots.get("position") is None:
            resolved_missing.append("position")

        if not resolved_missing:
            result_msg, success = execute_action("add_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            st.rerun()
            
        question = _polish_next_question(prompt, "add_slide", slots, resolved_missing, None)
        st.session_state.pending_intent = {"intent": "add_slide", "slots": slots, "missing_slots": resolved_missing, "next_question": question, "action": "ask"}
        st.session_state.pending_action = st.session_state.pending_intent
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if _is_bare_edit_slide_request(routing_prompt):
        st.session_state.pending_intent = None
        st.session_state.pending_action = None
        st.session_state.active_edit_context = None
        st.session_state.last_slide_index = None
        st.session_state.active_slide_index = None
        if "agent" in st.session_state:
            st.session_state.agent.clear_state()
        followup = f"Which slide would you like to edit in {_ppt_display_label(st.session_state.get('current_ppt_id'))}?"
        with st.chat_message("assistant"):
            st.markdown(followup)
        add_message("assistant", followup)
        st.rerun()

    if pending_state.get("intent") == "edit_slide" and "change_content" in (pending_state.get("missing_slots") or []) and not _is_add_slide_request(routing_prompt):
        if re.search(r"\b(?:cancel|stop|never mind|nevermind)\b", routing_prompt, re.IGNORECASE):
            agent = st.session_state.get("agent") if "agent" in st.session_state else None
            if agent:
                agent.clear_state()
            st.session_state.pending_intent = None
            st.session_state.pending_action = None
            with st.chat_message("assistant"):
                st.markdown("Okay, I cancelled that edit request.")
            add_message("assistant", "Okay, I cancelled that edit request.")
            st.rerun()
        slots = dict(pending_state.get("slots", {}))
        if _is_correction_message(prompt):
            corrected_slide = _extract_explicit_slide_number_from_text(prompt) or _resolve_slide_reference_text(prompt, slides)
            if corrected_slide is not None:
                slots["slide_number"] = corrected_slide
            question = _polish_next_question(prompt, "edit_slide", slots, ["change_content"], None)
            st.session_state.pending_intent = {
                "intent": "edit_slide",
                "slots": slots,
                "missing_slots": ["change_content"],
                "next_question": question,
                "action": "ask",
            }
            st.session_state.pending_action = st.session_state.pending_intent
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            with st.chat_message("assistant"):
                st.markdown(question)
            add_message("assistant", question)
            st.rerun()

        change_content = _extract_edit_change_content(prompt)
        if change_content and not _is_preview_request(routing_prompt) and not _is_topic_ideas_request(routing_prompt):
            slots["change_content"] = change_content
            requested_points = _parse_bullet_add_request(change_content)
            if requested_points and requested_points > _MAX_POINTS_PER_REQUEST:
                allowed_points = _resolved_add_point_count(_MAX_POINTS_PER_REQUEST, _parse_point_side_hint(change_content))
                slots["n_points"] = allowed_points
                question = _point_limit_confirmation_question(
                    int(slots.get("slide_number") or 0),
                    int(requested_points),
                    int(allowed_points),
                    st.session_state.get("current_ppt_id"),
                )
                pending_limit = {
                    "intent": "confirm_add_points_limit",
                    "slots": slots,
                    "requested_points": int(requested_points),
                    "allowed_points": int(allowed_points),
                    "next_question": question,
                    "action": "confirm",
                }
                _store_pending_action(pending_limit)
                with st.chat_message("assistant"):
                    st.markdown(question)
                add_message("assistant", question)
                st.rerun()
            result_msg, success = execute_action("edit_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            st.rerun()

        slide_num = slots.get("slide_number") or _extract_explicit_slide_number_from_text(prompt) or _resolve_slide_reference_text(prompt, slides)
        if slide_num is not None:
            slots["slide_number"] = slide_num
        question = _polish_next_question(prompt, "edit_slide", slots, ["change_content"], None)
        st.session_state.pending_intent = {
            "intent": "edit_slide",
            "slots": slots,
            "missing_slots": ["change_content"],
            "next_question": question,
            "action": "ask",
        }
        st.session_state.pending_action = st.session_state.pending_intent
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    if _is_bare_add_slide_request(routing_prompt) and pending_state.get("intent") != "add_slide":
        slots = {}
        content_candidate = _extract_add_slide_content(prompt)
        position_candidate = _extract_add_slide_position(prompt, slides)
        if content_candidate:
            slots["slide_content"] = content_candidate
        else:
            # Use classifier/reasoned content as a fallback so initial user
            # wording like "add summary slide in ppt 1" isn't lost.
            if reasoned and reasoned.get("content"):
                slots["slide_content"] = reasoned.get("content")
        if position_candidate is not None:
            slots["position"] = position_candidate

        missing_slots = []
        if not slots.get("slide_content"):
            missing_slots.append("slide_content")
        if slots.get("position") is None:
            missing_slots.append("position")

        if not missing_slots:
            result_msg, success = execute_action("add_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()

        st.session_state.pending_intent = {
            "intent": "add_slide",
            "slots": slots,
            "missing_slots": missing_slots,
            "next_question": None,
            "action": "ask",
        }
        st.session_state.pending_action = st.session_state.pending_intent
        st.session_state.active_edit_context = None
        if "agent" in st.session_state:
            st.session_state.agent.clear_state()
        followup = _polish_next_question(prompt, "add_slide", slots, missing_slots, None)
        with st.chat_message("assistant"):
            st.markdown(followup)
        add_message("assistant", followup)
        st.rerun()

    # If the user names a deck in an action request, make it the active context first.
    named_changed, named_ppt_id, named_clarification = resolve_named_ppt_context(routing_prompt)
    if named_clarification:
        with st.chat_message("assistant"):
            st.markdown(named_clarification)
        add_message("assistant", named_clarification)
        st.rerun()

    # 1. Resolve PPT reference (possibly switch active PPT)
    def should_resolve_ppt(user_input: str) -> bool:
        return bool(re.search(r"(ppt|presentation|deck)\s*\d+|switch|open|go to|previous ppt", user_input.lower()))
    if should_resolve_ppt(routing_prompt):
        changed, new_ppt_id, clarification = resolve_ppt_reference(routing_prompt)
    else:
        changed, new_ppt_id, clarification = False, None, None
    if named_changed:
        changed = True
        new_ppt_id = named_ppt_id
    if clarification:
        with st.chat_message("assistant"):
            st.markdown(clarification)
        add_message("assistant", clarification)
        st.rerun()
    if changed and _is_deck_level_edit_request(prompt):
        st.session_state.active_edit_context = {
            "level": "ppt",
            "ppt_id": new_ppt_id,
        }
        st.session_state.action_target_ppt_id = new_ppt_id
        ppt_label = _ppt_display_label(new_ppt_id)
        followup = f"What changes would you like to make in {ppt_label}?"
        with st.chat_message("assistant"):
            st.markdown(followup)
        add_message("assistant", followup)
        st.rerun()
    if changed:
        st.session_state.action_target_ppt_id = new_ppt_id
        if re.search(r"\b(?:open|switch to|go to|show)\b", routing_prompt, re.IGNORECASE):
            active_item = get_ppt_by_id(new_ppt_id) if new_ppt_id else None
            if active_item and active_item.get("outline_payload"):
                preview_slides = active_item["outline_payload"].get("slides", [])
                topic = active_item.get("topic", "Presentation")
                label = f"Preview — {topic} ({new_ppt_id})"
                add_message_with_preview(
                    role="assistant",
                    text=f"Showing preview for **{new_ppt_id}** — {topic}.",
                    preview_slides=preview_slides,
                    ppt_label=label,
                    ppt_bytes=active_item.get("ppt_bytes"),
                    ppt_filename=active_item.get("ppt_filename", "presentation.pptx"),
                )
                st.rerun()

        active_label = _ppt_display_label(new_ppt_id)
        active_topic = st.session_state.get("topic", "").strip()
        with st.chat_message("assistant"):
            message = f"✅ Active deck set to **{active_label}**."
            if active_topic:
                message += f" Topic: {active_topic}."
                message += " You can now say things like `edit slide 3`, `add a new slide at the end`, or `show preview`."
            st.markdown(message)
        add_message("assistant", f"Active deck set to {active_label}.")
    elif changed:
        st.session_state.active_edit_context = None
        if not _looks_like_change_instruction(prompt):
            st.rerun()

    # Directly answer topic lookup requests so they do not fall into suggest_topic.
    if _is_ppt_topic_lookup_request(routing_prompt):
        history = st.session_state.get("ppt_history", [])
        explicit_ref = _extract_explicit_ppt_ref_from_text(routing_prompt)
        resolved_item = None
        if explicit_ref:
            resolved_item = _resolve_ppt_history_item(explicit_ref)
        elif len(history) == 1:
            resolved_item = history[0]

        if resolved_item:
            topic = resolved_item.get("topic", "").strip() or "Untitled presentation"
            ppt_id = resolved_item.get("id", "ppt")
            reply = f"The topic for {_ppt_display_label(ppt_id)} is: {topic}."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()

        if len(history) > 1:
            topics = [
                f"**{i}.** {item.get('id', f'ppt_{i}')}: {item.get('topic', 'Untitled presentation')}"
                for i, item in enumerate(history, start=1)
            ]
            reply = "Here are the topics of your generated presentations:\n\n" + "\n".join(topics)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()

    if _is_preview_request(routing_prompt):
        current_id = st.session_state.get("current_ppt_id")
        current_item = get_ppt_by_id(current_id) if current_id else None
        if current_item and current_item.get("outline_payload"):
            preview_slides = current_item["outline_payload"].get("slides", [])
            topic = current_item.get("topic", "Presentation")
            label = f"Preview — {topic} ({current_id})"
            add_message_with_preview(
                role="assistant",
                text=f"Showing preview for **{current_id}** — {topic}.",
                preview_slides=preview_slides,
                ppt_label=label,
                ppt_bytes=current_item.get("ppt_bytes"),
                ppt_filename=current_item.get("ppt_filename", "presentation.pptx"),
            )
            st.rerun()

    view_request = _extract_view_slide_request(routing_prompt)
    if view_request:
        ppt_ref = view_request.get("ppt_ref")
        if ppt_ref:
            changed, new_ppt_id, clarification = resolve_ppt_reference(str(ppt_ref))
            if clarification:
                with st.chat_message("assistant"):
                    st.markdown(clarification)
                add_message("assistant", clarification)
                st.rerun()
            if changed and new_ppt_id:
                st.session_state.action_target_ppt_id = new_ppt_id
        current_outline = _current_outline_payload()
        slides_local = current_outline.get("slides", []) if current_outline else []
        slide_num = view_request.get("slide_number")
        if slide_num is not None and current_outline and 1 <= int(slide_num) <= len(slides_local):
            slide_idx = int(slide_num)
            slide = slides_local[slide_idx - 1] if isinstance(slides_local[slide_idx - 1], dict) else {}
            ppt_label = _ppt_display_label(st.session_state.get("current_ppt_id"))
            reply = _format_slide_view(slide, slide_idx, ppt_label)
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()
        slide_label = slide_num if slide_num is not None else "?"
        ppt_label = _ppt_display_label(st.session_state.get("current_ppt_id"))
        reply = f"Slide {slide_label} does not exist in {ppt_label}. This presentation only has {len(slides_local)} slide(s)."
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.rerun()

    if _is_transform_request(routing_prompt):
        slide_num = _extract_explicit_slide_number_from_text(routing_prompt) or _resolve_slide_reference_text(routing_prompt, slides)
        if slide_num is not None:
            slots = {"slide_number": slide_num, "change_content": routing_prompt}
            result_msg, success = execute_action("edit_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()

    slide_target = _extract_slide_navigation_target(routing_prompt)
    current_outline = _current_outline_payload()
    if slide_target and current_outline:
        slides = current_outline.get("slides", [])
        if 1 <= slide_target <= len(slides):
            slide = slides[slide_target - 1] if isinstance(slides[slide_target - 1], dict) else {}
            slide_title = str(slide.get("title", f"Slide {slide_target}")).strip() or f"Slide {slide_target}"
            _remember_action_context("focus_slide", slide_target, st.session_state.get("current_ppt_id"), f"Focused on slide {slide_target}")
            reply = f"Focused on slide {slide_target}: {slide_title}."
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
            st.rerun()

    if (
        re.search(r"\bimprove\b", routing_prompt, re.IGNORECASE)
        and re.search(r"\bcontent\b", routing_prompt, re.IGNORECASE)
        and not _is_slide_level_request(routing_prompt)
        and not _extract_deck_refinement_request(routing_prompt)
        and current_outline
    ):
        target_slide = st.session_state.get("active_slide_index") or st.session_state.get("last_slide_index")
        if target_slide:
            slots = {"slide_number": target_slide, "change_content": prompt}
            result_msg, success = execute_action("edit_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()
        question = f"Which slide should I improve in {_ppt_display_label(st.session_state.get('current_ppt_id'))}?"
        with st.chat_message("assistant"):
            st.markdown(question)
        add_message("assistant", question)
        st.rerun()

    refinement = _extract_deck_refinement_request(routing_prompt)
    if refinement and current_outline and not _is_slide_level_request(routing_prompt):
        current_id = st.session_state.get("current_ppt_id")
        current_item = get_ppt_by_id(current_id) if current_id else None
        if current_item:
            style_hint = refinement.get("style_hint") or routing_prompt
            if re.search(r"\b(shorter|less text|concise|condense|simpler|remove fluff|key points only|important points only)\b", routing_prompt, re.IGNORECASE):
                updated_slides = [_shorten_slide_structure(s) for s in slides]
                commit_changes(updated_slides, "Tightened the deck content while preserving the existing slide structure.")
                st.stop()
            updated_slides = _apply_design_to_existing_slides(slides, style_hint, refinement.get("tone") or "")
            commit_changes(updated_slides, "Updated the deck styling while preserving existing slide content and structure.")
            st.stop()

    # 🔧 FIX 3: Override for "add point in slide X" – force edit_slide
    if re.search(r"\b(add|insert)\s+\d*\s*(bullet|point|points)\b.*\bslide\s+\d+\b", routing_prompt, re.IGNORECASE):
        slide_num_match = re.search(r"\bslide\s+(\d+)\b", routing_prompt, re.IGNORECASE)
        if slide_num_match:
            slide_num = int(slide_num_match.group(1))
            # Clear any pending add_slide state
            st.session_state.pending_intent = None
            st.session_state.pending_action = None
            if "agent" in st.session_state:
                st.session_state.agent.clear_state()
            # Execute edit_slide directly
            slots = {"slide_number": slide_num, "change_content": routing_prompt}
            result_msg, success = execute_action("edit_slide", slots, slides)
            with st.chat_message("assistant"):
                st.markdown(result_msg)
            add_message("assistant", result_msg)
            st.rerun()

    # Prepare context for the agent
    context = {
        "has_deck": bool(outline_payload),
        "slides": slides,
    }

    # Initialize agent
    if "agent" not in st.session_state:
        st.session_state.agent = ConversationalAgent()
    agent = st.session_state.agent

    # Process with agent (intent detection + slot filling)
    response_text, action_data, should_execute = agent.process(routing_prompt, context)

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
