"""
intelligence_layer.py  (v3 — strict routing, no add_slide fallback)
====================================================================

Drop-in replacement for the intelligence layer in app.py.

KEY CHANGES vs v2:
  - reasoning_layer() applies hard semantic guards BEFORE touching intents
  - _should_execute_directly() has tighter per-intent rules
  - _confidence_preface() only fires when scope was inferred (not explicit)
  - _natural_success_message() unchanged API
  - New helper _safe_intent() prevents routing surprises
  - CRITICAL: unknown / ambiguous intent → ask clarification, NOT add_slide
"""

from __future__ import annotations

import re
from typing import Optional

import streamlit as st

from backend.services.intent_classifier import classify_ppt_intent
from backend.services.response_generator import generate_response


# ─────────────────────────────────────────────────────────────────────────────
#  Utilities
# ─────────────────────────────────────────────────────────────────────────────

def _norm_text(v: object) -> str:
    s = str(v or "").casefold()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _safe_str(v: object, default: str = "") -> str:
    return str(v).strip() if v is not None else default


_NUMBER_WORDS = {
    "one": 1, "a": 1, "an": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
    "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19,
    "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
}


def _extract_requested_point_count(text: str) -> Optional[int]:
    m = re.search(
        r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|"
        r"twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|"
        r"nineteen|twenty|thirty|forty|fifty)\s*(?:more\s+)?"
        r"(?:point|points|bullet|bullets|item|items)\b",
        text or "",
        re.IGNORECASE,
    )
    if not m:
        return None
    token = m.group(1).lower()
    return int(token) if token.isdigit() else _NUMBER_WORDS.get(token)


def _current_outline_payload() -> dict:
    outline = st.session_state.get("outline_payload")
    return outline if isinstance(outline, dict) else {}


def _ppt_display_label(ppt_id: str) -> str:
    if not ppt_id:
        return "ppt"
    m = re.search(r"ppt_(\d+)", str(ppt_id), re.IGNORECASE)
    return f"ppt {m.group(1)}" if m else str(ppt_id).replace("_", " ")


def _extract_explicit_slide_number(user_input: str) -> Optional[int]:
    for pat in (
        r"\bslide[_\s-]*#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
    ):
        m = re.search(pat, user_input or "", re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                pass
    return None


def _extract_explicit_ppt_ref(user_input: str) -> Optional[str]:
    for pat in (
        r"\b(?:ppt|presentation|deck)[_\s-]*#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck)\b",
    ):
        m = re.search(pat, user_input or "", re.IGNORECASE)
        if m:
            return f"ppt {m.group(1)}"
    return None


# Intents that operate on an existing slide
_SLIDE_INTENTS = {
    "transform_content", "add_points", "update_slide", "edit_slide",
    "view_slide", "explain_slide", "delete_slide", "merge_slides",
    "move_slide", "swap_slides", "clear_slide", "blank_slide",
}

# Intents that MUST NOT be silently redirected to add_slide
_NEVER_ADD_SLIDE = _SLIDE_INTENTS | {
    "preview_ppt", "switch_ppt", "download_ppt", "ppt_info",
    "suggest_topic", "refine_ppt", "create_ppt",
    "cancel", "greeting", "smalltalk", "unknown",
}


def _safe_intent(intent: str, fallback: str = "unknown") -> str:
    """Return intent unchanged, but guarantee we never smuggle in add_slide."""
    if not intent or intent not in {
        "transform_content", "add_points", "update_slide", "edit_slide",
        "view_slide", "explain_slide", "delete_slide", "merge_slides",
        "move_slide", "swap_slides", "clear_slide", "blank_slide",
        "add_slide", "preview_ppt", "switch_ppt", "download_ppt",
        "create_ppt", "ppt_info", "suggest_topic", "refine_ppt",
        "cancel", "greeting", "smalltalk", "unknown",
    }:
        return fallback
    return intent


# ─────────────────────────────────────────────────────────────────────────────
#  1. REASONING LAYER
# ─────────────────────────────────────────────────────────────────────────────

def reasoning_layer(user_input: str, state: dict) -> dict:
    """
    Unified conversational reasoning layer.

    Priority of overrides (highest wins):
      cancel > clear/blank > view_slide (explicit number) >
      transform_content > add_points/expand > add_slide (explicit) >
      context-continuation > LLM result
    """
    state  = state if isinstance(state, dict) else {}
    text   = (user_input or "").strip()
    lower  = _norm_text(text)

    current_ppt_id   = state.get("current_ppt_id")
    current_slide_id = state.get("current_slide_id")
    last_intent      = _safe_str(state.get("last_intent"))
    last_action      = _safe_str(state.get("last_action"))

    # ── Step 1: LLM-first classification ────────────────────────────────────
    base = classify_ppt_intent(text, ppt_id=current_ppt_id, slide_id=current_slide_id)

    result: dict = {
        "intent":                 _safe_intent(base.get("intent", "unknown")),
        "slide_id":               base.get("slide_id"),
        "slide_ids":              base.get("slide_ids", []),
        "ppt_id":                 base.get("ppt_id"),
        "operation":              base.get("operation"),
        "content":                base.get("content"),
        "action_type":            base.get("action_type"),
        "target":                 base.get("target"),
        "sub_intent":             base.get("sub_intent"),
        "missing_entities":       base.get("missing_entities", []),
        "clarification_question": base.get("clarification_question"),
        "position":               base.get("position"),
        "anchor_slide":           base.get("anchor_slide"),
        "confidence":             float(base.get("confidence", 0.30) or 0.30),
    }

    # Helper: override result cleanly
    def _override(
        intent: str,
        *,
        operation: Optional[str] = None,
        slide_id: Optional[int] = None,
        ppt_id: Optional[str] = None,
        content: Optional[str] = None,
        confidence: float = 0.95,
        sub_intent: Optional[str] = None,
    ) -> None:
        result["intent"]     = _safe_intent(intent)
        result["operation"]  = operation or intent
        result["confidence"] = confidence
        if slide_id  is not None: result["slide_id"]  = slide_id
        if ppt_id    is not None: result["ppt_id"]    = ppt_id
        if content   is not None: result["content"]   = content
        if sub_intent is not None: result["sub_intent"] = sub_intent

    # ── Step 2: Extract explicit references ──────────────────────────────────
    explicit_slide = _extract_explicit_slide_number(text)
    explicit_ppt   = _extract_explicit_ppt_ref(text)

    if explicit_ppt:
        result["ppt_id"] = explicit_ppt

    # For add_slide, an anchor slide ref is NOT the target slide
    add_slide_anchor = (
        result["intent"] == "add_slide"
        and re.search(r"\b(?:after|before)\s+slide\s*#?\d+\b", text, re.IGNORECASE)
    )
    if explicit_slide is not None and not add_slide_anchor:
        result["slide_id"] = explicit_slide

    if result["intent"] == "add_slide":
        result["slide_id"] = None

    # ── Step 3: Hard semantic overrides ─────────────────────────────────────

    # CANCEL — absolute priority
    if re.fullmatch(r"(?:stop|cancel|nevermind|never\s*mind|abort|quit)", lower):
        _override("cancel", operation="cancel", confidence=0.99)
        return result

    # CLEAR / BLANK
    if re.search(r"\b(?:remove everything|clear slide|clear all|erase all|wipe slide|delete all content|remove all content)\b", text, re.IGNORECASE):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override("clear_slide", operation="clear", slide_id=target, confidence=0.97)
        return result

    if re.search(r"\b(?:make blank slide|blank slide|empty slide)\b", text, re.IGNORECASE):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override("blank_slide", operation="clear", slide_id=target, confidence=0.97)
        return result

    # VIEW SLIDE — explicit number present (highest priority before transforms)
    view_match = re.search(
        r"\b(?:show|view|display|see|go\s+to|jump\s+to|navigate\s+to)\s+slide\s*#?(\d+)\b"
        r"|\bslide\s*#?(\d+)\s+(?:show|view|display|see)\b",
        text, re.IGNORECASE,
    )
    if view_match:
        num_str = next((g for g in view_match.groups() if g), None)
        slide_num = int(num_str) if num_str else explicit_slide
        _override("view_slide", operation="view", slide_id=slide_num, confidence=0.99)
        return result

    # EXPLAIN SLIDE
    if re.search(
        r"\b(?:explain|describe|summarize|summary\s+of|walk\s+me\s+through)\b.*\bslide\s*#?\d+\b"
        r"|\bslide\s*#?\d+\b.*\b(?:explain|describe|summarize|about)\b",
        text, re.IGNORECASE,
    ):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override("explain_slide", operation="explain", slide_id=target, confidence=0.97)
        return result

    # DELETE SLIDE
    if re.search(r"\b(?:delete|remove)\s+slide\s*#?\d+\b|\bslide\s*#?\d+\b.*\b(?:delete|remove)\b", text, re.IGNORECASE):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override("delete_slide", operation="delete", slide_id=target, confidence=0.97)
        return result

    # MERGE SLIDES
    if re.search(r"\b(?:merge|combine|join|consolidate)\s+slide", text, re.IGNORECASE):
        _override("merge_slides", operation="merge", confidence=0.96)
        return result

    # TRANSFORM — high priority
    if re.search(
        r"\b(?:improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|condense|"
        r"make\s+it\s+better|make\s+it\s+shorter|make\s+it\s+professional|make\s+it\s+cleaner|"
        r"make\s+it\s+clearer|make\s+it\s+concise|spice\s+up|jazz\s+up|tighten|upgrade|rephrase|"
        r"make\s+more\s+professional|make\s+more\s+concise|clean\s+up|fix\s+this|fix\s+it|"
        r"improve\s+this|improve\s+it|make\s+this\s+better|make\s+it\s+sound\s+better|"
        r"make\s+it\s+engaging|make\s+it\s+impactful|make\s+it\s+punchy|make\s+it\s+snappy)\b",
        text, re.IGNORECASE,
    ):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            "transform_content",
            operation="transform",
            slide_id=target,
            content=text,
            confidence=0.97,
        )
        return result

    # ADD POINTS / EXPAND
    if re.search(
        r"\b(?:add|append|insert)\b.*\b(?:point|points|bullet|bullets)\b"
        r"|\bmore\s+points?\b|\badd\s+one\s+more\b|\badd\s+another\s+point\b"
        r"|\b(?:make|expand|elaborate)\b.*\bslide\s*#?\d+\b.*\b(?:longer|more\s+detailed|more\s+detail)\b"
        r"|\bslide\s*#?\d+\b.*\b(?:longer|more\s+detailed|more\s+detail|expand)\b",
        text, re.IGNORECASE,
    ):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            "add_points",
            operation="append",
            slide_id=target,
            content=text,
            sub_intent="add_point",
            confidence=0.96,
        )
        return result

    # CONTEXT CONTINUATION ("more", "another", "continue" …)
    continuation_re = re.fullmatch(
        r"(?:add\s+)?(?:one\s+)?more|another|continue|carry\s+on|"
        r"add\s+one\s+more|keep\s+going|yes\s+more|one\s+more\s+point",
        lower,
    )
    if (
        continuation_re
        and last_intent in {"add_points", "edit_slide", "transform_content", "add_slide"}
        and not re.search(r"\b(?:show|view|display|see|go\s+to|jump\s+to)\s+slide\b", text, re.IGNORECASE)
    ):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            last_intent,
            operation="append" if last_intent == "add_points" else "modify",
            slide_id=target,
            content=last_action or text,
            confidence=0.88,
        )
        return result

    # VAGUE FOLLOW-UP ("this", "it", "that")
    if (
        re.fullmatch(r"(?:this|it|that|here|go\s+ahead|do\s+it)", lower)
        and last_intent
        and current_slide_id
        and not re.search(r"\b(?:show|view|display|see|go\s+to|jump\s+to)\s+slide\b", text, re.IGNORECASE)
    ):
        _override(
            last_intent,
            operation="append" if last_intent == "add_points" else "modify",
            slide_id=current_slide_id,
            content=last_action or text,
            confidence=0.82,
        )
        return result

    # ── Step 4: Scope resolution — fill missing context ──────────────────────
    if result["slide_id"] is None and current_slide_id and result["intent"] in _SLIDE_INTENTS:
        result["slide_id"]  = current_slide_id
        result["confidence"] = max(0.60, result["confidence"] - 0.10)

    if result["ppt_id"] is None and current_ppt_id:
        result["ppt_id"] = current_ppt_id

    # ── Step 5: Confidence floor for resolved slide intents ──────────────────
    if result["intent"] in _SLIDE_INTENTS and result["slide_id"] is not None:
        result["confidence"] = max(result["confidence"], 0.72)

    # ── Step 6: Safety net — unknown must not morph into add_slide ───────────
    if result["intent"] == "unknown":
        result["clarification_question"] = (
            result.get("clarification_question")
            or "What would you like me to do with the presentation?"
        )

    return result


# ─────────────────────────────────────────────────────────────────────────────
#  2. CANONICALIZE REASONED PROMPT
# ─────────────────────────────────────────────────────────────────────────────

def _canonicalize_reasoned_prompt(user_input: str, reasoned: dict) -> str:
    """Build a clean internal routing string from the reasoned intent."""
    text     = (user_input or "").strip()
    intent   = _safe_str((reasoned or {}).get("intent"))
    slide_id = (reasoned or {}).get("slide_id")
    ppt_id   = _safe_str((reasoned or {}).get("ppt_id"))
    content  = _safe_str((reasoned or {}).get("content"))

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
        return (f"show slide {slide_id} of {ppt_id}" if ppt_id else f"show slide {slide_id}")
    if intent == "explain_slide" and slide_id is not None:
        return f"explain slide {slide_id}"
    if intent == "add_points" and slide_id is not None:
        n = _extract_requested_point_count(content or text) or 1
        return f"add {n} point(s) in slide {slide_id}"
    if intent == "edit_slide" and slide_id is not None and content:
        return f"edit slide {slide_id}: {content}"
    if intent == "delete_slide" and slide_id is not None:
        return f"delete slide {slide_id}"
    if intent == "merge_slides":
        ids = (reasoned or {}).get("slide_ids") or []
        if len(ids) >= 2:
            return f"merge slide {ids[0]} and {ids[1]}"
    return text


# ─────────────────────────────────────────────────────────────────────────────
#  3. POLISH NEXT QUESTION
# ─────────────────────────────────────────────────────────────────────────────

def _polish_next_question(
    user_input: str,
    intent: Optional[str],
    slots: dict,
    missing: list,
    next_question: Optional[str],
) -> str:
    """Return a natural, context-specific follow-up question."""
    slide_num     = slots.get("slide_number")
    slide_content = slots.get("slide_content")
    ppt_id        = st.session_state.get("current_ppt_id") or slots.get("ppt_ref")
    ppt_label     = _ppt_display_label(str(ppt_id)) if ppt_id else "the presentation"
    missing       = missing or []

    def _slide_ref(num: object) -> str:
        try:
            outline = _current_outline_payload()
            slides  = (outline or {}).get("slides", [])
            n = int(num)  # type: ignore[arg-type]
            if 1 <= n <= len(slides):
                title = slides[n - 1].get("title", "")
                if title:
                    return f"**{title}** (slide {n})"
        except Exception:
            pass
        return f"slide {num}" if num else "that slide"

    if intent == "edit_slide":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide would you like to edit in {ppt_label}?"
        if "change_content" in missing:
            return (
                f"What should I change in {_slide_ref(slide_num)}? "
                f"(e.g. rewrite a bullet, update the title, add a point)"
            )

    if intent in {"clear_slide", "blank_slide"}:
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I clear in {ppt_label}?"
        if intent == "clear_slide":
            return f"This will remove everything from slide {slide_num}. Do you want me to continue?"
        return f"Should I make slide {slide_num} blank in {ppt_label}?"

    if intent == "transform_content":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I improve in {ppt_label}?"
        return (
            f"How should I improve {_slide_ref(slide_num)}? "
            f"(e.g. make it shorter, rewrite, more professional)"
        )

    if intent == "add_points":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I add points to in {ppt_label}?"

    if intent == "update_slide":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I update in {ppt_label}?"
        if "change_content" in missing:
            return f"What should I change in {_slide_ref(slide_num)}?"

    if intent == "explain_slide":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I explain in {ppt_label}?"

    if intent == "view_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide would you like to see in {ppt_label}?"

    if intent == "delete_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide should I delete from {ppt_label}?"

    if intent == "merge_slides":
        return "Which two slides should I merge?"

    if intent == "add_slide":
        if "slide_content" in missing or slide_content is None:
            return f"What topic should the new slide cover in {ppt_label}?"
        if "position" in missing:
            return (
                "Where should I place the new slide? "
                "(e.g. at the end, after slide 3, at the start)"
            )

    if intent == "move_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide would you like to move in {ppt_label}?"
        num = slots.get("slide_number")
        if "anchor_slide" in missing and "position" not in missing:
            return (
                f"Where should {_slide_ref(num)} go? "
                f"(e.g. before slide 4, after slide 6, to position 2)"
            )
        if "position" in missing:
            return f"What position should {_slide_ref(num)} move to?"

    if intent == "swap_slides":
        if "slide_a" in missing or slots.get("slide_a") is None:
            return f"Which two slides should I swap in {ppt_label}?"
        if "slide_b" in missing or slots.get("slide_b") is None:
            return f"Swap slide {slots.get('slide_a')} with which other slide?"

    if intent == "unknown":
        return next_question or "What would you like me to do with the presentation?"

    # Clean up generic LLM wording
    q = (next_question or "").strip()
    if q:
        q = re.sub(r"(?i)\bwhat point\b", "what bullet point", q)
        q = re.sub(r"(?i)\bwhat points\b", "what bullet points", q)
        q = re.sub(r"(?i)what change would you like to make\??", "", q).strip()
        if q:
            return q

    return "Could you give me a bit more detail?"


# ─────────────────────────────────────────────────────────────────────────────
#  4. CONFIDENCE-BASED EXECUTION GATE
# ─────────────────────────────────────────────────────────────────────────────

def _should_execute_directly(reasoned: dict) -> bool:
    """
    Return True when confidence is high enough to act without asking.

    confidence ≥ 0.60 → execute (possibly with a scope note)
    confidence  < 0.60 → ask for clarification
    """
    reasoned  = reasoned or {}
    intent    = _safe_str(reasoned.get("intent"))
    slide_id  = reasoned.get("slide_id")
    content   = _safe_str(reasoned.get("content"))
    conf      = float(reasoned.get("confidence") or 0.0)

    # unknown → never execute
    if intent == "unknown":
        return False

    # Slide-level intents require a slide_id
    if intent in _SLIDE_INTENTS and slide_id is None:
        return False

    # Intents that need specific sub-slots
    if intent == "edit_slide":
        return bool(slide_id and content) and conf >= 0.60

    if intent == "update_slide":
        return bool(slide_id and content) and conf >= 0.60

    if intent == "add_slide":
        has_placement = bool(reasoned.get("position") or reasoned.get("anchor_slide"))
        return bool(content and has_placement) and conf >= 0.60

    if intent in {"delete_slide", "move_slide", "swap_slides", "view_slide", "explain_slide"} and not slide_id:
        return False

    if intent == "merge_slides" and len(reasoned.get("slide_ids") or []) < 2:
        return False

    return conf >= 0.60


# ─────────────────────────────────────────────────────────────────────────────
#  5. CONFIDENCE PREFACE
# ─────────────────────────────────────────────────────────────────────────────

def _confidence_preface(reasoned: dict) -> Optional[str]:
    """
    For confidence 0.60–0.79, return a brief note when scope was inferred.
    Returns None when confidence is high (≥ 0.80) or not applicable.
    """
    conf     = float((reasoned or {}).get("confidence") or 1.0)
    intent   = (reasoned or {}).get("intent", "")
    slide_id = (reasoned or {}).get("slide_id")

    if conf >= 0.80:
        return None

    if intent in {"transform_content", "add_points", "update_slide", "edit_slide"} and slide_id:
        return f"*(Working on slide {slide_id} — let me know if that's not right.)*"

    return None


# ─────────────────────────────────────────────────────────────────────────────
#  6. NATURAL SUCCESS MESSAGES
# ─────────────────────────────────────────────────────────────────────────────

def _natural_success_message(intent: str, slots: dict) -> str:
    """Return a controlled but varied confirmation after an action completes."""
    slide_num = slots.get("slide_number") or slots.get("slide_id")
    n         = slots.get("n_points", 1)
    ppt_id    = st.session_state.get("current_ppt_id", "")
    ppt_label = _ppt_display_label(ppt_id) if ppt_id else "the deck"
    current_ppt = st.session_state.get("current_ppt")
    ppt_topic = st.session_state.get("topic") or (
        current_ppt.get("topic") if isinstance(current_ppt, dict) else ""
    )

    slide_title = ""
    if slide_num:
        try:
            outline = _current_outline_payload()
            slides  = (outline or {}).get("slides", [])
            n_int   = int(slide_num)
            if 1 <= n_int <= len(slides):
                slide_title = slides[n_int - 1].get("title", "")
        except Exception:
            pass

    client = None
    model = None
    try:
        from openai import AzureOpenAI
        from backend.config import AZURE_API_VERSION, AZURE_DEPLOYMENT, AZURE_ENDPOINT, AZURE_KEY

        client = AzureOpenAI(
            api_key=AZURE_KEY,
            api_version=AZURE_API_VERSION,
            azure_endpoint=AZURE_ENDPOINT,
        )
        model = AZURE_DEPLOYMENT
    except Exception:
        client = None
        model = None

    return generate_response(
        intent,
        slide_number=int(slide_num) if slide_num is not None else None,
        slide_title=slide_title,
        ppt_topic=str(ppt_topic or ""),
        ppt_label=ppt_label,
        change_type=str(slots.get("change_type") or slots.get("sub_intent") or ""),
        change_content=str(slots.get("change_content") or slots.get("content") or ""),
        n_points=int(n) if n is not None else None,
        user_request=str(slots.get("user_request") or ""),
        llm_client=client,
        llm_model=model,
    )
