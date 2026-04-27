"""
intelligence_layer.py
=====================
 
Drop-in replacements and additions for the intelligence section of app.py.
 
WHAT CHANGED (nothing else in app.py is touched):
  • reasoning_layer()                — LLM-first, context-aware, priority-correct
  • _canonicalize_reasoned_prompt()  — cleaner internal routing instructions
  • _polish_next_question()          — natural, slide-title-aware follow-ups
  • _should_execute_directly()       — NEW: confidence gate (replaces hardcoded flow)
  • _confidence_preface()            — NEW: subtle scope-inference notice
  • _natural_success_message()       — NEW: human-sounding confirmations
 
HOW TO INTEGRATE
----------------
1. Copy intent_classifier_final.py  →  backend/services/intent_classifier.py
2. In app.py, replace the four existing functions with the versions below.
3. Add _should_execute_directly, _confidence_preface, _natural_success_message
   anywhere before the "ACTION EXECUTORS" section.
4. In the main chat loop, update the three blocks that call reasoning_layer()
   to also call _should_execute_directly() and _natural_success_message().
   See the INTEGRATION GUIDE at the bottom of this file.
 
No FastAPI routes, file-upload logic, ppt_service, or UI structure is changed.
"""
 
import re
from typing import Optional

import streamlit as st

from backend.services.intent_classifier import classify_ppt_intent
 
def _norm_text(v):
    """Normalize text for fuzzy matching."""
    s = str(v or "").casefold()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def _safe_str(v, default=""):
    return str(v).strip() if v is not None else default


def _current_outline_payload() -> dict:
    outline = st.session_state.get("outline_payload")
    return outline if isinstance(outline, dict) else {}


def _ppt_display_label(ppt_id: str) -> str:
    if not ppt_id:
        return "ppt"
    m = re.search(r"ppt_(\d+)", str(ppt_id), re.IGNORECASE)
    if m:
        return f"ppt {m.group(1)}"
    return str(ppt_id).replace("_", " ")


def _extract_explicit_slide_number_from_text(user_input: str) -> Optional[int]:
    text = user_input or ""
    if not text:
        return None
    patterns = [
        r"\bslide[_\s-]*#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
        r"\bslide\s+(\d+)\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                return None
    return None


def _extract_explicit_ppt_ref_from_text(user_input: str) -> Optional[str]:
    text = user_input or ""
    if not text:
        return None
    patterns = [
        r"\b(?:ppt|presentation|deck)[_\s-]*#?(\d+)\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck)\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return f"ppt {m.group(1)}"
    return None
 
 
# ═══════════════════════════════════════════════════════════════════════════
#  1. REASONING LAYER
# ═══════════════════════════════════════════════════════════════════════════
 
def reasoning_layer(user_input: str, state: dict) -> dict:
    """
    Unified conversational reasoning layer.
 
    - Uses the LLM-backed classify_ppt_intent() for semantic understanding.
    - Applies hard overrides for high-confidence patterns (transform, add-points …).
    - Resolves scope from conversation context when the user omits slide/ppt refs.
    - Returns a clean dict used by the downstream router.
 
    Intent priority (hardcoded):
      transform_content > add_points > edit_slide > add_slide > …
    """
    state = state if isinstance(state, dict) else {}
    text  = (user_input or "").strip()
    lower = _norm_text(text)
 
    # Pull conversation context
    current_ppt_id   = state.get("current_ppt_id")
    current_slide_id = state.get("current_slide_id")
    last_intent      = _safe_str(state.get("last_intent", ""))
    last_action      = _safe_str(state.get("last_action", ""))
 
    # ── Step 1: Semantic classification (LLM-first, regex-fallback) ─────────
    base = classify_ppt_intent(
        text,
        ppt_id=current_ppt_id,
        slide_id=current_slide_id,
    )
 
    result = {
        "intent":     base.get("intent", "general_request"),
        "slide_id":   base.get("slide_id"),
        "ppt_id":     base.get("ppt_id"),
        "operation":  base.get("operation"),
        "content":    base.get("content"),
        "confidence": float(base.get("confidence", 0.35) or 0.35),
    }
 
    def _override(intent, *, operation=None, slide_id=None,
                  ppt_id=None, content=None, confidence=0.92):
        result["intent"]    = intent
        result["operation"] = operation or intent
        if slide_id  is not None: result["slide_id"]  = slide_id
        if ppt_id    is not None: result["ppt_id"]    = ppt_id
        if content   is not None: result["content"]   = content
        result["confidence"] = confidence

    # ── Step 2: Extract explicit references from the raw message ─────────────
    explicit_slide = _extract_explicit_slide_number_from_text(text)
    explicit_ppt   = _extract_explicit_ppt_ref_from_text(text)
 
    if explicit_ppt   is not None: result["ppt_id"]   = explicit_ppt
    if explicit_slide is not None: result["slide_id"] = explicit_slide
 
    # ── Step 3: Hard semantic overrides (keyword patterns win) ───────────────
 
    # Correction messages: "no, I meant slide 3" → update slide, keep intent
    if re.search(
        r"\b(?:sorry|no|nope|actually|wait|correction|instead|meant|i mean)\b",
        text, re.IGNORECASE,
    ):
        if explicit_slide is not None:
            result["slide_id"]  = explicit_slide
            result["confidence"] = max(result["confidence"], 0.95)

    # TRANSFORM — highest priority, always wins
    _TRANSFORM_PATT = re.compile(
        r"\b(?:improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|"
        r"condense|make it better|make it shorter|make it professional|"
        r"make it cleaner|make it clearer|make it concise|spice up|jazz up|"
        r"tighten|upgrade|rephrase|make more professional|make more concise|"
        r"make more clear|make it flow|clean up|clean it up|fix this|fix it|"
        r"improve this|improve it|make this better|make it sound better|"
        r"make it engaging|make it impactful|make it punchy|make it snappy)\b",
        re.IGNORECASE,
    )
    if _TRANSFORM_PATT.search(text):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            "transform_content",
            operation="transform",
            slide_id=target,
            content=text,
            confidence=0.97,
        )
 
    # ADD POINTS — second priority
    elif re.search(
        r"\b(?:add|append|insert)\b.*\b(?:point|points|bullet|bullets)\b"
        r"|\bmore points?\b|\badd one more\b|\badd another point\b",
        text, re.IGNORECASE,
    ):
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            "add_points",
            operation="append",
            slide_id=target,
            content=text,
            confidence=0.95,
        )
 
    # VIEW SLIDE
    elif re.search(
        r"\b(?:show|view|display|see)\s+slide\s+\d+\b"
        r"|\bslide\s+\d+\b.*\b(?:show|view|display)\b",
        text, re.IGNORECASE,
    ) and explicit_slide is not None:
        _override(
            "view_slide",
            operation="view",
            slide_id=explicit_slide,
            content=text,
            confidence=0.95,
        )
 
    # CONTEXT CONTINUATION — "more", "another", "add one more", "continue"
    elif re.fullmatch(
        r"(?:add\s+)?(?:one\s+)?more|another|continue|carry\s+on|"
        r"add\s+one\s+more|keep\s+going|yes\s+more|one\s+more\s+point",
        lower,
    ) and last_intent in {
        "add_points", "edit_slide", "transform_content", "add_slide",
    }:
        target = explicit_slide if explicit_slide is not None else current_slide_id
        _override(
            last_intent,
            operation="append" if last_intent == "add_points" else "modify",
            slide_id=target,
            content=last_action or text,
            confidence=0.88,
        )
 
    # VAGUE FOLLOW-UP: "this", "it", "that" → continue last action
    elif re.fullmatch(r"(?:this|it|that|here|go ahead|do it)", lower) \
            and last_intent and current_slide_id:
        _override(
            last_intent,
            operation="append" if last_intent == "add_points" else "modify",
            slide_id=current_slide_id,
            content=last_action or text,
            confidence=0.82,
        )
 
    # ── Step 4: Scope resolution — fill in missing context ───────────────────
    _SLIDE_INTENTS = {
        "transform_content", "add_points", "edit_slide", "view_slide",
        "delete_slide", "move_slide", "swap_slides",
    }
    if (
        result["slide_id"] is None
        and current_slide_id
        and result["intent"] in _SLIDE_INTENTS
    ):
        result["slide_id"]  = current_slide_id
        # Slightly lower confidence: we inferred the target slide
        result["confidence"] = max(0.60, result["confidence"] - 0.10)
 
    if result["ppt_id"] is None and current_ppt_id:
        result["ppt_id"] = current_ppt_id
 
    # ── Step 5: Confidence floor for fully-resolved slide intents ─────────────
    if result["intent"] in _SLIDE_INTENTS and result["slide_id"] is not None:
        result["confidence"] = max(result["confidence"], 0.72)
 
    return result
 
 
# ═══════════════════════════════════════════════════════════════════════════
#  2. CANONICALIZE REASONED PROMPT
# ═══════════════════════════════════════════════════════════════════════════
 
def _canonicalize_reasoned_prompt(user_input: str, reasoned: dict) -> str:
    """
    Build a clean, structured internal instruction from the reasoned intent.
    The result is passed to the downstream slot-filler / executor.
    """
    text     = (user_input or "").strip()
    intent   = _safe_str((reasoned or {}).get("intent", ""))
    slide_id = (reasoned or {}).get("slide_id")
    ppt_id   = _safe_str((reasoned or {}).get("ppt_id", ""))
    content  = _safe_str((reasoned or {}).get("content", ""))
 
    if intent == "transform_content" and slide_id is not None:
        instruction = content or text
        return f"improve slide {slide_id}: {instruction}"
 
    if intent == "view_slide" and slide_id is not None:
        return (f"show slide {slide_id} of {ppt_id}" if ppt_id
                else f"show slide {slide_id}")
 
    if intent == "add_points" and slide_id is not None:
        m = re.search(r"\b(\d+)\s*(?:more\s+)?(?:point|bullet)", text, re.IGNORECASE)
        n = m.group(1) if m else "1"
        return f"add {n} point(s) in slide {slide_id}"
 
    if intent == "edit_slide" and slide_id is not None and content:
        return f"edit slide {slide_id}: {content}"
 
    if intent == "delete_slide" and slide_id is not None:
        return f"delete slide {slide_id}"
 
    return text
 
 
# ═══════════════════════════════════════════════════════════════════════════
#  3. POLISH NEXT QUESTION  (replaces existing in app.py)
# ═══════════════════════════════════════════════════════════════════════════
 
def _polish_next_question(
    user_input: str,
    intent: Optional[str],
    slots: dict,
    missing: list,
    next_question: Optional[str],
) -> str:
    """
    Return a natural, context-specific follow-up question.
    Never uses robotic phrasing like "What change would you like to make?".
    """
    slide_num = slots.get("slide_number")
    ppt_id    = st.session_state.get("current_ppt_id") or slots.get("ppt_ref")
    ppt_label = _ppt_display_label(str(ppt_id)) if ppt_id else "the presentation"
 
    def _slide_title_or_ref(num):
        try:
            outline = _current_outline_payload()
            slides  = (outline or {}).get("slides", [])
            if num and 1 <= int(num) <= len(slides):
                t = slides[int(num) - 1].get("title", "")
                if t:
                    return f'**{t}** (slide {num})'
        except Exception:
            pass
        return f"slide {num}" if num else "that slide"
 
    missing = missing or []
 
    if intent == "edit_slide":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide would you like to edit in {ppt_label}?"
        if "change_content" in missing:
            return (f"What should I change in {_slide_title_or_ref(slide_num)}? "
                    f"(e.g. rewrite a bullet, update the title, add a point)")
 
    if intent == "transform_content":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I improve in {ppt_label}?"
        return (f"How should I improve {_slide_title_or_ref(slide_num)}? "
                f"(e.g. make it shorter, rewrite, make it more professional)")
 
    if intent == "add_points":
        if "slide_number" in missing or slide_num is None:
            return f"Which slide should I add points to in {ppt_label}?"
 
    if intent == "add_slide":
        if "slide_content" in missing:
            return f"What topic should the new slide cover?"
        if "position" in missing:
            return ("Where should I place the new slide? "
                    "(e.g. at the end, after slide 3, at the start)")
 
    if intent == "move_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide would you like to move in {ppt_label}?"
        num = slots.get("slide_number")
        if "anchor_slide" in missing and "position" not in missing:
            return (f"Where should {_slide_title_or_ref(num)} go? "
                    f"(e.g. before slide 4, after slide 6, to position 2)")
        if "position" in missing:
            return f"What position should {_slide_title_or_ref(num)} move to?"
 
    if intent == "swap_slides":
        if "slide_a" in missing or slots.get("slide_a") is None:
            return f"Which two slides should I swap in {ppt_label}?"
        if "slide_b" in missing or slots.get("slide_b") is None:
            return f"Swap slide {slots.get('slide_a')} with which other slide?"
 
    if intent == "delete_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide should I delete from {ppt_label}?"
 
    if intent == "view_slide":
        if "slide_number" in missing or slots.get("slide_number") is None:
            return f"Which slide would you like to see in {ppt_label}?"
 
    # Clean up generic wording if LLM returned it
    q = (next_question or "").strip()
    if q:
        q = re.sub(r"(?i)\bwhat point\b", "what bullet point", q)
        q = re.sub(r"(?i)\bwhat points\b", "what bullet points", q)
        q = re.sub(r"(?i)what change would you like to make\??", "", q).strip()
        if q:
            return q
 
    return "Could you give me a bit more detail?"
 
 
# ═══════════════════════════════════════════════════════════════════════════
#  4. CONFIDENCE-BASED EXECUTION GATE  (NEW)
# ═══════════════════════════════════════════════════════════════════════════
 
def _should_execute_directly(reasoned: dict) -> bool:
    """
    Return True when confidence is high enough to act without asking.
 
    confidence >= 0.60 → execute (may add a brief scope note)
    confidence  < 0.60 → ask for clarification first
    """
    conf = float((reasoned or {}).get("confidence", 0.0) or 0.0)
    return conf >= 0.60
 
 
def _confidence_preface(reasoned: dict) -> Optional[str]:
    """
    For confidence 0.60–0.79, return a brief note so the user knows the
    system inferred the target from context.  Returns None when not needed.
    """
    conf = float((reasoned or {}).get("confidence", 1.0) or 1.0)
    if conf >= 0.80:
        return None
    intent   = (reasoned or {}).get("intent", "")
    slide_id = (reasoned or {}).get("slide_id")
    if intent in {"transform_content", "add_points", "edit_slide"} and slide_id:
        return f"*(Working on slide {slide_id} — let me know if that's not right.)*"
    return None
 
 
# ═══════════════════════════════════════════════════════════════════════════
#  5. NATURAL SUCCESS MESSAGES  (NEW)
# ═══════════════════════════════════════════════════════════════════════════
 
def _natural_success_message(intent: str, slots: dict) -> str:
    """
    Return a conversational confirmation instead of "✅ Slide X updated."
    Pass this as the success_msg argument to commit_changes().
    """
    slide_num = slots.get("slide_number") or slots.get("slide_id")
    n         = slots.get("n_points", 1)
    ppt_id    = st.session_state.get("current_ppt_id", "")
    ppt_label = _ppt_display_label(ppt_id) if ppt_id else "the deck"
 
    slide_title = ""
    if slide_num:
        try:
            outline = _current_outline_payload()
            slides  = (outline or {}).get("slides", [])
            if 1 <= int(slide_num) <= len(slides):
                slide_title = slides[int(slide_num) - 1].get("title", "")
        except Exception:
            pass
 
    title_ref = f'**{slide_title}**' if slide_title else f"slide {slide_num}"
 
    messages = {
        "transform_content": f"Done! I've improved {title_ref}. Here's the updated deck 👇",
        "add_points":        f"Added {n} new point(s) to {title_ref} ✅",
        "edit_slide":        f"Got it — {title_ref} has been updated 👇",
        "add_slide":         f"New slide added to {ppt_label} ✅",
        "delete_slide":      f"Slide {slide_num} removed from {ppt_label}.",
        "move_slide":        f"Slide repositioned in {ppt_label} ✅",
        "swap_slides":       f"Slides swapped in {ppt_label} ✅",
        "refine_ppt":        f"Refreshed! Here's the updated {ppt_label} 👇",
        "create_ppt":        "Here's your new presentation 👇",
    }
    return messages.get(intent, "Done! Here's the updated presentation 👇")
