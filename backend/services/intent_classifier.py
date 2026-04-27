"""
intent_classifier.py  (v2 — LLM-first, regex-fallback)
=======================================================

LLM-backed semantic intent classifier for the AI PPT assistant.

Priority order (from highest to lowest):
  transform_content > add_points > edit_slide > add_slide >
  delete_slide > move_slide > swap_slides > view_slide >
  preview_ppt > switch_ppt > download_ppt > create_ppt >
  ppt_info > suggest_topic > refine_ppt > greeting > smalltalk > general_request
"""

import re
import json
from typing import Any, Optional


# ─────────────────────────────────────────────────────────────────────────────
#  Regex signals  (cheap pre-filter used when LLM is unavailable)
# ─────────────────────────────────────────────────────────────────────────────

_ADD_POINTS_RE = re.compile(
    r"\b(add|append|insert)\b.*\b(point|points|bullet|bullets)\b"
    r"|\bmore points?\b|\bone more\b",
    re.IGNORECASE,
)
_TRANSFORM_RE = re.compile(
    r"\b(improve|enhance|refine|shorten|simplify|rewrite|reword|"
    r"fix|polish|condense|make it better|make it shorter|"
    r"make it professional|make it cleaner|make it clearer|"
    r"make it concise|spice up|jazz up|tighten|upgrade|rephrase|"
    r"make more professional|make more concise|make more clear|"
    r"make it flow|clean up|clean it up|fix this|fix it|"
    r"improve this|improve it|make this better|make it sound better)\b",
    re.IGNORECASE,
)
_ADD_SLIDE_RE = re.compile(
    r"\b(add|create|make|build|insert)\b.*\b(new\s+)?slide\b"
    r"|\bthank you slide\b|\bnew slide\b|\banother slide\b",
    re.IGNORECASE,
)
_EDIT_SLIDE_RE = re.compile(
    r"\b(edit|change|update|modify|revise)\b.*\bslide\b"
    r"|\bchange title\b|\bchange content\b|\bchange subtitle\b",
    re.IGNORECASE,
)
_VIEW_SLIDE_RE = re.compile(
    r"\b(show|view|preview|display)\b.*\bslide\b"
    r"|\bslide\b.*\b(show|view|preview|display)\b",
    re.IGNORECASE,
)
_MOVE_RE = re.compile(
    r"\b(move|swap|reorder|rearrange|re-arrange|re-order|"
    r"switch positions|change position)\b.*\bslide\b",
    re.IGNORECASE,
)
_DELETE_RE = re.compile(r"\b(delete|remove)\b.*\bslide\b", re.IGNORECASE)
_SWITCH_PPT_RE = re.compile(
    r"\b(open|switch to|go to|show|view)\b.*\b(ppt|presentation|deck)\b"
    r"|\bppt\s*\d+\b",
    re.IGNORECASE,
)
_CREATE_RE = re.compile(
    r"\b(make|create|generate|build)\b.*\b(ppt|presentation|deck)\b",
    re.IGNORECASE,
)
_REFINE_RE = re.compile(
    r"\b(more professional|more formal|more casual|more friendly|"
    r"shorter slides|less text|key points only|important points only|"
    r"remove fluff|cleaner deck|cleaner presentation)\b",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _extract_slide_id(text: str) -> Optional[int]:
    match = re.search(
        r"\bslide\s+#?(\d+)\b|\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
        text or "",
        re.IGNORECASE,
    )
    if not match:
        return None
    num = next((g for g in match.groups() if g), None)
    try:
        return int(num) if num is not None else None
    except ValueError:
        return None


def _extract_ppt_id(text: str) -> Optional[str]:
    match = re.search(
        r"\b(?:ppt|presentation|deck)[_\s-]*#?(\d+)\b"
        r"|\b(\d+)(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck)\b",
        text or "",
        re.IGNORECASE,
    )
    if not match:
        return None
    num = next((g for g in match.groups() if g), None)
    return f"ppt {num}" if num else None


def _extract_content(text: str) -> Optional[str]:
    if not text:
        return None
    cleaned = re.sub(r"(?i)\b(?:sorry|no|actually|instead)\b", "", text)
    cleaned = re.sub(r"(?i)\b(?:slide|ppt|presentation|deck)\s+#?\d+\b", "", cleaned)
    cleaned = re.sub(
        r"(?i)^\s*(?:add|insert|create|make|build|edit|change|update|modify|"
        r"revise|improve|enhance|refine|shorten|simplify|rewrite|fix)\s+",
        "",
        cleaned,
    )
    cleaned = re.sub(r"(?i)\b(?:points?|bullets?)\b", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.;:-")
    return cleaned or None


# ─────────────────────────────────────────────────────────────────────────────
#  LLM-backed classification
# ─────────────────────────────────────────────────────────────────────────────

def _llm_classify(user_input: str, ppt_id: Any, slide_id: Any) -> Optional[dict]:
    """
    Use the Azure OpenAI client to semantically classify the user's intent.
    Returns None on any error → caller falls back to regex.
    """
    try:
        from openai import AzureOpenAI
        from backend.config import (
            AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT,
        )

        client = AzureOpenAI(
            api_key=AZURE_KEY,
            api_version=AZURE_API_VERSION,
            azure_endpoint=AZURE_ENDPOINT,
        )

        ctx_parts = []
        if ppt_id:
            ctx_parts.append(f"active_ppt={ppt_id}")
        if slide_id is not None:
            ctx_parts.append(f"active_slide={slide_id}")
        context_line = ", ".join(ctx_parts) if ctx_parts else "none"

        prompt = f"""You are an intent classifier for an AI PowerPoint assistant.

Active context: {context_line}
User said: "{user_input}"

Classify the intent. Return ONLY valid JSON:
{{
  "intent": "<intent>",
  "slide_id": <integer or null>,
  "ppt_id": <"ppt N" string or null>,
  "confidence": <float 0.0-1.0>,
  "operation": "<brief string>",
  "content": "<extracted instruction or null>"
}}

Valid intents (PRIORITY ORDER — pick highest applicable):
1. transform_content  - improve, fix, shorten, rewrite, polish, make better/professional/concise/cleaner, condense, spice up, rephrase, clean up, tighten
2. add_points         – add/append/insert bullet points or more points to a slide
3. edit_slide         – structural edit to a slide (title, layout, content replacement/removal)
4. add_slide          – insert a new slide
5. delete_slide       – remove/delete a slide
6. move_slide         – reposition or reorder a slide
7. swap_slides        – swap two slides
8. view_slide         – show/display/read a specific slide
9. preview_ppt        – show full deck preview
10. switch_ppt        – open or switch to a different deck
11. download_ppt      – download the PPTX file
12. create_ppt        – create a new presentation
13. ppt_info          – how many ppts, list ppts, get topic of a ppt
14. suggest_topic     – brainstorm new presentation topic ideas
15. refine_ppt        – deck-level tone/style refresh
16. greeting          – hi, hello, hey
17. smalltalk         – ok, thanks, cool, got it, yes, no
18. general_request   – anything else

Key rules:
- "improve", "fix", "shorten", "make it better", "make it professional",
  "rewrite", "polish", "condense", "simplify", "make it cleaner" → ALWAYS transform_content
- "add X points/bullets" → add_points (NOT edit_slide)
- slide_id: extract from message if present, else null (caller supplies context)
- ppt_id: extract "ppt N" style reference if present, else null
- confidence: 0.95+ for clear requests, 0.7-0.9 for inferred, 0.5-0.7 for ambiguous

Return ONLY valid JSON. No markdown. No explanation."""

        resp = client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=220,
        )
        raw = (resp.choices[0].message.content or "").strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
        data = json.loads(raw)
        if not isinstance(data, dict) or "intent" not in data:
            return None
        return data
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────────────

def classify_ppt_intent(
    user_input: str,
    ppt_id: Any = None,
    slide_id: Any = None,
) -> dict:
    """
    Classify the user's intent with LLM-first, regex-fallback strategy.

    Returns dict:
        intent, ppt_id, slide_id, content, operation, target, confidence
    """
    text = _norm(user_input)

    result = {
        "intent":     "general_request",
        "ppt_id":     ppt_id if ppt_id is not None else _extract_ppt_id(text),
        "slide_id":   slide_id if slide_id is not None else _extract_slide_id(text),
        "content":    None,
        "operation":  None,
        "target":     None,
        "confidence": 0.35,
    }

    if not text:
        return result

    _TARGET_MAP = {
        "transform_content": "slide",
        "add_points":        "slide",
        "edit_slide":        "slide",
        "add_slide":         "slide",
        "delete_slide":      "slide",
        "move_slide":        "slide",
        "swap_slides":       "slide",
        "view_slide":        "slide",
        "preview_ppt":       "ppt",
        "switch_ppt":        "ppt",
        "download_ppt":      "ppt",
        "create_ppt":        "ppt",
        "ppt_info":          "ppt",
        "suggest_topic":     "meta",
        "refine_ppt":        "ppt",
    }

    # ── LLM first ─────────────────────────────────────────────────────────
    llm = _llm_classify(text, ppt_id, slide_id)
    if llm and isinstance(llm, dict) and llm.get("intent"):
        intent = llm["intent"]
        conf   = float(llm.get("confidence", 0.7) or 0.7)

        result["intent"]     = intent
        result["confidence"] = conf
        result["content"]    = llm.get("content") or _extract_content(text)
        result["operation"]  = llm.get("operation") or intent
        result["target"]     = _TARGET_MAP.get(intent)

        # slide_id: explicit in message wins over LLM extraction over context
        explicit_slide = _extract_slide_id(text)
        llm_slide      = llm.get("slide_id")
        if explicit_slide is not None:
            result["slide_id"] = explicit_slide
        elif llm_slide is not None:
            result["slide_id"] = llm_slide
        # else keep passed-in context slide_id

        # ppt_id: explicit in message wins
        explicit_ppt = _extract_ppt_id(text)
        llm_ppt      = llm.get("ppt_id")
        if explicit_ppt:
            result["ppt_id"] = explicit_ppt
        elif llm_ppt:
            result["ppt_id"] = llm_ppt

        return result

    # ── Regex fallback ────────────────────────────────────────────────────
    if _SWITCH_PPT_RE.search(text):
        result.update(intent="switch_ppt",   target="ppt",   operation="switch",    confidence=0.96)
        return result

    if _CREATE_RE.search(text):
        result.update(intent="create_ppt",   target="ppt",   operation="create",    confidence=0.95)
        return result

    if _DELETE_RE.search(text):
        result.update(intent="delete_slide", target="slide", operation="delete",    confidence=0.95)
        return result

    if _MOVE_RE.search(text):
        result.update(intent="move_slide",   target="slide", operation="move",      confidence=0.94)
        return result

    if _VIEW_SLIDE_RE.search(text):
        result.update(intent="view_slide",   target="slide", operation="view",      confidence=0.93)
        return result

    # transform_content beats add_points
    if _TRANSFORM_RE.search(text):
        result.update(intent="transform_content", target="slide", operation="transform", confidence=0.97)
        if result["slide_id"] is None:
            result["slide_id"] = _extract_slide_id(text)
        return result

    if _REFINE_RE.search(text):
        result.update(intent="refine_ppt",   target="ppt",   operation="refine",    confidence=0.88)
        return result

    if _ADD_POINTS_RE.search(text):
        result.update(intent="add_points",   target="slide", operation="append",    confidence=0.96)
        result["content"] = _extract_content(text)
        return result

    if _ADD_SLIDE_RE.search(text):
        result.update(intent="add_slide",    target="slide", operation="insert",    confidence=0.95)
        result["content"] = _extract_content(text)
        return result

    if _EDIT_SLIDE_RE.search(text) or result["slide_id"] is not None:
        conf = 0.88 if result["slide_id"] is not None else 0.82
        result.update(intent="edit_slide", target="slide", operation="modify", confidence=conf)
        result["content"] = _extract_content(text)
        return result

    return result