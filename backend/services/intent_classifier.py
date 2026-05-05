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
    r"|\bmore points?\b|\bone more\b"
    r"|\b(?:make|expand|elaborate)\b.*\bslide\s*#?\d+\b.*\b(?:longer|more detailed|more detail|detailed|expanded)\b"
    r"|\bslide\s*#?\d+\b.*\b(?:longer|more detailed|more detail|detailed|expanded)\b",
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
    r"\b(add|insert)\b.*\b(new\s+)?slide\b"
    r"|\b(?:create|make|build|generate)\b.*\bnew\s+slide\b"
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
_EXPLAIN_SLIDE_RE = re.compile(
    r"\b(?:explain|describe|summarize|summary of|walk me through)\b.*\bslide\s*#?\d+\b"
    r"|\bslide\s*#?\d+\b.*\b(?:explain|describe|summarize|about)\b",
    re.IGNORECASE,
)
_MERGE_SLIDES_RE = re.compile(
    r"\b(?:merge|combine|join|consolidate)\b.*\bslides?\b",
    re.IGNORECASE,
)
_UPDATE_SLIDE_RE = re.compile(
    r"\b(?:include|add|insert|append|update|change|modify|improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|condense|expand|elaborate)\b",
    re.IGNORECASE,
)
_CLEAR_SLIDE_RE = re.compile(
    r"\b(remove everything|clear slide|clear all|erase all|wipe slide|wipe all|delete all content|remove all content|blank slide|make blank slide|empty slide)\b",
    re.IGNORECASE,
)
_CANCEL_RE = re.compile(r"^(stop|cancel|nevermind|never mind|abort|quit)$", re.IGNORECASE)
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


def _normalize_intent(intent: Any, text: str = "") -> str:
    raw = str(intent or "").strip().lower()
    aliases = {
        "show_slide": "view_slide",
        "display_slide": "view_slide",
        "open_slide": "view_slide",
        "remove_slide": "delete_slide",
        "merge_slide": "merge_slides",
        "merge": "merge_slides",
        "explain": "explain_slide",
        "explain_slide_content": "explain_slide",
        "update_slide_content": "update_slide",
        "add_points": "update_slide",
        "transform_content": "update_slide",
        "general_request": "unknown",
    }
    return aliases.get(raw, raw)


def _classify_update_sub_intent(text: str) -> Optional[str]:
    lower = _norm(text).lower()
    if re.search(r"\b(?:add|include|insert|append)\b.*\b(?:point|points|bullet|bullets|stat|stats|statistics|data|metric|metrics)\b", lower):
        return "add_point"
    if re.search(r"\b(?:longer|expand|expanded|elaborate|more detail|more detailed|add detail|add details)\b", lower):
        return "expand"
    if re.search(r"\b(?:shorten|shorter|condense|concise|trim|less text|summarize)\b", lower):
        return "shorten"
    if re.search(r"\b(?:rewrite|reword|rephrase|polish|improve|make it better|make better|fix)\b", lower):
        return "rewrite"
    return None


def _extract_slide_ids(text: str) -> list[int]:
    values = []
    for match in re.finditer(r"\bslide\s*#?(\d+)\b|\b(\d+)(?:st|nd|rd|th)?\s+slide\b", text or "", re.IGNORECASE):
        num = next((g for g in match.groups() if g), None)
        if num is None:
            continue
        try:
            parsed = int(num)
        except ValueError:
            continue
        if parsed not in values:
            values.append(parsed)
    if re.search(r"\b(?:merge|combine|join|consolidate)\b", text or "", re.IGNORECASE):
        for raw in re.findall(r"\b(\d+)\b", text or ""):
            try:
                parsed = int(raw)
            except ValueError:
                continue
            if parsed not in values:
                values.append(parsed)
    return values


def _extract_add_slide_placement(text: str) -> tuple[Optional[str], Optional[int]]:
    if not text:
        return None, None
    match = re.search(
        r"\b(after|before)\s+slide\s*#?(\d+)\b"
        r"|\bat\s+the\s+(start|beginning|end)\b"
        r"|\b(?:to|at)\s+position\s+(\d+)\b",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None, None
    if match.group(1) and match.group(2):
        return match.group(1).lower(), int(match.group(2))
    if match.group(3):
        pos = match.group(3).lower()
        return ("start" if pos in {"start", "beginning"} else "end"), None
    if match.group(4):
        return "position", int(match.group(4))
    return None, None


def _extract_content(text: str) -> Optional[str]:
    if not text:
        return None
    cleaned = re.sub(r"(?i)\b(?:sorry|no|actually|instead)\b", "", text)
    # Remove phrases like 'in ppt 2', 'of presentation', 'in the deck 3'
    cleaned = re.sub(r"(?i)\b(?:in|of|for|on|about)\s+(?:the\s+)?(?:ppt|presentation|deck)(?:\s*#?\d+)?\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:after|before)\s+slide\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\bat\s+the\s+(?:start|beginning|end)\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:to|at)\s+position\s+\d+\b", "", cleaned)
    # Remove concrete slide references like 'slide 3' or '3rd slide'.
    cleaned = re.sub(r"(?i)\bslide\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b\d+(?:st|nd|rd|th)?\s+slide\b", "", cleaned)
    # Remove standalone slide/ppt words (they can appear without numbers)
    cleaned = re.sub(r"(?i)\b(?:slide|ppt|presentation|deck)\b", "", cleaned)
    # Remove leading action verbs
    cleaned = re.sub(
        r"(?i)^\s*(?:add|insert|create|make|build|edit|change|update|modify|"
        r"revise|improve|enhance|refine|shorten|simplify|rewrite|fix)\s+",
        "",
        cleaned,
    )
    # Remove generic point/bullet mentions
    cleaned = re.sub(r"(?i)\b(?:points?|bullets?)\b", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.;:-")
    if _norm(cleaned).lower() in {"a new", "an new", "the new", "new", "another", "one more"}:
        return None
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
  "action_type": "<slide|deck|meta|unknown>",
  "target": "<existing_slide|new_slide|deck|conversation|unknown>",
  "slide_id": <integer or null>,
  "slide_ids": [<integer>, ...],
  "ppt_id": <"ppt N" string or null>,
  "sub_intent": "<add_point|expand|shorten|rewrite|null>",
  "confidence": <float 0.0-1.0>,
  "operation": "<brief string>",
  "content": "<extracted instruction or null>",
  "position": "<after|before|start|end|position or null>",
  "anchor_slide": <integer or null>,
  "missing_entities": ["slide_number", ...],
  "clarification_question": "<question or null>"
}}

Valid intents (PRIORITY ORDER — pick highest applicable):
1. view_slide         - show/display/read one specific existing slide
2. explain_slide      - explain/summarize/describe one existing slide
3. delete_slide       - remove/delete one existing slide
4. merge_slides       - merge/combine two or more existing slides
5. update_slide       - change content on an existing slide
6. edit_slide         – structural edit to an existing slide (title, layout, content replacement/removal)
4. clear_slide        – remove everything from a slide after confirmation
5. blank_slide        – clear a slide immediately without extra confirmation
6. add_slide          – insert a new slide
7. delete_slide       – remove/delete a slide
8. move_slide         – reposition or reorder a slide
9. swap_slides        – swap two slides
10. view_slide        – show/display/read a specific slide
11. preview_ppt       – show full deck preview
12. switch_ppt       – open or switch to a different deck
13. download_ppt      – download the PPTX file
14. create_ppt        – create a new presentation
15. ppt_info          – how many ppts, list ppts, get topic of a ppt
16. suggest_topic     – brainstorm new presentation topic ideas
17. refine_ppt        – deck-level tone/style refresh
18. cancel            – stop/cancel/abort the current pending request
19. greeting          – hi, hello, hey
20. smalltalk         – ok, thanks, cool, got it, yes, no, nothing to add
21. unknown           – anything unclear or unsupported

Key rules:
- NEVER classify unclear text as add_slide.
- Use add_slide ONLY when the user clearly asks for a new/another slide.
- "show slide 4 in ppt 3" → view_slide with slide_id=4 and ppt_id="ppt 3"; never preview_ppt
- "explain slide 4" → explain_slide
- "remove slide 3" → delete_slide
- "merge slide 4 and 5" → merge_slides with slide_ids=[4,5]
- "make slide 5 longer", "expand slide 5", "add more detail to slide 5" → update_slide, sub_intent=expand, not add_slide
- "include statistics" → update_slide only if active_slide is available; otherwise unknown with missing_entities=["slide_number"]
- "improve", "fix", "shorten", "make it better", "make it professional",
  "rewrite", "polish", "condense", "simplify", "make it cleaner" on an existing slide → update_slide
- For update_slide, set sub_intent to one of: add_point, expand, shorten, rewrite.
- "add X points/bullets" to an existing slide → update_slide (NOT add_slide)
- "add slide", "add a new slide", "add a summary slide" → add_slide
- If a message contains an explicit existing slide number like "slide 5", do not classify it as add_slide unless the user also clearly asks for a new slide.
- If confidence < 0.60, use intent="unknown" and include a clarification_question.
- "remove everything", "clear slide", "blank slide" → clear_slide or blank_slide
- "stop", "cancel" → cancel
- "add nothing", "nothing to add", "no change" → general_request or smalltalk
- slide_id: extract from message if present, else null (caller supplies context)
- ppt_id: extract "ppt N" style reference if present, else null
- confidence: 0.95+ for clear requests, 0.7-0.9 for inferred, 0.5-0.7 for ambiguous

Return ONLY valid JSON. No markdown. No explanation."""

        kwargs = {
            "model": AZURE_DEPLOYMENT,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.0,
            "max_tokens": 360,
        }
        try:
            resp = client.chat.completions.create(
                **kwargs,
                response_format={"type": "json_object"},
            )
        except Exception:
            resp = client.chat.completions.create(**kwargs)
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
    explicit_ppt = _extract_ppt_id(text)
    explicit_slide = _extract_slide_id(text)

    result = {
        "intent":     "unknown",
        "ppt_id":     explicit_ppt if explicit_ppt is not None else ppt_id,
        "slide_id":   explicit_slide,
        "slide_number": explicit_slide,
        "slide_ids":  _extract_slide_ids(text),
        "content":    None,
        "operation":  None,
        "target":     None,
        "action_type": "unknown",
        "sub_intent": None,
        "position":   None,
        "anchor_slide": None,
        "missing_entities": [],
        "clarification_question": None,
        "confidence": 0.35,
    }

    if not text:
        return result

    _TARGET_MAP = {
        "transform_content": "slide",
        "add_points":        "slide",
        "update_slide":      "slide",
        "explain_slide":     "slide",
        "edit_slide":        "slide",
        "add_slide":         "ppt",
        "delete_slide":      "slide",
        "merge_slides":      "slide",
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
        "unknown":           "unknown",
    }
    _ACTION_TYPE_MAP = {
        "view_slide": "slide",
        "explain_slide": "slide",
        "delete_slide": "slide",
        "merge_slides": "slide",
        "update_slide": "slide",
        "edit_slide": "slide",
        "add_slide": "deck",
        "create_ppt": "deck",
        "preview_ppt": "deck",
        "switch_ppt": "deck",
        "download_ppt": "deck",
        "ppt_info": "deck",
        "refine_ppt": "deck",
        "suggest_topic": "meta",
        "cancel": "meta",
        "greeting": "meta",
        "smalltalk": "meta",
        "unknown": "unknown",
    }

    # ── LLM first ─────────────────────────────────────────────────────────
    llm = _llm_classify(text, ppt_id, slide_id)
    if llm and isinstance(llm, dict) and llm.get("intent"):
        intent = _normalize_intent(llm["intent"], text)
        conf   = float(llm.get("confidence", 0.7) or 0.7)
        if conf < 0.60:
            intent = "unknown"

        result["intent"]     = intent
        result["confidence"] = conf
        result["content"]    = llm.get("content") or _extract_content(text)
        result["operation"]  = llm.get("operation") or intent
        result["target"]     = _TARGET_MAP.get(intent)
        result["action_type"] = llm.get("action_type") or _ACTION_TYPE_MAP.get(intent, "unknown")
        result["sub_intent"] = llm.get("sub_intent") or (_classify_update_sub_intent(text) if intent == "update_slide" else None)
        result["position"]   = llm.get("position")
        result["anchor_slide"] = llm.get("anchor_slide")
        result["missing_entities"] = llm.get("missing_entities") if isinstance(llm.get("missing_entities"), list) else []
        result["clarification_question"] = llm.get("clarification_question")

        # slide_id: explicit in message wins over LLM extraction.
        # Do not copy the active context slide here; the reasoning layer decides
        # when contextual scope is safe for the specific intent.
        explicit_slide = _extract_slide_id(text)
        llm_slide      = llm.get("slide_id")
        if explicit_slide is not None:
            result["slide_id"] = explicit_slide
        elif llm_slide is not None:
            result["slide_id"] = llm_slide
        result["slide_number"] = result["slide_id"]
        result["slide_ids"] = _extract_slide_ids(text) or ([result["slide_id"]] if result["slide_id"] else [])

        # ppt_id: explicit in message wins
        explicit_ppt = _extract_ppt_id(text)
        llm_ppt      = llm.get("ppt_id")
        if explicit_ppt:
            result["ppt_id"] = explicit_ppt
        elif llm_ppt:
            result["ppt_id"] = llm_ppt
        
        # CONFIDENCE BOOST for view_slide with explicit slide number
        if intent == "view_slide" and explicit_slide is not None:
            result["confidence"] = 0.99

        if intent == "add_slide":
            position, anchor_slide = _extract_add_slide_placement(text)
            if position:
                result["position"] = position
            if anchor_slide is not None:
                result["anchor_slide"] = anchor_slide
            result["slide_id"] = None
            result["slide_number"] = None
            missing = []
            if not result.get("content"):
                missing.append("slide_content")
            if not result.get("position"):
                missing.append("position")
            result["missing_entities"] = missing

        if intent in {"update_slide", "explain_slide", "delete_slide"} and result["slide_id"] is None:
            result["missing_entities"] = sorted(set(result["missing_entities"] + ["slide_number"]))
            result["confidence"] = min(result["confidence"], 0.59)
            result["intent"] = "unknown"
            result["clarification_question"] = result["clarification_question"] or "Which slide should I work on?"

        return result

    # ── Regex fallback ────────────────────────────────────────────────────
    # Conservative fallback: only high-precision patterns execute. Anything
    # unclear stays unknown so the app asks a clarification instead of adding a slide.
    if _EXPLAIN_SLIDE_RE.search(text):
        result.update(
            intent="explain_slide", target="slide", action_type="slide",
            operation="explain", confidence=0.96,
        )
        return result

    if _VIEW_SLIDE_RE.search(text):
        result.update(intent="view_slide", target="slide", action_type="slide", operation="view", confidence=0.98)
        if _extract_slide_id(text):
            result["slide_id"] = _extract_slide_id(text)
            result["slide_number"] = result["slide_id"]
            result["confidence"] = 0.99  # Explicit slide number = max confidence
        return result

    if _MERGE_SLIDES_RE.search(text):
        slide_ids = _extract_slide_ids(text)
        result.update(
            intent="merge_slides", target="slide", action_type="slide",
            operation="merge", slide_ids=slide_ids, confidence=0.94 if len(slide_ids) >= 2 else 0.58,
        )
        if len(slide_ids) < 2:
            result["intent"] = "unknown"
            result["missing_entities"] = ["slide_numbers"]
            result["clarification_question"] = "Which slides should I merge?"
        return result

    if _SWITCH_PPT_RE.search(text):
        result.update(intent="switch_ppt", target="ppt", action_type="deck", operation="switch", confidence=0.96)
        return result

    if _CREATE_RE.search(text):
        result.update(intent="create_ppt", target="ppt", action_type="deck", operation="create", confidence=0.95)
        return result

    if _DELETE_RE.search(text):
        result.update(intent="delete_slide", target="slide", action_type="slide", operation="delete", confidence=0.95)
        return result

    if _MOVE_RE.search(text):
        result.update(intent="move_slide", target="slide", action_type="slide", operation="move", confidence=0.94)
        return result

    if _CLEAR_SLIDE_RE.search(text):
        if re.search(r"\bmake blank slide|blank slide|empty slide\b", text, re.IGNORECASE):
            result.update(intent="blank_slide", target="slide", action_type="slide", operation="clear", confidence=0.96)
        else:
            result.update(intent="clear_slide", target="slide", action_type="slide", operation="clear", confidence=0.96)
        result["content"] = _extract_content(text)
        return result

    if _CANCEL_RE.search(text):
        result.update(intent="cancel", target="meta", action_type="meta", operation="cancel", confidence=0.99)
        return result

    if _REFINE_RE.search(text):
        result.update(intent="refine_ppt", target="ppt", action_type="deck", operation="refine", confidence=0.88)
        return result

    if _ADD_POINTS_RE.search(text) or _TRANSFORM_RE.search(text) or (
        _UPDATE_SLIDE_RE.search(text)
        and not _ADD_SLIDE_RE.search(text)
        and (explicit_slide is not None or slide_id is not None)
    ):
        result.update(intent="update_slide", target="slide", action_type="slide", operation="update", confidence=0.92)
        result["content"] = _extract_content(text)
        result["sub_intent"] = _classify_update_sub_intent(text)
        if result["slide_id"] is None and slide_id is not None:
            result["slide_id"] = slide_id
            result["slide_number"] = slide_id
        return result

    if _ADD_SLIDE_RE.search(text) and not re.search(
        r"\b(?:make|edit|change|update|modify|revise)\s+slide\s*#?\d+\b",
        text,
        re.IGNORECASE,
    ):
        result.update(intent="add_slide", target="ppt", action_type="deck", operation="insert", confidence=0.95)
        result["content"] = _extract_content(text)
        result["position"], result["anchor_slide"] = _extract_add_slide_placement(text)
        result["slide_id"] = None
        result["slide_number"] = None
        if not result.get("content"):
            result["missing_entities"].append("slide_content")
        if not result.get("position"):
            result["missing_entities"].append("position")
        return result

    if _EDIT_SLIDE_RE.search(text):
        conf = 0.88 if result["slide_id"] is not None else 0.58
        result.update(intent="edit_slide", target="slide", action_type="slide", operation="modify", confidence=conf)
        result["content"] = _extract_content(text)
        if result["slide_id"] is None:
            result["intent"] = "unknown"
            result["missing_entities"] = ["slide_number"]
            result["clarification_question"] = "Which slide would you like to edit?"
        return result

    if _UPDATE_SLIDE_RE.search(text) and not _ADD_SLIDE_RE.search(text):
        result.update(
            intent="unknown",
            target="existing_slide",
            action_type="slide",
            operation="clarify",
            content=_extract_content(text) or text,
            sub_intent=_classify_update_sub_intent(text),
            missing_entities=["slide_number"],
            clarification_question="Which slide should I update?",
            confidence=0.45,
        )
        if slide_id is not None:
            result.update(
                intent="update_slide",
                slide_id=slide_id,
                slide_number=slide_id,
                confidence=0.72,
                missing_entities=[],
                clarification_question=None,
                operation="update",
            )
        return result

    return result
