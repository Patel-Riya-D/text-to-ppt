"""
intent_classifier.py  (v4 — comprehensive NLU)
===============================================

Handles ALL natural language query categories:
  - Edit / modify existing slides
  - Add / remove slides
  - Content enhancement (examples, stats, context)
  - Design / style changes
  - Context-aware follow-ups
  - Regeneration control
  - Edge cases (1-slide summary, 20 slides, storytelling, Hindi)
  - Presentation mode (speaker notes, scripts, audience questions)
  - Deck-level bulk operations (make all slides shorter, add to each slide)

CRITICAL: NEVER default to add_slide for unclear intent → always ask clarification.
"""

from __future__ import annotations

import json
import re
from typing import Any, Optional

# ─────────────────────────────────────────────────────────────────────────────
#  Utility helpers
# ─────────────────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())

def _lower(text: str) -> str:
    return _norm(text).lower()

def _extract_slide_id(text: str) -> Optional[int]:
    m = re.search(
        r"\bslide[_\s-]*#?(\d+)\b|\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
        text or "", re.IGNORECASE,
    )
    if not m:
        return None
    num = next((g for g in m.groups() if g), None)
    try:
        return int(num) if num is not None else None
    except ValueError:
        return None

def _extract_slide_ids(text: str) -> list[int]:
    values: list[int] = []
    for m in re.finditer(
        r"\bslide[_\s-]*#?(\d+)\b|\b(\d+)(?:st|nd|rd|th)?\s+slide\b",
        text or "", re.IGNORECASE,
    ):
        num = next((g for g in m.groups() if g), None)
        try:
            v = int(num)  # type: ignore[arg-type]
            if v not in values:
                values.append(v)
        except (ValueError, TypeError):
            pass
    if re.search(r"\b(?:merge|combine|join|swap|exchange)\b", text or "", re.IGNORECASE):
        for raw in re.findall(r"\b(\d+)\b", text or ""):
            try:
                v = int(raw)
                if v not in values:
                    values.append(v)
            except ValueError:
                pass
    return values

def _extract_ppt_id(text: str) -> Optional[str]:
    m = re.search(
        r"\b(?:ppt|presentation|deck)[_\s-]*#?(\d+)\b"
        r"|\b(\d+)(?:st|nd|rd|th)?\s*(?:ppt|presentation|deck)\b",
        text or "", re.IGNORECASE,
    )
    if not m:
        return None
    num = next((g for g in m.groups() if g), None)
    return f"ppt {num}" if num else None

def _extract_slide_count(text: str) -> Optional[int]:
    for pat in (
        r"\b(?:into|to|make)\s+(\d+)\s+slide",
        r"\b(\d+)\s+slide",
        r"\bexpand\s+(?:to|into)\s+(\d+)",
    ):
        m = re.search(pat, text or "", re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                pass
    return None

def _normalize_intent(raw: Any) -> str:
    s = str(raw or "").strip().lower()
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
        "transform": "transform_content",
        "general_request": "unknown",
        "bulk_edit": "bulk_update",
        "all_slides": "bulk_update",
        "speaker_notes": "presentation_mode",
        "audience_questions": "presentation_mode",
        "slide_script": "presentation_mode",
    }
    return aliases.get(s, s)

def _classify_update_sub_intent(text: str) -> Optional[str]:
    low = _lower(text)
    if re.search(r"\b(?:add|include|insert|append)\b.*\b(?:point|bullet|stat|data|metric|example|case study|comparison|table|chart|context|trend|note)\b", low):
        return "add_content"
    if re.search(r"\b(?:longer|expand|elaborate|more detail|more detailed|add detail|add more)\b", low):
        return "expand"
    if re.search(r"\b(?:shorten|shorter|condense|concise|trim|less text|brief|compress|summarize)\b", low):
        return "shorten"
    if re.search(r"\b(?:rewrite|reword|rephrase|polish|improve|make it better|fix|engaging|impactful|punchy)\b", low):
        return "rewrite"
    if re.search(r"\b(?:bullet|bullets|bullet points|convert to bullet)\b", low):
        return "bulletize"
    if re.search(r"\b(?:formal|professional|business|academic)\b", low):
        return "formalize"
    if re.search(r"\b(?:casual|simple|beginner|easy|friendly|conversational)\b", low):
        return "simplify"
    if re.search(r"\b(?:storytelling|narrative|story|story format)\b", low):
        return "storytelling"
    return None

def _classify_presentation_mode_sub_intent(text: str) -> Optional[str]:
    low = _lower(text)
    if re.search(r"\b(?:speaker notes?|speaker script|notes)\b", low):
        return "speaker_notes"
    if re.search(r"\b(?:audience question|questions? audience|q\s*&\s*a|qa question)\b", low):
        return "audience_questions"
    if re.search(r"\b(?:script|short script|slide script|script for each)\b", low):
        return "slide_script"
    if re.search(r"\b(?:2 minute|2-minute|two minute|speech|summarize.*speech|summary.*speech)\b", low):
        return "speech_summary"
    if re.search(r"\b(?:explain.*present|present.*explain|like i.m present|presenting|explain\s+(?:this\s+)?(?:ppt|presentation|deck)|walk\s+me\s+through\s+(?:this\s+)?(?:ppt|presentation|deck))\b", low):
        return "presentation_explain"
    if re.search(r"\b(?:interactive question|one question per slide|question per slide)\b", low):
        return "interactive_questions"
    return "speaker_notes"

def _classify_design_sub_intent(text: str) -> Optional[str]:
    low = _lower(text)
    if re.search(r"\b(?:dark theme|dark mode|dark background)\b", low):
        return "dark_theme"
    if re.search(r"\b(?:minimal|minimalist|clean|simple design)\b", low):
        return "minimal_design"
    if re.search(r"\b(?:modern|contemporary|sleek)\b", low):
        return "modern_design"
    if re.search(r"\b(?:business|corporate|professional|investor|formal)\b", low):
        return "business_theme"
    if re.search(r"\b(?:visual|icon|image|graphic)\b", low):
        return "add_visuals"
    if re.search(r"\b(?:heading|headline|impactful heading|bold heading)\b", low):
        return "impactful_headings"
    return "general_design"

def _is_bulk_operation(text: str) -> bool:
    """Returns True if the operation targets ALL slides or the whole deck."""
    low = _lower(text)
    return bool(re.search(
        r"\b(?:all slides?|every slide|each slide|entire (?:ppt|deck|presentation)|"
        r"whole (?:ppt|deck|presentation)|throughout|across all|all content)\b",
        low,
    ))

def _extract_content(text: str) -> Optional[str]:
    cleaned = re.sub(r"(?i)\b(?:sorry|no|actually|instead)\b", "", text or "")
    cleaned = re.sub(r"(?i)\b(?:in|of|for|on|about)\s+(?:the\s+)?(?:ppt|presentation|deck)(?:\s*#?\d+)?\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:after|before)\s+slide\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\bat\s+the\s+(?:start|beginning|end)\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:to|at)\s+position\s+\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\bslide\s*#?\d+\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b\d+(?:st|nd|rd|th)?\s+slide\b", "", cleaned)
    cleaned = re.sub(r"(?i)\b(?:slide|ppt|presentation|deck)\b", "", cleaned)
    cleaned = re.sub(
        r"(?i)^\s*(?:add|insert|create|make|build|edit|change|update|modify|"
        r"revise|improve|enhance|refine|shorten|simplify|rewrite|fix|include|"
        r"convert|translate|regenerate|rewrite|give|provide)\s+", "", cleaned,
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.;:-")
    return cleaned or None


# ─────────────────────────────────────────────────────────────────────────────
#  LLM classification
# ─────────────────────────────────────────────────────────────────────────────

_LLM_SYSTEM = """You are a PowerPoint assistant intent classifier. Analyse what the user wants and return JSON.

VALID INTENTS (pick exactly one):
cancel            - stop/cancel/abort
clear_slide       - remove all content from a slide
blank_slide       - make slide blank immediately
view_slide        - show/display a specific slide
explain_slide     - explain/describe a specific slide
delete_slide      - remove a slide permanently
merge_slides      - merge/combine two slides
swap_slides       - swap positions of two slides
move_slide        - reposition a slide
transform_content - improve/fix/shorten/rewrite a specific slide
add_points        - add bullet points to a specific slide
update_slide      - general content update to a specific slide (stats, examples, context, table, etc.)
edit_slide        - structural edit (title/layout change)
add_slide         - insert a brand NEW slide
bulk_update       - operation affecting ALL slides or the whole deck
regenerate_slide  - regenerate only a specific slide from scratch
design_change     - theme/style/visual change (dark mode, minimal, business)
presentation_mode - speaker notes, scripts, audience questions, speech summary, interactive questions
refine_ppt        - deck-level tone/style/language refresh
preview_ppt       - show full deck preview
switch_ppt        - open or switch to a different deck
download_ppt      - download the PPTX file
create_ppt        - create a brand-new presentation
ppt_info          - info about decks (count/list/topic)
suggest_topic     - brainstorm new topic ideas
translate_ppt     - translate the entire PPT to another language
greeting          - hi/hello/hey
smalltalk         - ok/thanks/yes/no/cool
unknown           - anything unclear (ask clarification, NEVER use add_slide as fallback)

CRITICAL RULES:
1. NEVER use add_slide when intent is unclear → use unknown instead
2. "explain slide N" → explain_slide (never add_slide)
3. "remove/delete slide N" → delete_slide
4. "merge slide N and M" → merge_slides
5. "make all slides shorter" / "add examples to each slide" → bulk_update
6. "add statistics / examples / Indian context / real-world examples" → update_slide (if specific slide) OR bulk_update (if no slide specified or "each/all")
7. "speaker notes / script / audience questions / 2-minute speech" → presentation_mode
8. "dark theme / modern / business presentation / minimal" → design_change
9. "translate to Hindi" / "translate this PPT" → translate_ppt
10. "regenerate slide 2 / rewrite conclusion slide" → regenerate_slide
11. "1-slide summary / expand to 20 slides" → bulk_update with sub_intent=resize
12. "remove all text / keep only headings" → bulk_update with sub_intent=strip_text
13. "storytelling format / convert to story" → bulk_update with sub_intent=storytelling
14. "convert to bullet points" → bulk_update (if all slides) or update_slide (if one slide) with sub_intent=bulletize
15. confidence < 0.60 → intent=unknown, include clarification_question
16. For presentation_mode, always specify sub_intent: speaker_notes | audience_questions | slide_script | speech_summary | presentation_explain | interactive_questions"""

_LLM_SCHEMA = """{
  "intent": "<intent>",
  "action_type": "<slide|deck|meta|unknown>",
  "target": "<existing_slide|new_slide|all_slides|deck|conversation|unknown>",
  "slide_id": <integer or null>,
  "slide_ids": [<integers>],
  "ppt_id": "<'ppt N' or null>",
  "sub_intent": "<string or null>",
  "target_language": "<language name or null>",
  "slide_count_target": <integer or null>,
  "confidence": <float 0.0-1.0>,
  "operation": "<brief name>",
  "content": "<extracted instruction or null>",
  "position": "<after|before|start|end|position or null>",
  "anchor_slide": <integer or null>,
  "missing_entities": ["slide_number", ...],
  "clarification_question": "<question or null>"
}"""

def _build_prompt(user_input: str, ppt_id: Any, slide_id: Any) -> str:
    ctx = []
    if ppt_id:
        ctx.append(f"active_ppt={ppt_id}")
    if slide_id is not None:
        ctx.append(f"active_slide={slide_id}")
    context_line = ", ".join(ctx) or "none"

    return f"""{_LLM_SYSTEM}

Active context: {context_line}
User said: "{user_input}"

Confidence: 0.95+ explicit request, 0.70-0.94 clear inference, 0.50-0.69 ambiguous, <0.50 unclear

Return ONLY this JSON (no markdown):
{_LLM_SCHEMA}"""

def _llm_classify(user_input: str, ppt_id: Any, slide_id: Any) -> Optional[dict]:
    try:
        from openai import AzureOpenAI
        from backend.config import AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
        client = AzureOpenAI(api_key=AZURE_KEY, api_version=AZURE_API_VERSION, azure_endpoint=AZURE_ENDPOINT)
        prompt = _build_prompt(user_input, ppt_id, slide_id)
        kwargs = dict(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=450,
        )
        try:
            resp = client.chat.completions.create(**kwargs, response_format={"type": "json_object"})
        except Exception:
            resp = client.chat.completions.create(**kwargs)
        raw = (resp.choices[0].message.content or "").strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
        data = json.loads(raw)
        return data if isinstance(data, dict) and "intent" in data else None
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
#  Conservative regex fallback (LLM unavailable)
# ─────────────────────────────────────────────────────────────────────────────

_R = re.IGNORECASE

_RE_CANCEL       = re.compile(r"^(?:stop|cancel|nevermind|never\s*mind|abort|quit)$", _R)
_RE_CLEAR        = re.compile(r"\b(?:remove everything|clear slide|clear all|erase all|wipe slide|delete all content|remove all content)\b", _R)
_RE_BLANK        = re.compile(r"\b(?:make blank slide|blank slide|empty slide)\b", _R)
_RE_VIEW         = re.compile(r"\b(?:show|view|display|see|go to|jump to|navigate to)\s+slide\s*#?\d+\b", _R)
_RE_EXPLAIN      = re.compile(r"\b(?:explain|describe|summarize|walk me through)\b.{0,30}\bslide\s*#?\d+\b|\bslide\s*#?\d+\b.{0,20}\b(?:explain|describe|about)\b", _R)
_RE_DELETE       = re.compile(r"\b(?:delete|remove)\s+slide\s*#?\d+\b|\bslide\s*#?\d+\b.{0,10}\b(?:delete|remove)\b", _R)
_RE_MERGE        = re.compile(r"\b(?:merge|combine|join|consolidate)\s+slide", _R)
_RE_SWAP         = re.compile(r"\b(?:swap|exchange|switch\s+positions?)\s+slide", _R)
_RE_MOVE         = re.compile(r"\b(?:move|reorder|rearrange)\s+slide\b|\bslide\b.+\b(?:before|after)\s+slide\b", _R)
_RE_TRANSFORM    = re.compile(r"\b(?:improve|enhance|refine|shorten|simplify|rewrite|reword|fix|polish|condense|make\s+it\s+better|make\s+it\s+shorter|make\s+it\s+professional|make\s+it\s+cleaner|make\s+it\s+clearer|spice\s+up|upgrade|rephrase|tighten|clean\s+up|engaging|impactful)\b", _R)
_RE_ADD_POINTS   = re.compile(r"\b(?:add|append|insert)\b.{0,20}\b(?:point|points|bullet|bullets)\b|\bmore\s+points?\b|\bslide\s*#?\d+\b.{0,20}\b(?:longer|more detailed|more detail|expand)\b", _R)
_RE_BULK         = re.compile(r"\b(?:all slides?|every slide|each slide|entire (?:ppt|deck|presentation)|whole (?:ppt|deck|presentation)|throughout|across all)\b", _R)
_RE_REGEN        = re.compile(r"\b(?:regenerate|re-generate|redo|rebuild)\s+(?:slide|only slide|just slide)\b|\brewrite\s+(?:conclusion|introduction|intro|summary|thank you)\s+slide\b", _R)
_RE_DESIGN       = re.compile(r"\b(?:dark theme|dark mode|modern|minimal|minimalist|business presentation|investor|visually appealing|icons?|visuals?|impactful heading|use less text|visual\s+appeal)\b", _R)
_RE_PRES_MODE    = re.compile(r"\b(?:speaker notes?|slide script|script for each|audience question|questions? audience|audience.*questions?|q\s*&?\s*a|2[\s-]minute speech|two[\s-]minute|summarize.*speech|interactive question|question per slide|like i.m present|presenting|explain\s+(?:this\s+)?(?:ppt|presentation|deck)|walk\s+me\s+through\s+(?:this\s+)?(?:ppt|presentation|deck))\b", _R)
_RE_TRANSLATE    = re.compile(r"\b(?:translate|convert)\b.{0,30}\b(?:hindi|french|spanish|german|japanese|chinese|arabic|portuguese|urdu|bengali|tamil|telugu|kannada|marathi|gujarati|malayalam)\b", _R)
_RE_STATS        = re.compile(r"\b(?:include|add|insert)\b.{0,30}\b(?:statistics?|stats?|data|numbers?|figures?|facts?|percentage|metric|kpi)\b", _R)
_RE_EXAMPLES     = re.compile(r"\b(?:add|include|insert)\b.{0,30}\b(?:real[\s-]world examples?|examples?|case studi(?:es|y)|indian context|local context|context)\b", _R)
_RE_COMPARISON   = re.compile(r"\b(?:add|include|insert)\b.{0,30}\b(?:comparison table|table|chart|graph|advantages?\s+and\s+disadvantages?|pros?\s+and\s+cons?)\b", _R)
_RE_ADD_SLIDE    = re.compile(r"\b(?:add|insert)\b.{0,20}\b(?:new\s+)?slide\b|\bnew\s+slide\b|\banother\s+slide\b|\ba\s+(?:conclusion|summary|case study|intro|introduction|thank you|future trends?)\s+slide\b", _R)
_RE_CREATE       = re.compile(r"\b(?:make|create|generate|build)\b.{0,20}\b(?:ppt|presentation|deck)\b", _R)
_RE_SWITCH       = re.compile(r"\b(?:open|switch\s+to|go\s+to)\b.{0,20}\b(?:ppt|presentation|deck)\b|\bppt\s*\d+\b", _R)
_RE_PREVIEW      = re.compile(r"\bpreview\b.{0,20}\b(?:ppt|deck|presentation)\b|\bshow\s+(?:all|full)\s+(?:deck|slides)\b", _R)
_RE_DOWNLOAD     = re.compile(r"\b(?:download|export|save)\b.{0,20}\b(?:ppt|presentation|deck)\b", _R)
_RE_REFINE       = re.compile(r"\b(?:more professional|more formal|more casual|key points only|remove fluff|cleaner deck|change tone|tone to|suitable for beginners?|suitable for investors?)\b", _R)
_RE_GREETING     = re.compile(r"^(?:hi|hello|hey|howdy|greetings)[\s!.]*$", _R)
_RE_SMALLTALK    = re.compile(r"^(?:ok|okay|thanks|thank you|cool|got it|yes|no|sure|great|perfect|nice|alright)[\s!.]*$", _R)
_RE_RESIZE       = re.compile(
    r"\b(?:into|to|make(?:\s+this)?)\s+1\s+slide\b"
    r"|\b1[\s-]slide\s+summary\b"
    r"|\b(?:expand|resize|convert|turn)\b.{0,40}\b(?:to|into)?\s*\d+\s+slides?\b"
    r"|\b(?:ppt|presentation|deck)\b.{0,25}\b(?:to|into)\s+\d+\s+slides?\b",
    _R,
)
_RE_STRIP        = re.compile(r"\b(?:remove all text|keep only headings?|only headings?|headings? only|strip\s+text)\b", _R)


def _regex_fallback(text: str, explicit_slide: Optional[int], slide_id: Any) -> dict:
    """Conservative regex-only classification. Returns unknown if unsure."""

    if _RE_CANCEL.search(text):
        return {"intent": "cancel", "confidence": 0.99, "action_type": "meta"}

    if _RE_BLANK.search(text):
        return {"intent": "blank_slide", "confidence": 0.97, "action_type": "slide"}

    if _RE_CLEAR.search(text):
        return {"intent": "clear_slide", "confidence": 0.97, "action_type": "slide"}

    if _RE_EXPLAIN.search(text):
        return {"intent": "explain_slide", "confidence": 0.96, "action_type": "slide"}

    if _RE_VIEW.search(text):
        return {"intent": "view_slide", "confidence": 0.99, "action_type": "slide"}

    if _RE_MERGE.search(text):
        return {"intent": "merge_slides", "confidence": 0.95, "action_type": "slide"}

    if _RE_SWAP.search(text):
        return {"intent": "swap_slides", "confidence": 0.94, "action_type": "slide"}

    if _RE_DELETE.search(text):
        return {"intent": "delete_slide", "confidence": 0.95, "action_type": "slide"}

    if _RE_MOVE.search(text):
        return {"intent": "move_slide", "confidence": 0.93, "action_type": "slide"}

    if _RE_TRANSLATE.search(text):
        lang_m = re.search(r"\b(hindi|french|spanish|german|japanese|chinese|arabic|portuguese|urdu|bengali|tamil|telugu|kannada|marathi|gujarati|malayalam)\b", text, _R)
        return {"intent": "translate_ppt", "confidence": 0.96, "action_type": "deck", "target_language": lang_m.group(1).capitalize() if lang_m else None}

    if _RE_PRES_MODE.search(text):
        return {"intent": "presentation_mode", "confidence": 0.95, "action_type": "deck",
                "sub_intent": _classify_presentation_mode_sub_intent(text)}

    if _RE_DESIGN.search(text):
        return {"intent": "design_change", "confidence": 0.90, "action_type": "deck",
                "sub_intent": _classify_design_sub_intent(text)}

    if _RE_REGEN.search(text):
        return {"intent": "regenerate_slide", "confidence": 0.94, "action_type": "slide"}

    if _RE_RESIZE.search(text):
        return {"intent": "bulk_update", "confidence": 0.93, "action_type": "deck",
                "sub_intent": "resize", "slide_count_target": _extract_slide_count(text)}

    if _RE_STRIP.search(text):
        return {"intent": "bulk_update", "confidence": 0.93, "action_type": "deck", "sub_intent": "strip_text"}

    # Bulk operations (all/each/every slide)
    if _RE_BULK.search(text):
        sub = _classify_update_sub_intent(text)
        if _RE_STATS.search(text):
            sub = "add_statistics"
        elif _RE_EXAMPLES.search(text):
            sub = "add_examples"
        elif _RE_COMPARISON.search(text):
            sub = "add_comparison"
        return {"intent": "bulk_update", "confidence": 0.92, "action_type": "deck", "sub_intent": sub}

    if _RE_TRANSFORM.search(text):
        return {"intent": "transform_content", "confidence": 0.95, "action_type": "slide"}

    if _RE_ADD_POINTS.search(text):
        return {"intent": "update_slide", "confidence": 0.92, "action_type": "slide", "sub_intent": "add_content"}

    # Content enhancement (stats/examples/comparison) — with no bulk modifier
    if _RE_STATS.search(text):
        return {"intent": "update_slide", "confidence": 0.88, "action_type": "slide", "sub_intent": "add_statistics"}

    if _RE_EXAMPLES.search(text):
        return {"intent": "update_slide", "confidence": 0.88, "action_type": "slide", "sub_intent": "add_examples"}

    if _RE_COMPARISON.search(text):
        return {"intent": "update_slide", "confidence": 0.88, "action_type": "slide", "sub_intent": "add_comparison"}

    if _RE_REFINE.search(text):
        return {"intent": "refine_ppt", "confidence": 0.88, "action_type": "deck"}

    if _RE_DOWNLOAD.search(text):
        return {"intent": "download_ppt", "confidence": 0.95, "action_type": "deck"}

    if _RE_PREVIEW.search(text):
        return {"intent": "preview_ppt", "confidence": 0.93, "action_type": "deck"}

    if _RE_CREATE.search(text):
        return {"intent": "create_ppt", "confidence": 0.95, "action_type": "deck"}

    if _RE_SWITCH.search(text):
        return {"intent": "switch_ppt", "confidence": 0.94, "action_type": "deck"}

    if _RE_GREETING.search(text):
        return {"intent": "greeting", "confidence": 0.99, "action_type": "meta"}

    if _RE_SMALLTALK.search(text):
        return {"intent": "smalltalk", "confidence": 0.99, "action_type": "meta"}

    # add_slide — only when clearly a new slide request
    if _RE_ADD_SLIDE.search(text) and not re.search(r"\b(?:edit|change|update|modify)\s+slide\s*#?\d+\b", text, _R):
        return {"intent": "add_slide", "confidence": 0.92, "action_type": "deck"}

    # Slide-level update with active context
    if re.search(r"\b(?:include|add|update|change|modify|improve|expand|elaborate)\b", text, _R):
        if explicit_slide is not None or slide_id is not None:
            return {"intent": "update_slide", "confidence": 0.72, "action_type": "slide",
                    "sub_intent": _classify_update_sub_intent(text)}

    # Safe default
    return {
        "intent": "unknown", "confidence": 0.30, "action_type": "unknown",
        "clarification_question": "Could you clarify what you'd like me to do with the presentation?",
    }


# ─────────────────────────────────────────────────────────────────────────────
#  Maps
# ─────────────────────────────────────────────────────────────────────────────

_TARGET_MAP = {
    "transform_content":  "existing_slide",
    "update_slide":       "existing_slide",
    "edit_slide":         "existing_slide",
    "add_points":         "existing_slide",
    "view_slide":         "existing_slide",
    "explain_slide":      "existing_slide",
    "delete_slide":       "existing_slide",
    "merge_slides":       "existing_slide",
    "move_slide":         "existing_slide",
    "swap_slides":        "existing_slide",
    "clear_slide":        "existing_slide",
    "blank_slide":        "existing_slide",
    "regenerate_slide":   "existing_slide",
    "add_slide":          "new_slide",
    "bulk_update":        "all_slides",
    "design_change":      "deck",
    "presentation_mode":  "deck",
    "translate_ppt":      "deck",
    "refine_ppt":         "deck",
    "preview_ppt":        "deck",
    "switch_ppt":         "deck",
    "download_ppt":       "deck",
    "create_ppt":         "deck",
    "ppt_info":           "deck",
    "suggest_topic":      "meta",
    "cancel":             "meta",
    "greeting":           "conversation",
    "smalltalk":          "conversation",
    "unknown":            "unknown",
}

_ACTION_TYPE_MAP = {
    "view_slide":        "slide",
    "explain_slide":     "slide",
    "delete_slide":      "slide",
    "merge_slides":      "slide",
    "swap_slides":       "slide",
    "move_slide":        "slide",
    "update_slide":      "slide",
    "edit_slide":        "slide",
    "transform_content": "slide",
    "add_points":        "slide",
    "clear_slide":       "slide",
    "blank_slide":       "slide",
    "regenerate_slide":  "slide",
    "add_slide":         "deck",
    "bulk_update":       "deck",
    "design_change":     "deck",
    "presentation_mode": "deck",
    "translate_ppt":     "deck",
    "create_ppt":        "deck",
    "preview_ppt":       "deck",
    "switch_ppt":        "deck",
    "download_ppt":      "deck",
    "ppt_info":          "deck",
    "refine_ppt":        "deck",
    "suggest_topic":     "meta",
    "cancel":            "meta",
    "greeting":          "meta",
    "smalltalk":         "meta",
    "unknown":           "unknown",
}

# Intents that require a specific slide_id
_SLIDE_INTENTS = {
    "transform_content", "add_points", "update_slide", "edit_slide",
    "view_slide", "explain_slide", "delete_slide", "merge_slides",
    "move_slide", "swap_slides", "clear_slide", "blank_slide",
    "regenerate_slide",
}

# All valid intents
_ALL_VALID_INTENTS = set(_TARGET_MAP.keys())


def _safe_intent(raw: str, fallback: str = "unknown") -> str:
    return raw if raw in _ALL_VALID_INTENTS else fallback


# ─────────────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────────────

def classify_ppt_intent(
    user_input: str,
    ppt_id: Any = None,
    slide_id: Any = None,
) -> dict:
    """
    LLM-first intent classification with comprehensive natural language support.
    Falls back to regex when LLM unavailable.
    NEVER defaults to add_slide for unclear input.
    """
    text = _norm(user_input)

    explicit_ppt   = _extract_ppt_id(text)
    explicit_slide = _extract_slide_id(text)
    all_slide_ids  = _extract_slide_ids(text)

    result: dict = {
        "intent":                 "unknown",
        "ppt_id":                 explicit_ppt if explicit_ppt is not None else ppt_id,
        "slide_id":               explicit_slide,
        "slide_number":           explicit_slide,
        "slide_ids":              all_slide_ids,
        "content":                None,
        "operation":              None,
        "target":                 "unknown",
        "action_type":            "unknown",
        "sub_intent":             None,
        "target_language":        None,
        "slide_count_target":     None,
        "position":               None,
        "anchor_slide":           None,
        "missing_entities":       [],
        "clarification_question": None,
        "confidence":             0.30,
    }

    if not text:
        result["clarification_question"] = "What would you like me to do with the presentation?"
        return result

    # ── 1. LLM classification ─────────────────────────────────────────────────
    llm = _llm_classify(text, ppt_id, slide_id)

    if llm and isinstance(llm, dict) and llm.get("intent"):
        intent = _safe_intent(_normalize_intent(llm.get("intent", "")))
        conf   = float(llm.get("confidence") or 0.70)
        if conf < 0.60:
            intent = "unknown"

        result.update({
            "intent":                 intent,
            "confidence":             conf,
            "operation":              llm.get("operation") or intent,
            "target":                 _TARGET_MAP.get(intent, "unknown"),
            "action_type":            llm.get("action_type") or _ACTION_TYPE_MAP.get(intent, "unknown"),
            "sub_intent":             llm.get("sub_intent") or _classify_update_sub_intent(text) if intent in {
                "update_slide", "transform_content", "add_points", "bulk_update"
            } else llm.get("sub_intent"),
            "target_language":        llm.get("target_language"),
            "slide_count_target":     llm.get("slide_count_target"),
            "position":               llm.get("position"),
            "anchor_slide":           llm.get("anchor_slide"),
            "clarification_question": llm.get("clarification_question"),
            "missing_entities":       llm.get("missing_entities") if isinstance(llm.get("missing_entities"), list) else [],
            "content":                llm.get("content") or _extract_content(text),
        })

        # presentation_mode sub_intent
        if intent == "presentation_mode" and not result["sub_intent"]:
            result["sub_intent"] = _classify_presentation_mode_sub_intent(text)

        # design_change sub_intent
        if intent == "design_change" and not result["sub_intent"]:
            result["sub_intent"] = _classify_design_sub_intent(text)

        # Slide ID resolution
        llm_slide = llm.get("slide_id")
        if explicit_slide is not None:
            result["slide_id"] = explicit_slide
        elif llm_slide is not None:
            try:
                result["slide_id"] = int(llm_slide)
            except (TypeError, ValueError):
                pass
        result["slide_number"] = result["slide_id"]

        # Merge slide_ids
        merged = list(all_slide_ids)
        for sid in (llm.get("slide_ids") or []):
            try:
                v = int(sid)
                if v not in merged:
                    merged.append(v)
            except (TypeError, ValueError):
                pass
        result["slide_ids"] = merged or ([result["slide_id"]] if result["slide_id"] else [])

        # PPT ID resolution
        if explicit_ppt:
            result["ppt_id"] = explicit_ppt
        elif llm.get("ppt_id"):
            result["ppt_id"] = llm["ppt_id"]

        # Boost confidence for view_slide with explicit number
        if intent == "view_slide" and explicit_slide is not None:
            result["confidence"] = 0.99

        # add_slide: parse placement, clear slide_id
        if intent == "add_slide":
            if not result["position"]:
                pos_m = re.search(r"\b(after|before)\s+slide\s*#?(\d+)\b|\bat\s+the\s+(start|beginning|end)\b", text, _R)
                if pos_m:
                    if pos_m.group(1):
                        result["position"] = pos_m.group(1).lower()
                        result["anchor_slide"] = int(pos_m.group(2))
                    elif pos_m.group(3):
                        p = pos_m.group(3).lower()
                        result["position"] = "start" if p in {"start", "beginning"} else "end"
            result["slide_id"] = None
            result["slide_number"] = None
            miss = []
            if not result.get("content"):
                miss.append("slide_content")
            if not result.get("position"):
                miss.append("position")
            result["missing_entities"] = miss

        # Slide intents without slide_id → downgrade or use context
        if intent in _SLIDE_INTENTS and result["slide_id"] is None:
            if slide_id is not None:
                result["slide_id"] = slide_id
                result["slide_number"] = slide_id
                result["confidence"] = max(0.65, result["confidence"] - 0.10)
            else:
                result["missing_entities"] = sorted(set(result.get("missing_entities", []) + ["slide_number"]))
                result["confidence"] = min(result["confidence"], 0.55)
                result["intent"] = "unknown"
                result["clarification_question"] = result.get("clarification_question") or "Which slide should I work on?"

        return result

    # ── 2. Regex fallback ─────────────────────────────────────────────────────
    fb = _regex_fallback(text, explicit_slide, slide_id)
    intent = _safe_intent(_normalize_intent(fb.get("intent", "unknown")))

    result.update({
        "intent":              intent,
        "confidence":          float(fb.get("confidence", 0.30)),
        "operation":           fb.get("operation") or intent,
        "target":              _TARGET_MAP.get(intent, "unknown"),
        "action_type":         fb.get("action_type") or _ACTION_TYPE_MAP.get(intent, "unknown"),
        "sub_intent":          fb.get("sub_intent"),
        "target_language":     fb.get("target_language"),
        "slide_count_target":  fb.get("slide_count_target"),
        "content":             _extract_content(text),
        "clarification_question": fb.get("clarification_question"),
    })

    if intent == "presentation_mode" and not result["sub_intent"]:
        result["sub_intent"] = _classify_presentation_mode_sub_intent(text)

    if intent == "design_change" and not result["sub_intent"]:
        result["sub_intent"] = _classify_design_sub_intent(text)

    # Resolve slide context
    if intent in _SLIDE_INTENTS:
        if explicit_slide is not None:
            result["slide_id"] = explicit_slide
            result["slide_number"] = explicit_slide
        elif slide_id is not None:
            result["slide_id"] = slide_id
            result["slide_number"] = slide_id
            result["confidence"] = max(0.65, result["confidence"] - 0.10)
        else:
            result["missing_entities"] = ["slide_number"]
            result["intent"] = "unknown"
            result["confidence"] = 0.30
            result["clarification_question"] = "Which slide should I work on?"

    if intent == "add_slide":
        pos_m = re.search(r"\b(after|before)\s+slide\s*#?(\d+)\b|\bat\s+the\s+(start|beginning|end)\b", text, _R)
        if pos_m:
            if pos_m.group(1):
                result["position"] = pos_m.group(1).lower()
                result["anchor_slide"] = int(pos_m.group(2))
            elif pos_m.group(3):
                p = pos_m.group(3).lower()
                result["position"] = "start" if p in {"start", "beginning"} else "end"
        result["slide_id"] = None
        result["slide_number"] = None
        if not result.get("content"):
            result["missing_entities"].append("slide_content")
        if not result.get("position"):
            result["missing_entities"].append("position")

    if intent == "merge_slides":
        ids = _extract_slide_ids(text)
        result["slide_ids"] = ids
        if len(ids) < 2:
            result["missing_entities"] = ["slide_numbers"]
            result["intent"] = "unknown"
            result["confidence"] = 0.40
            result["clarification_question"] = "Which two slides should I merge?"

    if intent == "translate_ppt" and not result.get("target_language"):
        result["missing_entities"] = ["target_language"]
        result["clarification_question"] = "Which language should I translate the presentation into?"

    result["slide_ids"] = result["slide_ids"] or ([result["slide_id"]] if result["slide_id"] else [])
    return result
