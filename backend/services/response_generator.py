"""
response_generator.py
=====================

Hybrid response wording for the PPT assistant.

The action/routing layer stays authoritative. This module only turns completed
actions into short, natural confirmations:

1. Build a deterministic base message from templates.
2. Pick a light phrasing variation.
3. Optionally ask the LLM to polish wording for complex edits only.
"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass
from typing import Any, Optional


SIMPLE_ACTIONS = {
    "add_slide",
    "delete_slide",
    "move_slide",
    "swap_slides",
    "merge_slides",
    "clear_slide",
    "blank_slide",
    "view_slide",
    "download_ppt",
    "preview_ppt",
    "ppt_info",
    "smalltalk",
    "greeting",
}

LLM_ELIGIBLE_ACTIONS = {
    "update_slide",
    "edit_slide",
    "transform_content",
    "regenerate_slide",
}

_LAST_BY_ACTION: dict[str, str] = {}


@dataclass(frozen=True)
class ResponseContext:
    action: str
    success: bool = True
    slide_number: Optional[int] = None
    slide_title: str = ""
    ppt_topic: str = ""
    ppt_label: str = "the deck"
    change_type: str = ""
    change_content: str = ""
    n_points: Optional[int] = None
    user_request: str = ""
    use_llm: Optional[bool] = None


def generate_response(
    action: str,
    *,
    success: bool = True,
    slide_number: Optional[int] = None,
    slide_title: str = "",
    ppt_topic: str = "",
    ppt_label: str = "the deck",
    change_type: str = "",
    change_content: str = "",
    n_points: Optional[int] = None,
    user_request: str = "",
    use_llm: Optional[bool] = None,
    llm_client: Any = None,
    llm_model: Optional[str] = None,
) -> str:
    """
    Return a short assistant response for a completed action.

    `llm_client` is optional. When omitted, the function still returns a varied
    template. When provided, the LLM is used only to polish wording for complex
    edit actions.
    """
    ctx = ResponseContext(
        action=_normalize_action(action),
        success=success,
        slide_number=slide_number,
        slide_title=slide_title.strip(),
        ppt_topic=ppt_topic.strip(),
        ppt_label=ppt_label.strip() or "the deck",
        change_type=_infer_change_type(change_type, change_content, action),
        change_content=change_content.strip(),
        n_points=n_points,
        user_request=user_request.strip(),
        use_llm=use_llm,
    )
    base = _base_response(ctx)
    varied = _vary_response(base, ctx)
    if _should_use_llm(ctx) and llm_client and llm_model:
        return _enhance_with_llm(varied, ctx, llm_client, llm_model)
    return varied


def _normalize_action(action: str) -> str:
    value = str(action or "").strip().lower()
    if value == "add_points":
        return "update_slide"
    return value or "unknown"


def _slide_ref(ctx: ResponseContext) -> str:
    if ctx.slide_title:
        return f"slide {ctx.slide_number} ({ctx.slide_title})" if ctx.slide_number else ctx.slide_title
    if ctx.slide_number:
        return f"slide {ctx.slide_number}"
    return "the slide"


def _infer_change_type(change_type: str, change_content: str, action: str) -> str:
    if change_type:
        return change_type
    text = f"{action} {change_content}".casefold()
    if re.search(r"\b(shorten|shorter|condense|trim|concise|less text)\b", text):
        return "shortened"
    if re.search(r"\b(rewrite|reword|rephrase|polish|improve|enhance|professional)\b", text):
        return "rewritten"
    if re.search(r"\b(examples?|case stud(?:y|ies)|context)\b", text):
        return "added examples"
    if re.search(r"\b(stats?|statistics?|data|metrics?|numbers?|figures?)\b", text):
        return "added data"
    if re.search(r"\b(point|bullet)\b", text):
        return "added points"
    if "regenerate" in text:
        return "regenerated"
    return "updated"


def _base_response(ctx: ResponseContext) -> str:
    if not ctx.success:
        return "I could not complete that request."

    slide = _slide_ref(ctx)
    topic_hint = f" in {ctx.ppt_topic}" if ctx.ppt_topic else ""
    points = ctx.n_points if ctx.n_points is not None else 1

    if ctx.action in {"update_slide", "edit_slide", "transform_content"}:
        if ctx.change_type == "added points":
            return f"Added {points} point{'s' if points != 1 else ''} to {slide}."
        if ctx.change_type == "shortened":
            return f"Shortened {slide} and kept the message focused."
        if ctx.change_type == "rewritten":
            return f"Reworked {slide} so it reads more clearly."
        if ctx.change_type == "added examples":
            return f"Added relevant examples to {slide}."
        if ctx.change_type == "added data":
            return f"Added supporting data to {slide}."
        return f"Updated {slide}."

    if ctx.action == "regenerate_slide":
        return f"Regenerated {slide}."
    if ctx.action == "add_slide":
        return f"Added the new slide to {ctx.ppt_label}{topic_hint}."
    if ctx.action == "delete_slide":
        return f"Removed {slide} from {ctx.ppt_label}."
    if ctx.action == "move_slide":
        return f"Moved {slide} in {ctx.ppt_label}."
    if ctx.action == "swap_slides":
        return f"Swapped the selected slides in {ctx.ppt_label}."
    if ctx.action == "merge_slides":
        return f"Merged the selected slides in {ctx.ppt_label}."
    if ctx.action == "clear_slide":
        return f"Cleared {slide}."
    if ctx.action == "blank_slide":
        return f"{slide.capitalize()} is now blank."
    if ctx.action == "create_ppt":
        return f"Built the presentation{topic_hint}."
    if ctx.action == "refine_ppt":
        return f"Refreshed {ctx.ppt_label}{topic_hint}."
    if ctx.action == "view_slide":
        return f"Showing {slide}."
    if ctx.action == "explain_slide":
        return f"Here is a clear explanation of {slide}."
    return "Done."


def _vary_response(base: str, ctx: ResponseContext) -> str:
    variants = _template_variants(ctx)
    if not variants:
        return base
    last = _LAST_BY_ACTION.get(ctx.action)
    choices = [v for v in variants if v != last] or variants
    selected = random.choice(choices)
    _LAST_BY_ACTION[ctx.action] = selected
    return selected


def _template_variants(ctx: ResponseContext) -> list[str]:
    slide = _slide_ref(ctx)
    ppt_label = ctx.ppt_label
    points = ctx.n_points if ctx.n_points is not None else 1

    if ctx.action in {"update_slide", "edit_slide", "transform_content"}:
        if ctx.change_type == "added points":
            noun = "point" if points == 1 else "points"
            return [
                f"Added {points} {noun} to {slide}.",
                f"{slide.capitalize()} now has {points} new {noun}.",
                f"Done, I added {points} {noun} to {slide}.",
            ]
        if ctx.change_type == "shortened":
            return [
                f"Shortened {slide} and kept the main idea intact.",
                f"Done, {slide} is tighter now.",
                f"I trimmed {slide} so it reads more cleanly.",
            ]
        if ctx.change_type == "rewritten":
            return [
                f"Reworked {slide} so it reads more clearly.",
                f"Done, {slide} has a cleaner version now.",
                f"I polished the wording on {slide}.",
            ]
        if ctx.change_type == "added examples":
            return [
                f"Added relevant examples to {slide}.",
                f"Done, {slide} now includes examples.",
                f"I worked examples into {slide}.",
            ]
        return [
            f"Updated {slide}.",
            f"Done, {slide} has been updated.",
            f"I made that change on {slide}.",
        ]

    simple_variants = {
        "regenerate_slide": [
            f"Regenerated {slide}.",
            f"Done, {slide} has a fresh version now.",
            f"I rebuilt {slide} with updated wording.",
        ],
        "add_slide": [
            f"Added the new slide to {ppt_label}.",
            f"Done, the new slide is in {ppt_label}.",
            f"I added that slide to {ppt_label}.",
        ],
        "delete_slide": [
            f"Removed {slide} from {ppt_label}.",
            f"Done, {slide} is out of {ppt_label}.",
            f"I deleted {slide} from {ppt_label}.",
        ],
        "clear_slide": [
            f"Cleared {slide}.",
            f"Done, {slide} has been cleared.",
            f"I removed the content from {slide}.",
        ],
        "blank_slide": [
            f"{slide.capitalize()} is now blank.",
            f"Done, {slide} is blank now.",
            f"I cleared {slide} into a blank slide.",
        ],
    }
    return simple_variants.get(ctx.action, [])


def _should_use_llm(ctx: ResponseContext) -> bool:
    if ctx.use_llm is not None:
        return bool(ctx.use_llm)
    if ctx.action in SIMPLE_ACTIONS:
        return False
    if ctx.action not in LLM_ELIGIBLE_ACTIONS:
        return False
    text = f"{ctx.change_type} {ctx.change_content} {ctx.user_request}".casefold()
    return bool(
        re.search(r"\b(rewrite|improve|enhance|polish|example|context|vague|make it|professional|shorten)\b", text)
        or len(text.split()) > 10
    )


def _enhance_with_llm(base: str, ctx: ResponseContext, llm_client: Any, llm_model: str) -> str:
    prompt = f"""
Rewrite this assistant confirmation to sound natural and varied.

Base confirmation: {base}
Action: {ctx.action}
Slide number: {ctx.slide_number or "none"}
Slide title: {ctx.slide_title or "none"}
PPT topic: {ctx.ppt_topic or "none"}
Change type: {ctx.change_type or "none"}

Rules:
- Do not change the action outcome.
- Do not add new facts.
- Do not mention that an LLM was used.
- Keep it under 22 words.
- Return one sentence only.
"""
    try:
        resp = llm_client.chat.completions.create(
            model=llm_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.45,
            max_tokens=60,
        )
        text = (resp.choices[0].message.content or "").strip()
        text = re.sub(r"^['\"]|['\"]$", "", text).strip()
        if not text or len(text.split()) > 28:
            return base
        return text
    except Exception:
        return base
