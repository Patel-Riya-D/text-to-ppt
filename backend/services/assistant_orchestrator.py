"""
assistant_orchestrator.py

Small, framework-neutral routing helpers for the PPT assistant.

This module is the first extraction toward a LangGraph-style orchestrator:
it keeps conversational policy and pre-routing decisions outside Streamlit,
while app.py still owns UI, session state, and mutation executors.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import re
from typing import Any, Optional


def norm_text(value: Any) -> str:
    text = str(value or "").casefold()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def safe_str(value: Any, default: str = "") -> str:
    return str(value).strip() if value is not None else default


def extract_first_json_object(text: str) -> dict[str, Any]:
    raw = str(text or "").strip()
    start = raw.find("{")
    if start == -1:
        raise ValueError("No JSON object found")
    depth = 0
    in_string = False
    escape = False
    for idx, ch in enumerate(raw[start:], start=start):
        if escape:
            escape = False
            continue
        if ch == "\\":
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
        if in_string:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                parsed = json.loads(raw[start : idx + 1])
                if not isinstance(parsed, dict):
                    raise ValueError("JSON root is not an object")
                return parsed
    raise ValueError("Unbalanced JSON object")


@dataclass
class AssistantDecision:
    mode: str
    intent: str = "unknown"
    slots: dict[str, Any] = field(default_factory=dict)
    message: Optional[str] = None
    confidence: float = 0.0
    reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "intent": self.intent,
            "slots": dict(self.slots),
            "message": self.message,
            "confidence": self.confidence,
            "reason": self.reason,
        }


def is_generic_presentation_context(value: str) -> bool:
    norm = norm_text(value)
    if not norm:
        return True
    norm = re.sub(r"\bclg\b", "college", norm)
    generic_patterns = [
        r"^(?:me|myself|for me|please|pls)\s*$",
        r"^(?:my|a|the)?\s*(?:college|school|class|seminar|project|assignment|presentation|ppt|deck)\s*$",
        r"^(?:my|a|the)?\s*(?:college|school|class)\s+(?:seminar|project|assignment|presentation)\s*$",
        r"^(?:my|a|the)?\s*(?:college|school|class)\s+(?:ppt|deck|slides?)\s*$",
        r"^(?:seminar|project|assignment|class|college)\s+(?:ppt|presentation|deck)\s*$",
        r"^(?:my|a|the)?\s*(?:college|school|class|seminar|project|assignment)\s+(?:ppt|presentation|deck|slides?)\s*$",
        r"^(?:something|anything)\s+(?:to\s+)?(?:show|present)(?:\s+(?:in|for)\s+(?:college|school|class|seminar))?\s*$",
        r"^(?:something|anything)\s+for\s+(?:my\s+)?(?:college|school|class|seminar|presentation)\s*$",
        r"^(?:show|present)\s+(?:in|for)\s+(?:college|school|class|seminar)\s*$",
        r"^(?:another|different|new|other|some other|any other)\s+topic(?:\s+for\s+me)?\s*$",
        r"^(?:another|different|new|other|some other|any other)\s+(?:ppt|presentation|deck)(?:\s+for\s+me)?\s*$",
    ]
    return any(re.fullmatch(pattern, norm, re.IGNORECASE) for pattern in generic_patterns)


def is_implicit_create_request(user_input: str) -> bool:
    text = str(user_input or "")
    if not text:
        return False
    return bool(
        re.search(
            r"\b(?:help\s+me\s+(?:to\s+)?(?:create|make|build)|can you help me|i need|need|want|looking for)\b"
            r".{0,60}\b(?:ppt|presentation|deck|slides|seminar|project|assignment)\b",
            text,
            re.IGNORECASE,
        )
    )


def extract_conversational_create_topic(
    user_input: str,
    *,
    explicit_topic: Optional[str] = None,
    pending_need_topic: bool = False,
) -> Optional[str]:
    text = str(user_input or "").strip()
    if not text:
        return None
    if explicit_topic and not is_generic_presentation_context(explicit_topic):
        return explicit_topic

    patterns = [
        r"\b(?:seminar|project|assignment|class|college)\s+(?:on|about|regarding)\s+(.+)$",
        r"\b(?:ppt|presentation|deck|slides?)\s+(?:on|about|regarding)\s+(.+)$",
        r"\b(?:topic|subject)\s+(?:is|:)\s+(.+)$",
        r"\b(?:i was thinking|i am thinking|thinking|maybe|let'?s do it on|it is about|it's about|the subject is|topic will be)\s+(.+)$",
        r"^(?:on|about)\s+(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if not match:
            continue
        topic = re.sub(
            r"\b(?:with|for|of)\s+\d+\s+slides?\b.*$",
            "",
            match.group(1),
            flags=re.IGNORECASE,
        )
        topic = re.sub(
            r"\bfor\s+(?:my|a|the)?\s*(?:college|clg|school|class|seminar|project|assignment|presentation)\b.*$",
            "",
            topic,
            flags=re.IGNORECASE,
        )
        topic = re.sub(r"[\.\!\?]+$", "", topic).strip(" ,.;:-\"'")
        topic = re.sub(r"\s+", " ", topic).strip()
        if topic and not is_generic_presentation_context(topic):
            return topic

    if pending_need_topic:
        topic = re.sub(
            r"^\s*(?:i was thinking|i am thinking|thinking|maybe|let'?s do|let'?s go with|go with|use|topic is|subject is)\s+",
            "",
            text,
            flags=re.IGNORECASE,
        )
        topic = topic.strip(" ,.;:-\"'")
        if topic and topic != text and not is_generic_presentation_context(topic):
            return topic
        if (
            1 <= len(text.split()) <= 12
            and not re.search(
                r"\b(?:slide|slides|ppt|presentation|deck|seminar|project|assignment|help|create|make|generate|build)\b",
                text,
                re.IGNORECASE,
            )
        ):
            return text.strip(" ,.;:-\"'")
    return None


def _extract_slide_count_from_text(user_input: str) -> Optional[int]:
    text = str(user_input or "")
    match = re.search(r"\b(\d{1,2})\s*(?:slides?|pages?)?\b", text, re.IGNORECASE)
    if not match:
        return None
    try:
        count = int(match.group(1))
    except ValueError:
        return None
    return count if 1 <= count <= 60 else None


def create_topic_followup(user_input: str = "") -> str:
    text = str(user_input or "")
    if re.search(r"\bseminar\b", text, re.IGNORECASE):
        return "Sure - what topic is your seminar about?"
    if re.search(r"\b(?:college|class|school|clg)\b", text, re.IGNORECASE):
        return "Got it - what topic should the presentation cover?"
    if re.search(r"\b(?:project|assignment)\b", text, re.IGNORECASE):
        return "Of course - what is the topic for your project presentation?"
    return "Of course! What topic should the presentation be about?"


def has_explicit_deterministic_command(user_input: str) -> bool:
    text = str(user_input or "")
    if not text:
        return False
    explicit_patterns = [
        r"\b(?:show|view|display|go to|jump to|navigate to)\s+slide\s*#?\d+\b",
        r"\b(?:delete|remove)\s+slide\s*#?\d+\b",
        r"\b(?:merge|combine)\s+slide\s*#?\d+.*\bslide\s*#?\d+\b",
        r"\b(?:swap|exchange)\s+slide\s*#?\d+.*\bslide\s*#?\d+\b",
        r"\b(?:move|reorder)\s+slide\s*#?\d+\b",
        r"\b(?:download|export)\b.*\b(?:ppt|presentation|deck)\b",
        r"\b(?:create|make|generate|build)\b.*\b(?:ppt|presentation|deck)\b",
        r"\b(?:add|insert)\b.*\b(?:new\s+|another\s+|one more\s+)?slide\b",
    ]
    return any(re.search(pattern, text, re.IGNORECASE) for pattern in explicit_patterns)

def is_create_deck_request(user_input: str) -> bool:
    text = str(user_input or "")
    if not text:
        return False
    return bool(
        re.search(r"\b(?:create|make|generate|build|draft|prepare)\b", text, re.IGNORECASE)
        and re.search(r"\b(?:ppt|presentation|deck|slides?)\b", text, re.IGNORECASE)
    )

def is_topic_suggestion_request(user_input: str) -> bool:
    text = str(user_input or "")
    return bool(
        re.search(r"\b(?:suggest|recommend|ideas?|topics?|options?|choose from)\b", text, re.IGNORECASE)
        and re.search(r"\b(?:topic|ppt|presentation|deck|slides?)\b", text, re.IGNORECASE)
    )

def should_skip_open_llm_router(user_input: str) -> bool:
    """
    Keep exact operational commands deterministic, but let create-deck language
    go through the LLM router because topic/missing-field inference is natural
    language heavy.
    """
    text = str(user_input or "")
    if not text:
        return True
    if is_create_deck_request(text):
        return False
    return has_explicit_deterministic_command(text)


def is_contextual_followup_candidate(user_input: str) -> bool:
    text = str(user_input or "").strip()
    if not text or should_skip_open_llm_router(text):
        return False
    if len(text.split()) <= 8 and re.search(
        r"\b(this|that|it|one|looks good|good|boring|shorter|improve|modern|style|actually|instead|not that|wrong|yes|no|ok|okay)\b",
        text,
        re.IGNORECASE,
    ):
        return True
    return bool(
        re.search(
            r"\b(?:looks good|actually|instead|not that one|not this one|wrong one|make it|make this|"
            r"this feels|feels boring|boring|go with|modern style|improve it|improve this|"
            r"change this|shorten it|shorter|simplify it|better|cleaner)\b",
            text,
            re.IGNORECASE,
        )
    )


def fallback_contextual_intent(user_input: str, assistant_state: dict[str, Any]) -> Optional[dict[str, Any]]:
    text = safe_str(user_input)
    active_slide = assistant_state.get("active_slide_index")
    slide_target = active_slide if active_slide and assistant_state.get("active_slide") else None
    if re.search(r"\b(?:looks good|good|ok|okay|fine|perfect)\b", text, re.IGNORECASE):
        return {"intent": "smalltalk", "confidence": 0.86, "content": text, "clarification_question": None}
    if re.search(r"\b(?:not that one|not this one|wrong one)\b", text, re.IGNORECASE):
        return {
            "intent": "unknown",
            "confidence": 0.55,
            "content": text,
            "clarification_question": "Which slide or PPT did you mean?",
        }
    if re.search(r"\b(?:modern style|modern design|go with modern|minimal style|minimal design)\b", text, re.IGNORECASE):
        return {"intent": "design_change", "confidence": 0.88, "content": text, "sub_intent": "modern_design"}
    if re.search(r"\b(?:shorter|shorten|simplify|concise|less text)\b", text, re.IGNORECASE):
        if slide_target:
            return {"intent": "transform_content", "slide_id": slide_target, "confidence": 0.84, "content": text}
        return {"intent": "bulk_update", "sub_intent": "shorten", "confidence": 0.82, "content": text}
    if re.search(r"\b(?:improve|better|boring|change this|actually change)\b", text, re.IGNORECASE):
        if slide_target:
            return {"intent": "transform_content", "slide_id": slide_target, "confidence": 0.82, "content": text}
        return {"intent": "refine_ppt", "confidence": 0.78, "content": text, "style_hint": text}
    return None


def resolve_contextual_intent(
    user_input: str,
    assistant_state: dict[str, Any],
    *,
    llm_client: Any = None,
    deployment: Optional[str] = None,
) -> Optional[dict[str, Any]]:
    """
    Contextual resolver for natural follow-ups. The LLM dependency is injected
    by app.py today; a future LangGraph node can pass the same dependency.
    """
    if not is_contextual_followup_candidate(user_input):
        return None
    if llm_client is None or not deployment:
        return fallback_contextual_intent(user_input, assistant_state)

    prompt = f"""
You are a contextual intent resolver for a PowerPoint assistant.

User message:
{json.dumps(user_input)}

Compact assistant state:
{json.dumps(assistant_state, ensure_ascii=True)}

Return ONLY valid JSON with these fields:
{{
  "intent": "smalltalk|unknown|edit_slide|update_slide|transform_content|add_points|add_slide|design_change|bulk_update|refine_ppt|view_slide|explain_slide|delete_slide|regenerate_slide|cancel",
  "target_level": "conversation|slide|deck|new_slide|unknown",
  "slide_id": integer_or_null,
  "ppt_id": string_or_null,
  "content": string_or_null,
  "sub_intent": string_or_null,
  "confidence": number_between_0_and_1,
  "clarification_question": string_or_null
}}

Rules:
- Do not invent slide numbers. Use active_slide_index only when the message says "this", "it", "that", or clearly continues the last slide action.
- If user approves with "looks good", "ok", or similar, intent=smalltalk.
- If user says "make it shorter" and an active slide exists, target that slide. If no active slide exists, target the deck with intent=bulk_update and sub_intent=shorten.
- If user says "go with modern style", intent=design_change. Preserve content.
- If user says "this feels boring" or "can you improve it", choose active slide if available, otherwise deck refinement.
- If user says "not that one", ask a clarification.
- confidence >= 0.80 only when the target is clear.
"""
    try:
        resp = llm_client.chat.completions.create(
            model=deployment,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=420,
        )
        raw = safe_str(resp.choices[0].message.content)
        parsed = extract_first_json_object(raw)
        intent = safe_str(parsed.get("intent", "unknown"), "unknown")
        confidence = float(parsed.get("confidence") or 0.0)
        return {
            "intent": intent,
            "slide_id": parsed.get("slide_id"),
            "slide_ids": [],
            "ppt_id": parsed.get("ppt_id") or assistant_state.get("active_ppt_id"),
            "operation": parsed.get("sub_intent") or intent,
            "content": parsed.get("content") or user_input,
            "action_type": parsed.get("target_level"),
            "target": parsed.get("target_level"),
            "sub_intent": parsed.get("sub_intent"),
            "missing_entities": [],
            "clarification_question": parsed.get("clarification_question"),
            "confidence": max(0.0, min(1.0, confidence)),
            "contextual": True,
        }
    except Exception:
        return fallback_contextual_intent(user_input, assistant_state)


def is_readonly_conversation_request(user_input: str) -> bool:
    text = str(user_input or "").strip()
    if not text:
        return False
    if re.search(
        r"\b(?:delete|remove|add|insert|replace|apply|do it|make the change|change it to|"
        r"regenerate|redo|rebuild|move|swap|merge|create|generate|build)\b",
        text,
        re.IGNORECASE,
    ) and not re.search(
        r"\b(?:suggest|suggestion|feedback|review|how can|what should|what can|is this|is my|why|explain)\b",
        text,
        re.IGNORECASE,
    ):
        return False
    return bool(
        re.search(
            r"\b(?:how can i|how should i|what should i|what can i|any suggestions?|suggest improvements?|"
            r"give me feedback|feedback on|review slide|review this|is this good|is my .*good|"
            r"is this correct|is it correct|is this accurate|is it accurate|any improvement needed|"
            r"improvement needed|needs improvement|can this be improved|can it be improved|"
            r"what do you think|does this work|does this look|look okay|looks okay|is it okay|"
            r"is this okay|is it suitable|is this suitable|"
            r"what should i say|how should i present|present this|talking points?|"
            r"why is|what is|explain this|can you explain|help me understand)\b",
            text,
            re.IGNORECASE,
        )
    )


def fallback_open_conversation_decision(
    user_input: str,
    assistant_state: dict[str, Any],
    *,
    explicit_topic: Optional[str] = None,
    pending_need_topic: bool = False,
) -> Optional[AssistantDecision]:
    decision = pre_route_message(
        user_input,
        explicit_topic=explicit_topic,
        pending_need_topic=pending_need_topic,
    )
    if decision:
        return decision

    text = safe_str(user_input)
    if pending_need_topic:
        topic = extract_conversational_create_topic(text, pending_need_topic=True)
        if topic:
            return AssistantDecision(
                mode="ask",
                intent="create_ppt",
                slots={"topic": topic},
                message=f"Great - how many slides should the presentation on {topic} have?",
                confidence=0.78,
                reason="pending_topic_supplied",
            )

    if is_readonly_conversation_request(text):
        return AssistantDecision(
            mode="answer",
            intent="readonly_advice",
            confidence=0.84,
            reason="read_only_conversation_fallback",
        )

    if is_create_deck_request(text):
        topic = extract_conversational_create_topic(
            text,
            explicit_topic=explicit_topic,
            pending_need_topic=pending_need_topic,
        )
        if topic:
            return AssistantDecision(
                mode="ask",
                intent="create_ppt",
                slots={"topic": topic},
                message=f"Great - how many slides should the presentation on {topic} have?",
                confidence=0.78,
                reason="create_deck_fallback_with_topic",
            )
        return AssistantDecision(
            mode="ask",
            intent="create_ppt",
            slots={"topic": None},
            message=create_topic_followup(text),
            confidence=0.76,
            reason="create_deck_fallback_missing_topic",
        )

    if re.search(
        r"\b(?:present something|have to present|need slides|need something to show|show in college|"
        r"class presentation|college presentation|seminar tomorrow|presentation tomorrow)\b",
        text,
        re.IGNORECASE,
    ):
        return AssistantDecision(
            mode="ask",
            intent="create_ppt",
            slots={"topic": None},
            message=create_topic_followup(text),
            confidence=0.74,
            reason="natural_create_missing_topic",
        )
    return None


def resolve_open_conversation(
    user_input: str,
    assistant_state: dict[str, Any],
    *,
    llm_client: Any = None,
    deployment: Optional[str] = None,
    explicit_topic: Optional[str] = None,
    pending_need_topic: bool = False,
) -> Optional[AssistantDecision]:
    """
    LLM-backed router for open natural language. This is intentionally safe:
    it may ask a missing create-PPT question or provide read-only advice, but
    it does not perform deck mutations.
    """
    text = safe_str(user_input)
    if not text or should_skip_open_llm_router(text):
        return None

    fallback = fallback_open_conversation_decision(
        text,
        assistant_state,
        explicit_topic=explicit_topic,
        pending_need_topic=pending_need_topic,
    )
    if llm_client is None or not deployment:
        return fallback

    prompt = f"""
You are the open-conversation router for a PowerPoint assistant.

User message:
{json.dumps(text)}

Compact assistant state:
{json.dumps(assistant_state or {}, ensure_ascii=True)}

Known deterministic extraction:
{json.dumps({"explicit_topic": explicit_topic, "pending_need_topic": pending_need_topic}, ensure_ascii=True)}

Return ONLY valid JSON:
{{
  "mode": "ask|answer|execute|none",
  "intent": "create_ppt|readonly_advice|general_help|smalltalk|unknown",
  "slots": {{"topic": string_or_null, "slide_count": integer_or_null}},
  "message": string_or_null,
  "confidence": number_between_0_and_1,
  "reason": string
}}

Rules:
- Use this router as the primary decision maker for create-PPT requests and natural conversation.
- If the user wants help creating slides/presentation/deck for class, college, seminar, tomorrow, or presenting something, infer intent=create_ppt.
- If the user says "another PPT", "new presentation", "make a deck for me", or similar while a deck exists, infer intent=create_ppt for a fresh deck.
- Generic context like "college", "class", "seminar", "something to show", "my presentation", "another topic", or "for me" is NOT a topic. Ask for the topic.
- Only extract a topic when the user gives a real subject, usually after words like "on", "about", or clear subject wording.
- If pending_need_topic=true and the user gives a topic idea, extract that topic and ask for slide count.
- If create_ppt has a topic but no slide_count, mode=ask and ask how many slides.
- If create_ppt has no topic, ask exactly for the topic. Do NOT list topic suggestions unless the user explicitly asks for topic ideas.
- For "can you create another PPT for me?" or "I want to create another PPT for me", message should be like "Of course! What topic should the presentation be about?"
- If the user asks for feedback/advice like "Does this look okay?", "any suggestions?", or "is this correct?", mode=answer and intent=readonly_advice.
- Do not choose execute for edits, deletes, additions, design changes, or rebuilds here.
- Use mode=none when the message should continue through the deterministic app router.
- Keep message concise and contextual.
"""
    try:
        resp = llm_client.chat.completions.create(
            model=deployment,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=380,
        )
        parsed = extract_first_json_object(safe_str(resp.choices[0].message.content))
        mode = safe_str(parsed.get("mode"), "none")
        intent = safe_str(parsed.get("intent"), "unknown")
        confidence = max(0.0, min(1.0, float(parsed.get("confidence") or 0.0)))
        if mode not in {"ask", "answer", "execute", "none"} or mode == "none" or confidence < 0.55:
            return fallback
        slots = parsed.get("slots") if isinstance(parsed.get("slots"), dict) else {}
        if is_create_deck_request(text) and intent != "create_ppt" and not is_topic_suggestion_request(text):
            intent = "create_ppt"
            mode = "ask"
        topic = safe_str(slots.get("topic"))
        if topic and is_generic_presentation_context(topic):
            slots["topic"] = None
        if intent == "create_ppt":
            topic = safe_str(slots.get("topic"))
            slide_count = slots.get("slide_count") or _extract_slide_count_from_text(text)
            if slide_count:
                slots["slide_count"] = slide_count
            if not topic:
                mode = "ask"
                slots["topic"] = None
                parsed["message"] = create_topic_followup(text)
            elif not slide_count:
                mode = "ask"
                parsed["message"] = parsed.get("message") or f"Great - how many slides should the presentation on {topic} have?"
        if intent == "readonly_advice":
            mode = "answer"
        return AssistantDecision(
            mode=mode,
            intent=intent,
            slots=slots,
            message=parsed.get("message"),
            confidence=confidence,
            reason=safe_str(parsed.get("reason"), "llm_open_router"),
        )
    except Exception:
        return fallback

def pre_route_message(
    user_input: str,
    *,
    explicit_topic: Optional[str] = None,
    pending_need_topic: bool = False,
) -> Optional[AssistantDecision]:
    """
    Early, pure routing decision. app.py can execute the decision while this
    module remains UI/framework independent.
    """
    if is_readonly_conversation_request(user_input):
        return AssistantDecision(
            mode="answer",
            intent="readonly_advice",
            confidence=0.90,
            reason="read_only_conversation",
        )
    if is_implicit_create_request(user_input):
        topic = extract_conversational_create_topic(
            user_input,
            explicit_topic=explicit_topic,
            pending_need_topic=pending_need_topic,
        )
        if topic:
            return AssistantDecision(
                mode="ask",
                intent="create_ppt",
                slots={"topic": topic},
                message=f"Great - how many slides should the presentation on {topic} have?",
                confidence=0.86,
                reason="implicit_create_with_topic",
            )
        return AssistantDecision(
            mode="ask",
            intent="create_ppt",
            slots={"topic": None},
            message=create_topic_followup(user_input),
            confidence=0.84,
            reason="implicit_create_missing_topic",
        )
    return None
