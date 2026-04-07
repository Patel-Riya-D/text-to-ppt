import base64
import copy
import json
from uuid import uuid4

import requests
import streamlit as st

OUTLINE_URL = "http://127.0.0.1:8000/generate-outline"
BUILD_URL   = "http://127.0.0.1:8000/build-ppt"

LAYOUT_OPTIONS = [
    "title_cover","section_index","bullets","two_column","big_stat",
    "timeline","icon_grid","case_study","table","chart",
    "image_text_split","hybrid_insight",
]

st.set_page_config(page_title="AI PPT Chat Builder", page_icon="💬", layout="wide")


# ─── STATE ────────────────────────────────────────────────────────────────

def init_state():
    defaults = {
        "messages": [{
            "role": "assistant",
            "content": (
                "Tell me the presentation topic and I'll draft slide sections first. "
                "You can then rename headings, add or remove slides, and tweak the content "
                "before generating the PPT."
            ),
        }],
        "outline_payload": None,
        "ppt_result":      None,
        "topic":           "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def add_message(role, content):
    st.session_state.messages.append({"role": role, "content": content})


# ─── SLIDE UTILS ──────────────────────────────────────────────────────────

def ensure_editor_id(slide: dict) -> dict:
    slide = dict(slide or {})
    slide.setdefault("_editor_id",  uuid4().hex)
    slide.setdefault("title",       "Untitled Slide")
    slide.setdefault("subtitle",    "")
    slide.setdefault("layout",      "bullets")
    slide.setdefault("icon",        "▸")
    slide.setdefault("content",     [])
    slide.setdefault("style",       {})
    if not isinstance(slide["content"], list): slide["content"] = []
    slide["content"] = [str(x).strip() for x in slide["content"] if str(x).strip()]
    if not isinstance(slide["style"], dict):   slide["style"] = {}
    return slide

def bullets_to_text(items) -> str:
    if not isinstance(items, list): return ""
    return "\n".join(str(x).strip() for x in items if str(x).strip())

def text_to_bullets(value: str) -> list:
    bullets = []
    for line in str(value or "").splitlines():
        c = line.strip().lstrip("-").lstrip("•").strip()
        if c: bullets.append(c)
    return bullets

def summarize_outline(slides: list) -> str:
    if not slides:
        return "I couldn't generate an outline yet. Try another prompt."
    lines = ["I drafted these sections for the deck:"]
    for i, s in enumerate(slides, 1):
        lines.append(f"{i}. {s.get('title', f'Slide {i}')}")
    lines.append("Edit anything below, then generate the PPT when the structure looks right.")
    return "\n".join(lines)


# ─── LAYOUT-AWARE EDITOR CONTENT ──────────────────────────────────────────

def get_editor_content(slide: dict) -> str:
    """Return text for the textarea based on the slide's actual data fields."""
    layout = slide.get("layout","bullets")

    if layout == "two_column":
        lp = slide.get("left_points",[]) or []
        rp = slide.get("right_points",[]) or []
        combined = list(lp) + list(rp)
        if combined: return bullets_to_text(combined)

    elif layout == "timeline":
        steps = slide.get("steps",[]) or []
        if steps:
            return "\n".join(
                f"{s.get('label',''): <20}{s.get('detail','')}"
                for s in steps if isinstance(s, dict)
            )

    elif layout == "icon_grid":
        items = slide.get("grid_items",[]) or []
        if items:
            return "\n".join(
                f"{g.get('title','')}: {g.get('detail','')}"
                for g in items if isinstance(g, dict)
            )

    elif layout == "table":
        cols = slide.get("table_columns",[]) or []
        rows = slide.get("table_rows",[]) or []
        if cols:
            lines = [" | ".join(str(c) for c in cols)]
            for row in rows:
                lines.append(" | ".join(str(v) for v in row))
            return "\n".join(lines)

    elif layout == "chart":
        data = slide.get("chart_data",[]) or []
        if data:
            return "\n".join(
                f"{d.get('label','')} : {d.get('value','')}"
                for d in data if isinstance(d, dict)
            )

    elif layout == "case_study":
        lines = []
        if slide.get("company"):    lines.append(f"Company: {slide['company']}")
        if slide.get("result"):     lines.append(f"Result: {slide['result']}")
        for m in (slide.get("metrics") or []):
            if isinstance(m,dict): lines.append(f"  {m.get('label','')}: {m.get('value','')}")
        lines.extend(slide.get("content",[]) or [])
        return "\n".join(lines)

    elif layout in ("big_stat","hybrid_insight"):
        lines = []
        if slide.get("stat"):       lines.append(f"STAT: {slide['stat']}")
        if slide.get("stat_label"): lines.append(f"LABEL: {slide['stat_label']}")
        lines.extend(slide.get("content",[]) or [])
        return "\n".join(lines)

    elif layout == "section_index":
        secs = slide.get("sections",[]) or slide.get("content",[]) or []
        return bullets_to_text(secs)

    # Default
    return bullets_to_text(slide.get("content",[]))


def save_editor_content(slide: dict, text: str) -> dict:
    """Parse textarea text back into the slide's data fields."""
    layout = slide.get("layout","bullets")
    lines  = text_to_bullets(text)

    if layout == "two_column":
        mid = max(1, len(lines)//2)
        slide["left_points"]  = lines[:mid]
        slide["right_points"] = lines[mid:]
        slide["content"]      = lines

    elif layout == "timeline":
        # Each line: "label   detail" OR "label: detail"
        steps = []
        for line in lines:
            if ":" in line:
                parts = line.split(":",1)
                steps.append({"label":parts[0].strip(),"detail":parts[1].strip()})
            elif len(line) > 25:
                steps.append({"label":line[:20].strip(),"detail":line[20:].strip()})
            else:
                steps.append({"label":line,"detail":""})
        slide["steps"]   = steps
        slide["content"] = lines

    elif layout == "icon_grid":
        items = []
        for line in lines:
            if ":" in line:
                p = line.split(":",1)
                items.append({"icon":"▸","title":p[0].strip(),"detail":p[1].strip()})
            else:
                items.append({"icon":"▸","title":line[:30],"detail":line})
        slide["grid_items"] = items[:4]
        slide["content"]    = lines

    elif layout == "table":
        # First line = headers, rest = rows
        if lines:
            slide["table_columns"] = [c.strip() for c in lines[0].split("|") if c.strip()]
            slide["table_rows"]    = [
                [c.strip() for c in row.split("|")] for row in lines[1:]
            ]
        slide["content"] = lines

    elif layout == "chart":
        data = []
        for line in lines:
            if ":" in line:
                p = line.split(":",1)
                try: data.append({"label":p[0].strip(),"value":int(float(p[1].strip()))})
                except: data.append({"label":p[0].strip(),"value":50})
            else:
                data.append({"label":line,"value":50})
        slide["chart_data"] = data[:5]
        slide["content"]    = lines

    elif layout == "case_study":
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("company:"):
                slide["company"] = line.split(":",1)[1].strip()
            elif ll.startswith("result:"):
                slide["result"] = line.split(":",1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines

    elif layout in ("big_stat","hybrid_insight"):
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("stat:"):
                slide["stat"] = line.split(":",1)[1].strip()
            elif ll.startswith("label:"):
                slide["stat_label"] = line.split(":",1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines

    elif layout == "section_index":
        slide["sections"] = lines
        slide["content"]  = lines

    else:
        slide["content"] = lines

    return slide


# ─── API CALLS ────────────────────────────────────────────────────────────

def request_outline(topic: str, num_slides: int, tone: str) -> dict:
    resp = requests.post(
        OUTLINE_URL,
        data={"topic": topic, "num_slides": num_slides, "tone": tone},
        timeout=(10, 180),
    )
    if resp.status_code != 200: raise RuntimeError(resp.text)
    payload = resp.json()
    payload["slides"] = [
        ensure_editor_id(s) for s in payload.get("slides", []) if isinstance(s, dict)
    ]
    return payload


def sanitize_outline_for_build(payload: dict) -> dict:
    """
    Pass EVERY layout-specific field through to the backend.
    Nothing is dropped.
    """
    safe = {
        "design_system": payload.get("design_system", {}) if isinstance(payload, dict) else {},
        "slides": [],
    }
    for slide in (payload.get("slides",[]) if isinstance(payload,dict) else []):
        if not isinstance(slide, dict): continue
        layout = str(slide.get("layout","bullets")).strip() or "bullets"
        s = {
            "title":    str(slide.get("title","")).strip(),
            "subtitle": str(slide.get("subtitle","")).strip(),
            "layout":   layout,
            "icon":     str(slide.get("icon","▸")).strip() or "▸",
            "content":  text_to_bullets(bullets_to_text(slide.get("content",[]))),
            "style":    slide.get("style",{}) if isinstance(slide.get("style"),dict) else {},
        }

        # ── Pass through every layout-specific field ──────────────────────
        if layout == "two_column":
            s["left_title"]   = slide.get("left_title","Left")
            s["right_title"]  = slide.get("right_title","Right")
            lp = slide.get("left_points",[])  or []
            rp = slide.get("right_points",[]) or []
            if not lp and not rp:
                mid = max(1, len(s["content"])//2)
                lp, rp = s["content"][:mid], s["content"][mid:]
            s["left_points"]  = lp
            s["right_points"] = rp

        elif layout == "big_stat":
            s["stat"]       = slide.get("stat","—")
            s["stat_label"] = slide.get("stat_label","")
            s["stat_source"]= slide.get("stat_source","")

        elif layout == "timeline":
            s["steps"] = slide.get("steps",[]) or []

        elif layout == "icon_grid":
            s["grid_items"] = slide.get("grid_items",[]) or []

        elif layout == "case_study":
            s["company"] = slide.get("company","")
            s["result"]  = slide.get("result","")
            s["metrics"] = slide.get("metrics",[]) or []

        elif layout == "table":
            s["table_columns"] = slide.get("table_columns",[]) or []
            s["table_rows"]    = slide.get("table_rows",[])    or []

        elif layout == "chart":
            s["chart_title"]  = slide.get("chart_title","")
            s["chart_data"]   = slide.get("chart_data",[])  or []
            s["chart_source"] = slide.get("chart_source","")

        elif layout == "image_text_split":
            s["image_caption"] = slide.get("image_caption","")
            s["image_side"]    = slide.get("image_side","right")

        elif layout == "hybrid_insight":
            s["stat"]       = slide.get("stat","—")
            s["stat_label"] = slide.get("stat_label","")
            s["chart_data"] = slide.get("chart_data",[]) or []

        elif layout == "section_index":
            s["sections"] = slide.get("sections", s["content"])

        safe["slides"].append(s)
    return safe


# ─── PREVIEW ──────────────────────────────────────────────────────────────

def preview_slide(slide: dict, index: int):
    st.markdown(f"### {index}. {slide.get('title','')}")
    if slide.get("subtitle"): st.caption(slide["subtitle"])
    layout  = slide.get("layout","bullets")
    content = slide.get("content",[])

    if layout == "bullets":
        for p in content: st.write(f"• {p}")
    elif layout == "two_column":
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**{slide.get('left_title','Left')}**")
            for p in slide.get("left_points", content[:3]): st.write(f"• {p}")
        with c2:
            st.markdown(f"**{slide.get('right_title','Right')}**")
            for p in slide.get("right_points", content[3:]): st.write(f"• {p}")
    elif layout == "big_stat":
        st.metric(label=slide.get("stat_label",""), value=slide.get("stat","—"))
        for p in content: st.write(f"• {p}")
    elif layout == "timeline":
        for step in (slide.get("steps",[]) or []):
            st.write(f"**{step.get('label','')}**: {step.get('detail','')}")
    elif layout == "icon_grid":
        cols = st.columns(2)
        for i, g in enumerate(slide.get("grid_items",[]) or []):
            with cols[i%2]:
                st.write(f"**{g.get('title','')}**: {g.get('detail','')}")
    elif layout == "table":
        cols = slide.get("table_columns",[]); rows = slide.get("table_rows",[])
        if cols and rows:
            try:
                import pandas as pd
                st.dataframe(pd.DataFrame(rows, columns=cols[:len(rows[0])] if rows else cols))
            except: st.write(" | ".join(cols))
        for p in content: st.write(f"• {p}")
    elif layout == "chart":
        for d in (slide.get("chart_data",[]) or []):
            st.write(f"**{d.get('label','')}**: {d.get('value','')}")
    elif layout == "case_study":
        st.write(f"🏢 **{slide.get('company','')}** — {slide.get('result','')}")
        for m in (slide.get("metrics",[]) or []):
            st.write(f"  +{m.get('value','')}% {m.get('label','')}")
        for p in content: st.write(f"• {p}")
    elif layout == "section_index":
        for i, sec in enumerate(slide.get("sections",content), 1):
            st.write(f"{i:02d}. {sec}")
    else:
        for p in content: st.write(f"• {p}")
    st.markdown("---")


# ─── MAIN UI ──────────────────────────────────────────────────────────────

init_state()

st.title("AI PPT Chat Builder")
st.caption("Plan the deck like a conversation, then edit the slide structure before generating.")

# Sidebar
st.sidebar.title("Deck Settings")
num_slides     = st.sidebar.slider("Number of Slides", 3, 12, 6)
tone           = st.sidebar.selectbox("Presentation Style", ["Professional","Creative","Educational"])
uploaded_logo  = st.sidebar.file_uploader("Upload Company Logo (optional)", type=["png","jpg","jpeg"])
if uploaded_logo:
    st.sidebar.image(uploaded_logo, caption="Logo preview", width=140)
if st.sidebar.button("Reset Conversation", use_container_width=True):
    for k in ("messages","outline_payload","ppt_result","topic"):
        st.session_state.pop(k, None)
    st.rerun()

# Chat history
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Chat input
prompt = st.chat_input("Describe the presentation you want to build")
if prompt:
    prompt = prompt.strip()
    if prompt:
        add_message("user", prompt)
        effective = (
            f"{st.session_state.topic}\n\nAdditional refinement: {prompt}"
            if st.session_state.topic else prompt
        )
        if not st.session_state.topic:
            st.session_state.topic = prompt
        with st.spinner("Drafting slide sections..."):
            try:
                op = request_outline(effective, num_slides, tone)
                st.session_state.outline_payload = op
                st.session_state.ppt_result      = None
                add_message("assistant", summarize_outline(op.get("slides",[])))
            except requests.exceptions.Timeout:
                add_message("assistant","The outline request timed out. Try fewer slides.")
            except requests.exceptions.ConnectionError:
                add_message("assistant","Can't reach the backend at `http://127.0.0.1:8000`.")
            except Exception as exc:
                add_message("assistant", f"Outline generation failed: `{exc}`")
        st.rerun()

# ─── OUTLINE EDITOR ───────────────────────────────────────────────────────

outline_payload = st.session_state.outline_payload
if outline_payload:
    st.markdown("## Edit Sections")
    st.caption("Edit headings, content, and layouts. Changes sync back to layout-specific fields automatically.")

    if st.button("Add Section", use_container_width=False):
        n = len(outline_payload.get("slides",[])) + 1
        outline_payload["slides"].append(ensure_editor_id({
            "title":f"New Section {n}","subtitle":"","layout":"bullets",
            "content":["Add your main point here.","Add a supporting point.","Add a takeaway."],
            "style":{},
        }))
        st.session_state.ppt_result = None
        st.rerun()

    slides = outline_payload.get("slides",[])
    for index, slide in enumerate(slides):
        slide    = ensure_editor_id(slide)
        slides[index] = slide
        sid      = slide["_editor_id"]
        layout   = slide.get("layout","bullets")

        # Layout-specific hint
        hints = {
            "two_column":       "One bullet per line. First half → Left column, second half → Right column.",
            "timeline":         "Format: `Phase Label: Detail description` (one per line)",
            "icon_grid":        "Format: `Title: Detail description` (4 lines for 4 grid items)",
            "table":            "First line = column names separated by |. Then data rows with |.",
            "chart":            "Format: `Label : value` (numeric value 1-100, one per line)",
            "case_study":       "Start lines with `Company:` and `Result:` then add bullet points.",
            "big_stat":         "Start with `STAT: value` and `LABEL: description` then add bullet points.",
            "hybrid_insight":   "Start with `STAT: value` and `LABEL: description` then add bullet points.",
            "section_index":    "One section title per line.",
        }
        hint = hints.get(layout, "One bullet point per line. Use **text** for bold.")

        with st.expander(f"Slide {index+1}: {slide.get('title','Untitled Slide')}", expanded=index<2):
            title    = st.text_input("Heading",   value=slide.get("title",""),    key=f"t_{sid}")
            subtitle = st.text_input("Subtitle",  value=slide.get("subtitle",""), key=f"s_{sid}")
            li = LAYOUT_OPTIONS.index(layout) if layout in LAYOUT_OPTIONS else 1
            new_layout = st.selectbox("Layout", LAYOUT_OPTIONS, index=li, key=f"l_{sid}")

            # If layout changed, clear stale layout-specific fields
            if new_layout != layout:
                for fld in ("steps","grid_items","left_points","right_points",
                            "table_columns","table_rows","chart_data","metrics"):
                    slide.pop(fld, None)
                slide["layout"] = new_layout
                layout = new_layout

            editor_val = get_editor_content(slide)
            bullet_text = st.text_area(
                "Content or talking points",
                value=editor_val,
                height=200,
                help=hint,
                key=f"c_{sid}",
            )
            st.caption(f"💡 {hint}")

            # Save updated values
            slide["title"]    = title.strip() or f"Slide {index+1}"
            slide["subtitle"] = subtitle.strip()
            slide["layout"]   = layout
            slide = save_editor_content(slide, bullet_text)
            slides[index] = slide

            uc, dc, rc = st.columns(3)
            if uc.button("Move Up",   key=f"u_{sid}", disabled=index==0,              use_container_width=True):
                slides[index-1], slides[index] = slides[index], slides[index-1]
                st.session_state.ppt_result = None; st.rerun()
            if dc.button("Move Down", key=f"d_{sid}", disabled=index==len(slides)-1,  use_container_width=True):
                slides[index+1], slides[index] = slides[index], slides[index+1]
                st.session_state.ppt_result = None; st.rerun()
            if rc.button("Remove",    key=f"r_{sid}",                                 use_container_width=True):
                slides.pop(index)
                st.session_state.ppt_result = None; st.rerun()

    # Build / Regenerate buttons
    bc, rc2 = st.columns(2)
    build_clicked = bc.button("Generate PPT", type="primary", use_container_width=True)
    regen_clicked = rc2.button("Regenerate Outline", use_container_width=True)

    if regen_clicked:
        if not st.session_state.topic.strip():
            st.warning("Enter a presentation request in chat first.")
        else:
            with st.spinner("Refreshing outline..."):
                try:
                    refreshed = request_outline(st.session_state.topic, num_slides, tone)
                    st.session_state.outline_payload = refreshed
                    st.session_state.ppt_result      = None
                    add_message("assistant", summarize_outline(refreshed.get("slides",[])))
                    st.rerun()
                except Exception as exc:
                    st.error(f"Outline refresh failed: {exc}")

    if build_clicked:
        build_payload = sanitize_outline_for_build(copy.deepcopy(outline_payload))
        files = None
        if uploaded_logo:
            files = {"logo":(uploaded_logo.name, uploaded_logo.getvalue(), uploaded_logo.type)}
        with st.spinner("Building your presentation..."):
            try:
                resp = requests.post(
                    BUILD_URL,
                    data={
                        "topic":       st.session_state.topic or "Presentation",
                        "tone":        tone,
                        "slides_json": json.dumps(build_payload),
                    },
                    files=files,
                    timeout=(10, 180),
                )
                if resp.status_code != 200:
                    st.error(resp.text)
                else:
                    st.session_state.ppt_result = resp.json()
                    add_message("assistant","The PPT is ready — preview and download below.")
                    st.rerun()
            except requests.exceptions.Timeout:
                st.error("PPT generation timed out. Try reducing the slide count.")
            except requests.exceptions.ConnectionError:
                st.error("Cannot connect to the backend. Ensure FastAPI is running on `http://127.0.0.1:8000`.")
            except Exception as exc:
                st.error(f"Error: {exc}")

# ─── PREVIEW + DOWNLOAD ───────────────────────────────────────────────────

if st.session_state.ppt_result:
    result = st.session_state.ppt_result
    st.markdown("## Slide Preview")
    for i, slide in enumerate(result.get("slides",[]), 1):
        if isinstance(slide, dict):
            preview_slide(slide, i)

    ppt_bytes  = base64.b64decode(result["ppt_base64"])
    file_topic = (st.session_state.topic or "presentation").strip().replace(" ","_")
    st.download_button(
        label="⬇️ Download PPT",
        data=ppt_bytes,
        file_name=f"{file_topic}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )