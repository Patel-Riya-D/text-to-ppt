import streamlit as st
import requests
import base64

API_URL = "http://127.0.0.1:8000/generate-ppt"

# =========================
# Page Config
# =========================
st.set_page_config(
    page_title="AI PPT Generator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI PPT Generator")
st.caption("Create professional presentations with your company branding")

# =========================
# Sidebar Controls
# =========================
st.sidebar.title("⚙️ Settings")

num_slides = st.sidebar.slider("Number of Slides", 3, 10, 5)

tone = st.sidebar.selectbox(
    "Presentation Style",
    ["Professional", "Creative", "Educational"]
)

# =========================
# Input Section
# =========================
topic = st.text_input("Enter Topic", placeholder="e.g. Artificial Intelligence")

uploaded_logo = st.file_uploader(
    "🏢 Upload Company Logo (optional)",
    type=["png", "jpg", "jpeg"]
)
if uploaded_logo:
    st.image(uploaded_logo, caption="Your Company Logo", width=150)

# =========================
# Generate Button
# =========================
if st.button("🚀 Generate PPT"):
    if not topic.strip():
        st.warning("Please enter a topic")
    else:
        with st.spinner("Generating slides... ⏳"):
            try:
                files = None
                if uploaded_logo:
                    files = {
                        "logo": (
                            uploaded_logo.name,
                            uploaded_logo.getvalue(),
                            uploaded_logo.type
                        )
                    }

                response = requests.post(
                    API_URL,
                    data={
                        "topic": topic,
                        "num_slides": num_slides,
                        "tone": tone
                    },
                    files=files,
                    timeout=(10, 180),
                )

                if response.status_code != 200:
                    st.error(response.text)
                else:
                    data = response.json()
                    st.success("PPT Generated Successfully!")

                    # =========================
                    # Slide Preview
                    # =========================
                    st.subheader("📊 Slide Preview")

                    for i, slide in enumerate(data.get("slides", []), 1):
                        with st.container():
                            st.markdown(f"### {i}. {slide.get('title', '')}")

                            if slide.get("subtitle"):
                                st.caption(slide["subtitle"])

                            layout = slide.get("layout", "bullets")

                            if layout == "bullets" and slide.get("content"):
                                for point in slide["content"]:
                                    st.write(f"• {point}")

                            elif layout == "two_column":
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown(f"**{slide.get('left_title', 'Left')}**")
                                    for p in slide.get("left_points", slide.get("content", [])[:3]):
                                        st.write(f"• {p}")
                                with col2:
                                    st.markdown(f"**{slide.get('right_title', 'Right')}**")
                                    for p in slide.get("right_points", slide.get("content", [])[3:]):
                                        st.write(f"• {p}")

                            elif layout == "big_stat":
                                st.metric(
                                    label=slide.get("stat_label", ""),
                                    value=slide.get("stat", "—")
                                )
                                for point in slide.get("content", []):
                                    st.write(f"• {point}")

                            elif layout == "timeline":
                                for step in slide.get("steps", []):
                                    label  = step.get("label", "")
                                    detail = step.get("detail", "")
                                    st.write(f"**{label}**: {detail}")

                            elif layout == "icon_grid":
                                cols = st.columns(2)
                                for idx, gi in enumerate(slide.get("grid_items", [])):
                                    with cols[idx % 2]:
                                        st.write(f"**{gi.get('title', '')}**: {gi.get('detail', '')}")

                            elif layout == "table":
                                cols = slide.get("table_columns", [])
                                rows = slide.get("table_rows", [])
                                if cols and rows:
                                    try:
                                        import pandas as pd
                                        normalized_rows = []
                                        for r in rows:
                                            if not isinstance(r, list):
                                                continue
                                            rr = [str(x) for x in r[:len(cols)]]
                                            while len(rr) < len(cols):
                                                rr.append("")
                                            normalized_rows.append(rr)
                                        if normalized_rows:
                                            st.dataframe(pd.DataFrame(normalized_rows, columns=cols), width="stretch")
                                    except Exception:
                                        # Fallback without pandas
                                        st.write(" | ".join([f"**{c}**" for c in cols]))
                                        for r in rows:
                                            if isinstance(r, list):
                                                st.write(" | ".join([str(x) for x in r[:len(cols)]]))
                                for point in slide.get("content", []):
                                    st.write(f"• {point}")

                            else:
                                # Fallback: show content bullets if available
                                for point in slide.get("content", []):
                                    st.write(f"• {point}")

                            st.markdown("---")

                    # =========================
                    # Download
                    # =========================
                    ppt_bytes = base64.b64decode(data["ppt_base64"])
                    st.download_button(
                        label="⬇️ Download PPT",
                        data=ppt_bytes,
                        file_name=f"{topic}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except requests.exceptions.Timeout:
                st.error("Generation timed out. Try reducing slides (e.g., 3-5) and retry.")
            except requests.exceptions.ConnectionError:
                st.error("Cannot connect to backend. Ensure FastAPI server is running on http://127.0.0.1:8000.")
            except Exception as e:
                st.error(f"Error: {e}")
