"""Streamlit UI for the AI Presentation Builder (P3 ownership)."""
from __future__ import annotations

from typing import Any
import streamlit as st

from services.pipeline import (
    build_deck, export_deck, get_history_items, load_history_deck,
    regenerate_deck_slide,
)

st.set_page_config(page_title="PitchPilot AI", page_icon="✨", layout="wide")


def apply_styles() -> None:
    st.markdown("""
    <style>
    .block-container {padding-top:1.35rem; padding-bottom:3rem; max-width:1280px}
    [data-testid="stSidebar"] {border-right:1px solid rgba(128,128,128,.18)}
    .hero {padding:1.3rem 1.5rem; border-radius:18px; margin-bottom:1.1rem;
      color:#fff; background:linear-gradient(120deg,#4338ca,#7c3aed 55%,#db2777)}
    .hero h1 {color:#fff; margin:0; font-size:2.05rem}
    .hero p {margin:.35rem 0 0; opacity:.9}
    </style>""", unsafe_allow_html=True)


def value(obj: Any, field: str, default: Any = None) -> Any:
    return obj.get(field, default) if isinstance(obj, dict) else getattr(obj, field, default)


def set_value(obj: Any, field: str, new_value: Any) -> None:
    if isinstance(obj, dict):
        obj[field] = new_value
    else:
        setattr(obj, field, new_value)


def sidebar_inputs() -> tuple[dict[str, Any], bool]:
    with st.sidebar:
        st.title("✨ PitchPilot")
        st.caption("AI-powered presentation builder")
        st.divider()
        mode = st.radio("Start from", ["An idea", "A document"], horizontal=True)
        idea, uploaded_bytes, uploaded_name = "", None, None
        if mode == "An idea":
            idea = st.text_area(
                "Describe your presentation",
                placeholder="Example: A 90-day AI adoption plan for Indian hospitals",
                height=135,
            )
            source_ready = bool(idea.strip())
        else:
            uploaded = st.file_uploader("Upload source material", type=["pdf", "docx", "txt"])
            source_ready = uploaded is not None
            if uploaded is not None:
                uploaded_bytes, uploaded_name = uploaded.getvalue(), uploaded.name
                st.success(f"Ready: {uploaded.name}")

        st.subheader("Tailor the deck")
        audience = st.selectbox(
            "Audience", ["Executives", "Technical team", "Clients", "Students", "General"]
        )
        tone = st.selectbox(
            "Tone", ["Professional", "Persuasive", "Educational", "Inspirational", "Concise"]
        )
        duration = st.slider("Duration (minutes)", 3, 30, 10)
        slide_count = st.slider("Number of slides", 3, 15, 7)
        qa_count = st.slider("Q&A items", 1, 8, 3)
        clicked = st.button(
            "Generate presentation", type="primary", use_container_width=True,
            disabled=not source_ready,
        )
        st.caption("Sensitive data is redacted before content is sent to the LLM.")

    return ({
        "source_type": "idea" if mode == "An idea" else "document",
        "idea": idea.strip(), "uploaded_bytes": uploaded_bytes,
        "uploaded_name": uploaded_name, "audience": audience, "tone": tone,
        "duration_minutes": duration, "slide_count": slide_count, "qa_count": qa_count,
    }, clicked)


def render_slide(slide: Any, index: int, deck: Any, settings: dict[str, Any]) -> None:
    version = st.session_state.deck_version
    title = value(slide, "title", f"Slide {index + 1}")
    with st.expander(f"{index + 1:02d} · {title}", expanded=index == 0):
        edited_title = st.text_input("Slide title", title, key=f"title_{version}_{index}")
        edited_bullets = st.text_area(
            "Key points — one per line", "\n".join(value(slide, "bullets", [])),
            height=125, key=f"bullets_{version}_{index}",
        )
        notes_col, visual_col = st.columns(2)
        with notes_col:
            edited_notes = st.text_area(
                "Speaker notes", value(slide, "speaker_notes", ""), height=115,
                key=f"notes_{version}_{index}",
            )
        with visual_col:
            edited_visual = st.text_area(
                "Visual suggestion", value(slide, "visual_suggestion", ""), height=115,
                key=f"visual_{version}_{index}",
            )
        set_value(slide, "title", edited_title.strip() or f"Slide {index + 1}")
        set_value(slide, "bullets", [x.strip() for x in edited_bullets.splitlines() if x.strip()])
        set_value(slide, "speaker_notes", edited_notes.strip())
        set_value(slide, "visual_suggestion", edited_visual.strip())

        if st.button("↻ Regenerate this slide", key=f"regen_{index}"):
            try:
                with st.spinner(f"Improving slide {index + 1}…"):
                    updated, demo = regenerate_deck_slide(
                        deck=deck, slide_index=index, settings=settings
                    )
                st.session_state.deck = updated
                st.session_state.deck_version += 1
                st.toast("Slide regenerated" + (" in demo mode" if demo else ""), icon="✅")
                st.rerun()
            except Exception as exc:
                st.error(f"This slide couldn't be regenerated. Details: {exc}")


def render_builder(deck: Any, settings: dict[str, Any]) -> None:
    heading, metric = st.columns([4, 1])
    with heading:
        st.subheader(value(deck, "title", "Generated presentation"))
        if value(deck, "subtitle", ""):
            st.caption(value(deck, "subtitle"))
    slides = value(deck, "slides", [])
    with metric:
        st.metric("Slides", len(slides))
    st.caption("Review, edit, or regenerate only the slide you want to improve.")
    for index, slide in enumerate(slides):
        render_slide(slide, index, deck, settings)

    st.subheader("Likely audience questions")
    for index, item in enumerate(value(deck, "qa", []), start=1):
        with st.expander(f"Q{index}. {value(item, 'question', 'Question')}"):
            st.write(value(item, "answer", "No suggested answer available."))

    st.divider()
    copy_col, download_col = st.columns([2, 1])
    with copy_col:
        st.markdown("**Your presentation is ready.** Download it or refine individual slides.")
    with download_col:
        try:
            pptx_bytes, demo = export_deck(deck)
            st.download_button(
                "Download .pptx", data=pptx_bytes, file_name="presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary", use_container_width=True,
            )
            if demo:
                st.caption("Using the built-in demo template.")
        except Exception as exc:
            st.error(f"The PowerPoint file couldn't be created: {exc}")


def render_history() -> None:
    st.subheader("Presentation history")
    st.caption("Open a previous deck and continue editing it.")
    try:
        records = get_history_items()
    except Exception as exc:
        st.error(f"History is temporarily unavailable: {exc}")
        return
    if not records:
        st.info("Your generated presentations will appear here.")
        return
    for position, record in enumerate(records):
        record_id = value(record, "id", position)
        with st.container(border=True):
            info, action = st.columns([4, 1])
            with info:
                st.markdown(f"**{value(record, 'title', 'Untitled presentation')}**")
                details = " · ".join(str(x) for x in [
                    value(record, "audience", ""), value(record, "tone", ""),
                    value(record, "created_at", ""),
                ] if x)
                st.caption(details)
            with action:
                if st.button("Open", key=f"history_{record_id}", use_container_width=True):
                    try:
                        with st.spinner("Loading presentation…"):
                            st.session_state.deck = load_history_deck(record_id)
                        st.session_state.deck_version += 1
                        st.toast("Loaded. Open the Builder tab.", icon="✅")
                    except Exception as exc:
                        st.error(f"This presentation couldn't be loaded: {exc}")


def main() -> None:
    apply_styles()
    st.session_state.setdefault("deck", None)
    st.session_state.setdefault("deck_version", 0)
    st.session_state.setdefault("last_settings", {})
    settings, generate_clicked = sidebar_inputs()
    st.markdown("""
    <div class="hero"><h1>Ideas in. Presentation out.</h1>
    <p>Build an audience-ready deck with key points, speaker notes, visuals, and Q&amp;A.</p></div>
    """, unsafe_allow_html=True)

    if generate_clicked:
        try:
            with st.spinner("Planning the narrative and writing your slides…"):
                deck, demo = build_deck(**settings)
            st.session_state.deck, st.session_state.last_settings = deck, settings
            st.session_state.deck_version += 1
            if demo:
                st.warning("Demo mode is active until the LLM/database modules are merged.")
            else:
                st.success("Presentation generated and saved.")
        except Exception as exc:
            st.error(f"I couldn't generate the presentation. Check the configuration and retry. Details: {exc}")

    builder_tab, history_tab = st.tabs(["Builder", "History"])
    with builder_tab:
        if st.session_state.deck is None:
            st.info("Choose an idea or upload a document in the sidebar to begin.")
            columns = st.columns(3)
            columns[0].markdown("**Audience-aware**\n\nContent adapts to who is listening.")
            columns[1].markdown("**Speaker-ready**\n\nEvery slide includes useful notes.")
            columns[2].markdown("**Exportable**\n\nDownload an editable PowerPoint.")
        else:
            render_builder(st.session_state.deck, st.session_state.last_settings or settings)
    with history_tab:
        render_history()


if __name__ == "__main__":
    main()
