"""app.py — Streamlit UI only. All logic lives in services/pipeline.py.  Owner: P3
Run:  streamlit run app.py
"""
import streamlit as st

from db.database import DB_MODE, init_db

st.set_page_config(page_title="AI Presentation Builder", page_icon="🎯", layout="wide")
init_db()

st.title("🎯 AI Presentation Builder")
st.sidebar.caption("🟢 Shared cloud DB" if DB_MODE == "shared" else "🟡 Local DB (offline mode)")

# TODO P3:
# 1. Sidebar: idea text OR file upload, audience, tone, duration, number of questions
# 2. Button -> services.pipeline.build_deck(...) inside st.spinner, catch GenerationError / ValueError
# 3. Show each slide in st.expander: title, key points, speaker notes, visual suggestion, Regenerate
# 4. Audience Q&A section
# 5. Download button -> utils.exporter.export_pptx(deck)
# 6. History tab -> db.crud.get_history() / get_presentation()
st.info("Skeleton running. UI coming soon.")
