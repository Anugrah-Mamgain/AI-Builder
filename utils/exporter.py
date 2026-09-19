"""utils/exporter.py — Build a .pptx from a Presentation.  Owner: P2
Use python-pptx. Put speaker_notes in slide.notes_slide.notes_text_frame.
Add visual_suggestion as a small italic line or in the notes. Return bytes for st.download_button.
"""
from schemas import Presentation


def export_pptx(deck: Presentation) -> bytes:
    raise NotImplementedError("P2: implement export_pptx")
