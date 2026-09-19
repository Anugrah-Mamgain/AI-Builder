"""services/pipeline.py — The one function the UI calls. No Streamlit code here.  Owner: P3
Flow: parse -> redact -> generate -> save
"""
from typing import Optional, Tuple

from db.crud import save_presentation
from llm.generator import generate_presentation
from schemas import Presentation
from utils.logger import get_logger
from utils.parser import extract_text
from utils.pii import redact_pii

log = get_logger(__name__)


def build_deck(idea_text: str = "", uploaded_file=None, audience: str = "General",
               tone: str = "Professional", duration_min: int = 10,
               num_questions: int = 5) -> Tuple[Presentation, Optional[int]]:
    """Returns (deck, presentation_id). presentation_id is None if saving failed."""
    if uploaded_file is not None:
        raw, input_type, filename = extract_text(uploaded_file), "document", uploaded_file.name
    else:
        raw, input_type, filename = idea_text, "idea", None

    clean = redact_pii(raw)
    deck = generate_presentation(clean, audience, tone, duration_min, num_questions, input_type)

    try:
        pres_id = save_presentation(deck, tone, duration_min, input_type, filename,
                                    clean if input_type == "document" else None)
    except Exception as e:
        # A DB failure should not lose the user's generated deck
        log.error(f"Save failed: {type(e).__name__}")
        pres_id = None
    return deck, pres_id
