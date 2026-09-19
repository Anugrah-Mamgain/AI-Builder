"""db/crud.py — Database operations.  Owner: P2
Input/output uses the Pydantic models from schemas.py so the rest of the app never touches SQL.
"""
from typing import List, Optional

from schemas import Presentation


def save_presentation(deck: Presentation, tone: str = "", duration_min: int = 0,
                      input_type: str = "idea", filename: Optional[str] = None,
                      redacted_text: Optional[str] = None) -> int:
    """Save deck + slides + questions (+ source doc if uploaded). Return presentation id."""
    raise NotImplementedError("P2: implement save_presentation")


def get_history(limit: int = 20) -> List[dict]:
    """Return [{'id', 'title', 'audience', 'created_at'}], newest first."""
    raise NotImplementedError("P2: implement get_history")


def get_presentation(presentation_id: int) -> Optional[Presentation]:
    """Load a saved deck back as a Presentation object (None if not found)."""
    raise NotImplementedError("P2: implement get_presentation")


def update_slide(presentation_id: int, slide_no: int, new_slide: dict) -> None:
    """Replace one slide after regeneration."""
    raise NotImplementedError("P2: implement update_slide")
