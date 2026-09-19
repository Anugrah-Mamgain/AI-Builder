"""db/crud.py — Database operations.
Input/output uses the Pydantic models from schemas.py so the rest of the app never touches SQL.
"""
import hashlib
from typing import List, Optional, Union

from db.database import SessionLocal
from db.models import PresentationRow, QuestionRow, SlideRow, SourceDocumentRow
from schemas import AudienceQuestion, Presentation, Slide


def save_presentation(deck: Presentation, tone: str = "", duration_min: int = 0,
                      input_type: str = "idea", filename: Optional[str] = None,
                      redacted_text: Optional[str] = None) -> int:
    """Save deck + slides + questions (+ source doc if uploaded). Return presentation id."""
    session = SessionLocal()
    try:
        pres_row = PresentationRow(
            title=deck.title,
            audience=deck.audience,
            tone=tone,
            duration_min=duration_min,
            input_type=input_type,
        )
        session.add(pres_row)
        session.flush()  # obtains pres_row.id

        for slide in deck.slides:
            slide_row = SlideRow(
                presentation_id=pres_row.id,
                slide_no=slide.slide_no,
                title=slide.title,
                key_points=slide.key_points,
                speaker_notes=slide.speaker_notes,
                visual_suggestion=slide.visual_suggestion,
            )
            session.add(slide_row)

        for q in deck.audience_questions:
            q_row = QuestionRow(
                presentation_id=pres_row.id,
                question=q.question,
                suggested_answer=q.suggested_answer,
            )
            session.add(q_row)

        if filename or redacted_text:
            content_hash = None
            if redacted_text:
                content_hash = hashlib.sha256(redacted_text.encode('utf-8')).hexdigest()
            doc_row = SourceDocumentRow(
                presentation_id=pres_row.id,
                filename=filename,
                content_hash=content_hash,
                redacted_text=redacted_text,
            )
            session.add(doc_row)

        session.commit()
        return pres_row.id
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def get_history(limit: int = 20) -> List[dict]:
    """Return [{'id', 'title', 'audience', 'created_at', 'slide_count', 'input_type'}], newest first."""
    session = SessionLocal()
    try:
        rows = (
            session.query(PresentationRow)
            .order_by(PresentationRow.created_at.desc())
            .limit(limit)
            .all()
        )
        history = []
        for r in rows:
            history.append({
                "id": r.id,
                "title": r.title,
                "audience": r.audience,
                "tone": r.tone,
                "duration_min": r.duration_min,
                "input_type": r.input_type,
                "created_at": r.created_at.isoformat() if r.created_at else None,
                "slide_count": len(r.slides),
            })
        return history
    finally:
        session.close()


def get_presentation(presentation_id: int) -> Optional[Presentation]:
    """Load a saved deck back as a Presentation object (None if not found)."""
    session = SessionLocal()
    try:
        row = session.query(PresentationRow).filter(PresentationRow.id == presentation_id).first()
        if not row:
            return None

        slides = [
            Slide(
                slide_no=s.slide_no,
                title=s.title,
                key_points=s.key_points if isinstance(s.key_points, list) else list(s.key_points),
                speaker_notes=s.speaker_notes or "",
                visual_suggestion=s.visual_suggestion or "",
            )
            for s in row.slides
        ]

        questions = [
            AudienceQuestion(
                question=q.question,
                suggested_answer=q.suggested_answer or "",
            )
            for q in row.questions
        ]

        return Presentation(
            title=row.title,
            audience=row.audience or "General",
            slides=slides,
            audience_questions=questions,
        )
    finally:
        session.close()


def update_slide(presentation_id: int, slide_no: int, new_slide: Union[dict, Slide]) -> bool:
    """Replace one slide after regeneration. Returns True if updated, False if slide not found."""
    if isinstance(new_slide, Slide):
        new_slide_dict = new_slide.model_dump()
    elif isinstance(new_slide, dict):
        new_slide_dict = new_slide
    else:
        raise ValueError("new_slide must be a Slide object or dict")

    session = SessionLocal()
    try:
        slide_row = (
            session.query(SlideRow)
            .filter(SlideRow.presentation_id == presentation_id, SlideRow.slide_no == slide_no)
            .first()
        )
        if not slide_row:
            return False

        if "title" in new_slide_dict:
            slide_row.title = new_slide_dict["title"]
        if "key_points" in new_slide_dict:
            slide_row.key_points = new_slide_dict["key_points"]
        if "speaker_notes" in new_slide_dict:
            slide_row.speaker_notes = new_slide_dict["speaker_notes"]
        if "visual_suggestion" in new_slide_dict:
            slide_row.visual_suggestion = new_slide_dict["visual_suggestion"]

        session.commit()
        return True
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
