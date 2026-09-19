"""db/crud.py — Database operations. Owner: P2.

Input/output uses Pydantic models from schemas.py so the rest of the app
does not need to interact with SQL directly.

Security:
- Only redacted_text may be persisted.
- content_hash must represent the ORIGINAL uploaded file bytes.
- Raw source text must never be stored.
"""

from typing import List, Optional, Union

from db.database import SessionLocal
from db.models import (
    PresentationRow,
    QuestionRow,
    SlideRow,
    SourceDocumentRow,
)
from schemas import AudienceQuestion, Presentation, Slide


def save_presentation(
    deck: Presentation,
    tone: str = "",
    duration_min: int = 0,
    input_type: str = "idea",
    filename: Optional[str] = None,
    redacted_text: Optional[str] = None,
    content_hash: Optional[str] = None,
) -> int:
    """Save presentation, slides, questions and optional sanitized source.

    Returns:
        Database presentation ID.

    Raises:
        ValueError: if source-document metadata is incomplete or unsafe.
    """

    # A source document must never be stored without a content hash.
    if filename is not None and not content_hash:
        raise ValueError(
            "content_hash is required when saving a source document."
        )

    # We only persist sanitized/redacted text.
    if redacted_text is not None and not isinstance(redacted_text, str):
        raise TypeError("redacted_text must be a string or None.")

    if content_hash is not None:
        if len(content_hash) != 64:
            raise ValueError("content_hash must be a SHA-256 hex digest.")
        try:
            int(content_hash, 16)
        except ValueError as exc:
            raise ValueError(
                "content_hash must contain only hexadecimal characters."
            ) from exc

    session = SessionLocal()

    try:
        presentation_row = PresentationRow(
            title=deck.title,
            audience=deck.audience,
            tone=tone,
            duration_min=duration_min,
            input_type=input_type,
        )

        session.add(presentation_row)
        session.flush()  # Gets presentation_row.id.

        for slide in deck.slides:
            slide_row = SlideRow(
                presentation_id=presentation_row.id,
                slide_no=slide.slide_no,
                title=slide.title,
                key_points=slide.key_points,
                speaker_notes=slide.speaker_notes,
                visual_suggestion=slide.visual_suggestion,
            )

            session.add(slide_row)

        for question in deck.audience_questions:
            question_row = QuestionRow(
                presentation_id=presentation_row.id,
                question=question.question,
                suggested_answer=question.suggested_answer,
            )

            session.add(question_row)

        # Only sanitized text is persisted.
        if filename is not None or redacted_text is not None:
            document_row = SourceDocumentRow(
                presentation_id=presentation_row.id,
                filename=filename,
                content_hash=content_hash,
                redacted_text=redacted_text,
            )

            session.add(document_row)

        session.commit()

        return presentation_row.id

    except Exception:
        session.rollback()
        raise

    finally:
        session.close()


def get_history(limit: int = 20) -> List[dict]:
    """Return recent presentation metadata, newest first."""

    if limit <= 0:
        return []

    session = SessionLocal()

    try:
        rows = (
            session.query(PresentationRow)
            .order_by(PresentationRow.created_at.desc())
            .limit(limit)
            .all()
        )

        history = []

        for row in rows:
            history.append(
                {
                    "id": row.id,
                    "title": row.title,
                    "audience": row.audience,
                    "tone": row.tone,
                    "duration_min": row.duration_min,
                    "input_type": row.input_type,
                    "created_at": (
                        row.created_at.isoformat()
                        if row.created_at
                        else None
                    ),
                    "slide_count": len(row.slides),
                }
            )

        return history

    finally:
        session.close()


def get_presentation(
    presentation_id: int,
) -> Optional[Presentation]:
    """Load a saved presentation as a Pydantic Presentation model."""

    session = SessionLocal()

    try:
        row = (
            session.query(PresentationRow)
            .filter(PresentationRow.id == presentation_id)
            .first()
        )

        if row is None:
            return None

        slides = [
            Slide(
                slide_no=slide.slide_no,
                title=slide.title,
                key_points=(
                    slide.key_points
                    if isinstance(slide.key_points, list)
                    else list(slide.key_points)
                ),
                speaker_notes=slide.speaker_notes or "",
                visual_suggestion=slide.visual_suggestion or "",
            )
            for slide in row.slides
        ]

        questions = [
            AudienceQuestion(
                question=question.question,
                suggested_answer=question.suggested_answer or "",
            )
            for question in row.questions
        ]

        return Presentation(
            title=row.title,
            audience=row.audience or "General",
            slides=slides,
            audience_questions=questions,
        )

    finally:
        session.close()


def update_slide(
    presentation_id: int,
    slide_no: int,
    new_slide: Union[dict, Slide],
) -> bool:
    """Replace one slide after regeneration.

    Returns:
        True if the slide was found and updated, otherwise False.
    """

    if isinstance(new_slide, Slide):
        slide_data = new_slide.model_dump()
    elif isinstance(new_slide, dict):
        slide_data = new_slide
    else:
        raise ValueError("new_slide must be a Slide or dict.")

    session = SessionLocal()

    try:
        slide_row = (
            session.query(SlideRow)
            .filter(
                SlideRow.presentation_id == presentation_id,
                SlideRow.slide_no == slide_no,
            )
            .first()
        )

        if slide_row is None:
            return False

        if "title" in slide_data:
            slide_row.title = slide_data["title"]

        if "key_points" in slide_data:
            slide_row.key_points = slide_data["key_points"]

        if "speaker_notes" in slide_data:
            slide_row.speaker_notes = slide_data["speaker_notes"]

        if "visual_suggestion" in slide_data:
            slide_row.visual_suggestion = slide_data["visual_suggestion"]

        session.commit()

        return True

    except Exception:
        session.rollback()
        raise

    finally:
        session.close()