"""tests/test_db_fallback.py — Unit tests for DB fallback and CRUD operations."""
import os
import pytest
from db.database import _make_engine, init_db
from db.crud import save_presentation, get_presentation, get_history, update_slide
from schemas import Presentation, Slide, AudienceQuestion


def test_invalid_database_url_fallback(monkeypatch):
    """Setting an unreachable DATABASE_URL must fall back to local SQLite cleanly."""
    monkeypatch.setenv("DATABASE_URL", "postgresql://invalid_user:invalid_pass@localhost:59999/nonexistent_db")
    engine, db_mode = _make_engine()
    assert db_mode == "local"
    assert "sqlite" in str(engine.url)


def test_crud_lifecycle():
    """Verify save, get, history, and update slide operations on the database."""
    init_db()

    deck = Presentation(
        title="CRUD Test Deck",
        audience="Developers",
        slides=[
            Slide(
                slide_no=1,
                title="Slide 1 Title",
                key_points=["Key point 1", "Key point 2"],
                speaker_notes="Speaker notes for slide 1.",
                visual_suggestion="Visual idea 1"
            )
        ],
        audience_questions=[
            AudienceQuestion(
                question="What is the testing framework?",
                suggested_answer="pytest is used for automated testing."
            )
        ]
    )

    pres_id = save_presentation(deck, tone="Professional", duration_min=10, input_type="idea")
    assert isinstance(pres_id, int) and pres_id > 0

    # Fetch history
    history = get_history(limit=5)
    assert any(h["id"] == pres_id for h in history)

    # Fetch presentation detail
    fetched = get_presentation(pres_id)
    assert fetched is not None
    assert fetched.title == "CRUD Test Deck"
    assert len(fetched.slides) == 1
    assert fetched.slides[0].title == "Slide 1 Title"

    # Update slide
    updated = update_slide(
        pres_id,
        slide_no=1,
        new_slide={
            "title": "New Slide 1 Title",
            "key_points": ["New key point 1", "New key point 2"],
            "speaker_notes": "Updated speaker notes for slide 1.",
            "visual_suggestion": "New visual idea"
        }
    )
    assert updated is True

    # Re-fetch and verify update
    re_fetched = get_presentation(pres_id)
    assert re_fetched.slides[0].title == "New Slide 1 Title"
