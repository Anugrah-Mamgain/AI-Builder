"""Run from project root: python -m scripts.check_db"""
from sqlalchemy import inspect, text

from db.crud import get_history, get_presentation, save_presentation, update_slide
from db.database import DB_MODE, engine, init_db
from schemas import AudienceQuestion, Presentation, Slide


def main():
    print("--- Checking DB Connection & Initialization ---")
    init_db()
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print(f"DB mode : {DB_MODE}")
    tables = sorted(inspect(engine).get_table_names())
    print(f"Tables  : {tables}")

    print("\n--- Testing CRUD Lifecycle ---")
    sample_deck = Presentation(
        title="Test Presentation",
        audience="Engineers",
        slides=[
            Slide(
                slide_no=1,
                title="Introduction to AI",
                key_points=["Point 1", "Point 2"],
                speaker_notes="Welcome everyone to this test presentation.",
                visual_suggestion="Diagram showing AI architecture"
            ),
            Slide(
                slide_no=2,
                title="Deep Dive into Data Layer",
                key_points=["SQLAlchemy models", "Supabase PostgreSQL", "SQLite fallback"],
                speaker_notes="Here we explain how portable JSON column types work.",
                visual_suggestion="Flowchart showing fallback mechanism"
            )
        ],
        audience_questions=[
            AudienceQuestion(
                question="Does SQLite fallback support all fields?",
                suggested_answer="Yes, portable JSON and text types work identically on SQLite and Postgres."
            )
        ]
    )

    pres_id = save_presentation(
        deck=sample_deck,
        tone="Informative",
        duration_min=15,
        input_type="idea"
    )
    print(f"[OK] Created test presentation ID: {pres_id}")

    fetched_deck = get_presentation(pres_id)
    assert fetched_deck is not None, "Failed to retrieve presentation"
    print(f"[OK] Retrieved presentation: '{fetched_deck.title}' with {len(fetched_deck.slides)} slides")

    history = get_history()
    print(f"[OK] History query returned {len(history)} record(s)")

    # Test updating slide
    updated = update_slide(
        presentation_id=pres_id,
        slide_no=2,
        new_slide={
            "title": "Updated Deep Dive Title",
            "key_points": ["Updated Point 1", "Updated Point 2"],
            "speaker_notes": "Updated speaker notes for slide 2.",
            "visual_suggestion": "Updated visual suggestion."
        }
    )
    print(f"[OK] Slide update status: {updated}")

    updated_deck = get_presentation(pres_id)
    assert updated_deck.slides[1].title == "Updated Deep Dive Title"
    print(f"[OK] Verified updated slide title: '{updated_deck.slides[1].title}'")

    print("\n[SUCCESS] DB & CRUD Verification completed successfully!")


if __name__ == "__main__":
    main()
