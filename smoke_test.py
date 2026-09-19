"""Run the full P3 pipeline without Streamlit: python smoke_test.py."""
from services.pipeline import build_deck, export_deck


def main() -> None:
    deck, demo = build_deck(
        source_type="idea", idea="AI adoption roadmap for Indian hospitals",
        uploaded_bytes=None, uploaded_name=None, audience="Executives",
        tone="Professional", duration_minutes=10, slide_count=6, qa_count=3,
    )
    pptx_bytes, export_demo = export_deck(deck)
    slides = deck["slides"] if isinstance(deck, dict) else deck.slides
    assert len(slides) == 6
    assert len(pptx_bytes) > 1_000
    print("Smoke test passed:", {"slides": len(slides), "pptx_bytes": len(pptx_bytes),
                                  "demo_mode": demo or export_demo})


if __name__ == "__main__":
    main()
