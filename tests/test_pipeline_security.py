from services.pipeline import build_deck


def test_document_is_redacted_before_llm(monkeypatch):
    captured = {}

    def fake_generate_presentation(
        text,
        audience="General",
        tone="Professional",
        duration_min=10,
        num_questions=5,
        input_type="idea",
    ):
        captured["text"] = text

        return {
            "title": "Security Test",
            "subtitle": "Test",
            "slides": [
                {
                    "title": "Test",
                    "bullets": ["Sanitized input received"],
                    "speaker_notes": "Test",
                    "visual_suggestion": "None",
                }
            ],
            "qa": [],
        }

    monkeypatch.setattr(
        "llm.generator.generate_presentation",
        fake_generate_presentation,
    )

    # Prevent DB from being involved in this security test.
    monkeypatch.setattr(
    "services.pipeline._save",
    lambda deck, settings, **kwargs: False,
)

    source = (
        "Confidential report\n"
        "Email: secret@example.com\n"
        "Phone: +91 9876543210\n"
        "PAN: ABCDE1234F\n"
        "Aadhaar: 1234 5678 9012\n"
        "Card: 4111 1111 1111 1111\n"
    )

    deck, demo = build_deck(
        source_type="document",
        idea="",
        uploaded_bytes=source.encode("utf-8"),
        uploaded_name="security_test.txt",
        audience="Hackathon judges",
        tone="Professional",
        duration_minutes=5,
        slide_count=4,
        qa_count=2,
    )

    sanitized = captured["text"]

    assert "secret@example.com" not in sanitized
    assert "9876543210" not in sanitized
    assert "ABCDE1234F" not in sanitized
    assert "1234 5678 9012" not in sanitized
    assert "4111 1111 1111 1111" not in sanitized

    assert "[EMAIL]" in sanitized
    assert "[PHONE]" in sanitized
    assert "[PAN]" in sanitized
    assert "[AADHAAR]" in sanitized
    assert "[CARD]" in sanitized

    print("\nSECURITY BOUNDARY PASSED")
    print("LLM received:")
    print(sanitized)