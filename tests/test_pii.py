from utils.pii import redact_pii


def test_pii_redaction():
    text = (
        "Email: secret@example.com\n"
        "Phone: +91 9876543210\n"
        "Aadhaar: 1234 5678 9012\n"
        "PAN: ABCDE1234F\n"
        "Card: 4111 1111 1111 1111"
    )

    result = redact_pii(text)

    assert "secret@example.com" not in result
    assert "9876543210" not in result
    assert "1234 5678 9012" not in result
    assert "ABCDE1234F" not in result
    assert "4111 1111 1111 1111" not in result

    assert "[EMAIL]" in result
    assert "[PHONE]" in result
    assert "[AADHAAR]" in result
    assert "[PAN]" in result
    assert "[CARD]" in result