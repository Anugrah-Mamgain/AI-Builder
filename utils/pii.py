from __future__ import annotations

import hashlib
import re


EMAIL_RE = re.compile(
    r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
    re.IGNORECASE,
)

PHONE_RE = re.compile(
    r"(?<!\d)(?:\+?91[\s.-]?)?[6-9]\d{2}[\s.-]?\d{3}[\s.-]?\d{4}(?!\d)"
)

AADHAAR_RE = re.compile(
    r"(?<!\d)(?:\d[\s-]?){11}\d(?!\d)"
)

PAN_RE = re.compile(
    r"(?<![A-Z0-9])[A-Z]{5}\d{4}[A-Z](?![A-Z0-9])",
    re.IGNORECASE,
)

CARD_RE = re.compile(
    r"(?<!\d)(?:\d[\s-]?){12,18}\d(?!\d)"
)


def _luhn_valid(number: str) -> bool:
    digits = [int(char) for char in number if char.isdigit()]

    if not 13 <= len(digits) <= 19:
        return False

    checksum = 0
    parity = len(digits) % 2

    for index, digit in enumerate(digits):
        if index % 2 == parity:
            digit *= 2
            if digit > 9:
                digit -= 9
        checksum += digit

    return checksum % 10 == 0


def _redact_cards(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        candidate = match.group(0)
        digits = "".join(char for char in candidate if char.isdigit())

        if _luhn_valid(digits):
            return "[CARD]"

        return candidate

    return CARD_RE.sub(replace, text)


def redact_pii(text: str) -> str:
    """Redact sensitive identifiers before LLM/cloud-DB access."""

    if not isinstance(text, str):
        raise TypeError("text must be a string.")

    # Card numbers first.
    text = _redact_cards(text)

    text = EMAIL_RE.sub("[EMAIL]", text)
    text = PHONE_RE.sub("[PHONE]", text)
    text = AADHAAR_RE.sub("[AADHAAR]", text)
    text = PAN_RE.sub("[PAN]", text)

    return text


def content_hash(content: bytes | str) -> str:
    """Generate SHA-256 without persisting the original content."""

    if isinstance(content, str):
        content = content.encode("utf-8")

    return hashlib.sha256(content).hexdigest()