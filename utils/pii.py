"""utils/pii.py — Redact PII BEFORE text reaches the LLM or the cloud DB.  Owner: P2
Patterns to cover: emails, Indian phone numbers, PAN, Aadhaar, card numbers.
Replace with tags like [EMAIL], [PHONE], [PAN], [AADHAAR], [CARD].
"""


def redact_pii(text: str) -> str:
    raise NotImplementedError("P2: implement redact_pii")
