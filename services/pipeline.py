"""P3 orchestration: parse -> redact -> generate -> save -> export.

This version integrates P2 document security with the current P1/P3 contracts:
- documents are parsed before generation
- PII redaction is mandatory before LLM/database use
- original uploaded bytes are hashed before parsing/redaction
- only redacted text is persisted
- LLM failures may use the in-memory mock generator
- database failures fall back to session history
"""

from __future__ import annotations

import copy
import io
import logging
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


LOGGER = logging.getLogger(__name__)
_DEMO_HISTORY: list[dict[str, Any]] = []


MOCK_DECK = {
    "title": "AI in Indian Healthcare",
    "subtitle": "From operational efficiency to better patient outcomes",
    "slides": [
        {
            "title": "The Opportunity",
            "bullets": [
                "Growing patient demand is stretching hospital capacity",
                "Clinical teams lose time to repetitive administrative work",
                "Responsible AI augments decisions without replacing clinicians",
            ],
            "key_points": [
                "Growing patient demand is stretching hospital capacity",
                "Clinical teams lose time to repetitive administrative work",
                "Responsible AI augments decisions without replacing clinicians",
            ],
            "speaker_notes": "Lead with the human impact: more time for patients.",
            "visual_suggestion": "Before-and-after hospital workflow",
        },
        {
            "title": "Where AI Creates Value",
            "bullets": [
                "Automate appointment triage and routine documentation",
                "Predict capacity needs and optimize resources",
                "Surface relevant information at the point of care",
            ],
            "key_points": [
                "Automate appointment triage and routine documentation",
                "Predict capacity needs and optimize resources",
                "Surface relevant information at the point of care",
            ],
            "speaker_notes": "Tie every use case to a measurable KPI.",
            "visual_suggestion": "Patient, clinician, and hospital value map",
        },
        {
            "title": "Responsible Implementation",
            "bullets": [
                "Keep clinicians in the loop for consequential decisions",
                "Minimize and encrypt personally identifiable information",
                "Monitor accuracy, bias, latency, and adoption",
            ],
            "key_points": [
                "Keep clinicians in the loop for consequential decisions",
                "Minimize and encrypt personally identifiable information",
                "Monitor accuracy, bias, latency, and adoption",
            ],
            "speaker_notes": "Governance is part of product design.",
            "visual_suggestion": "Circular governance diagram",
        },
        {
            "title": "90-Day Pilot Roadmap",
            "bullets": [
                "Weeks 1–3: select one workflow and baseline metrics",
                "Weeks 4–8: launch a controlled pilot with training",
                "Weeks 9–12: measure impact and decide whether to scale",
            ],
            "key_points": [
                "Weeks 1–3: select one workflow and baseline metrics",
                "Weeks 4–8: launch a controlled pilot with training",
                "Weeks 9–12: measure impact and decide whether to scale",
            ],
            "speaker_notes": "Close with a low-risk next step.",
            "visual_suggestion": "Horizontal roadmap with three milestones",
        },
    ],
    "qa": [
        {
            "question": "How is sensitive information protected?",
            "answer": "Use minimization, encryption, access controls, logs, and retention limits.",
        },
        {
            "question": "How will we measure success?",
            "answer": "Compare time, quality, satisfaction, adoption, and cost against a baseline.",
        },
        {
            "question": "Will AI replace staff?",
            "answer": "AI assists repetitive work; accountable humans retain final control.",
        },
    ],
    "audience_questions": [
        {
            "question": "How is sensitive information protected?",
            "suggested_answer": "Use minimization, encryption, access controls, logs, and retention limits.",
        },
        {
            "question": "How will we measure success?",
            "suggested_answer": "Compare time, quality, satisfaction, adoption, and cost against a baseline.",
        },
        {
            "question": "Will AI replace staff?",
            "suggested_answer": "AI assists repetitive work; accountable humans retain final control.",
        },
    ],
}


def _as_dict(model: Any) -> dict[str, Any]:
    """Return a deep-copied plain dict and normalize legacy aliases."""
    if model is None:
        return {}

    if isinstance(model, dict):
        result = copy.deepcopy(model)
    elif hasattr(model, "model_dump"):
        result = model.model_dump()
    elif hasattr(model, "dict"):
        result = model.dict()
    else:
        raise TypeError("Presentation must be a dictionary or Pydantic model.")

    slides = result.get("slides", [])
    if isinstance(slides, list):
        for slide in slides:
            if not isinstance(slide, dict):
                continue

            if "key_points" not in slide and "bullets" in slide:
                slide["key_points"] = slide["bullets"]
            if "bullets" not in slide and "key_points" in slide:
                slide["bullets"] = slide["key_points"]

    questions = result.get("audience_questions")
    if questions is not None and "qa" not in result:
        result["qa"] = [
            {
                "question": q.get("question", ""),
                "answer": q.get(
                    "suggested_answer",
                    q.get("answer", ""),
                ),
            }
            for q in questions
            if isinstance(q, dict)
        ]
    elif "qa" in result and "audience_questions" not in result:
        result["audience_questions"] = [
            {
                "question": q.get("question", ""),
                "suggested_answer": q.get(
                    "answer",
                    q.get("suggested_answer", ""),
                ),
            }
            for q in result["qa"]
            if isinstance(q, dict)
        ]

    return result


def _parse_document(data: bytes, name: str) -> tuple[str, bool]:
    """Parse a document.

    This is fail-closed: if the parser is unavailable or unimplemented,
    document processing stops instead of sending raw text downstream.
    """
    try:
        from utils.parser import extract_text
    except (ImportError, ModuleNotFoundError) as exc:
        raise RuntimeError(
            "Document parser is unavailable. Document processing was stopped."
        ) from exc

    try:
        return extract_text(data, name), False
    except TypeError:
        # Compatibility with Streamlit UploadedFile-style signatures.
        file_obj = io.BytesIO(data)
        file_obj.name = name
        file_obj.size = len(data)
        return extract_text(file_obj), False
    except NotImplementedError as exc:
        raise RuntimeError(
            "Document parser is not implemented. Document processing was stopped."
        ) from exc


def _redact(text: str) -> tuple[str, bool]:
    """Redact PII before text reaches the LLM or database.

    This is also fail-closed: a redaction failure stops document processing.
    """
    try:
        from utils.pii import redact_pii
    except (ImportError, ModuleNotFoundError) as exc:
        raise RuntimeError(
            "PII redaction is unavailable. Document processing was stopped."
        ) from exc

    try:
        return redact_pii(text), False
    except NotImplementedError as exc:
        raise RuntimeError(
            "PII redaction is not implemented. Document processing was stopped."
        ) from exc


def _mock_generate(
    text: str,
    slide_count: int,
    qa_count: int,
) -> dict[str, Any]:
    """Generate a deterministic in-memory fallback deck."""
    topic = (
        text.strip().splitlines()[0][:80]
        if text.strip()
        else "Generated presentation"
    )

    deck = {
        "title": topic,
        "subtitle": "A focused proposal tailored to the selected audience",
        "slides": [
            {
                "title": "Problem and context",
                "bullets": [
                    f"{topic} addresses a clear operational or user need",
                    "Current approaches require avoidable time and manual effort",
                    "A focused first release can demonstrate value quickly",
                ],
                "key_points": [
                    f"{topic} addresses a clear operational or user need",
                    "Current approaches require avoidable time and manual effort",
                    "A focused first release can demonstrate value quickly",
                ],
                "speaker_notes": (
                    "Describe the current situation and the people most "
                    "affected by it."
                ),
                "visual_suggestion": (
                    "A before-and-after comparison of the current "
                    "and proposed experience"
                ),
            },
            {
                "title": "User needs",
                "bullets": [
                    "A simple experience that requires little training",
                    "Reliable results with clear explanations and human control",
                    "Privacy safeguards throughout the workflow",
                ],
                "key_points": [
                    "A simple experience that requires little training",
                    "Reliable results with clear explanations and human control",
                    "Privacy safeguards throughout the workflow",
                ],
                "speaker_notes": (
                    "Connect each need to a specific user or stakeholder."
                ),
                "visual_suggestion": (
                    "Three user profiles with one priority highlighted for each"
                ),
            },
            {
                "title": "Proposed solution",
                "bullets": [
                    f"A practical application centred on {topic}",
                    "Guided inputs produce structured, actionable outputs",
                    "Users can review and refine results before using them",
                ],
                "key_points": [
                    f"A practical application centred on {topic}",
                    "Guided inputs produce structured, actionable outputs",
                    "Users can review and refine results before using them",
                ],
                "speaker_notes": (
                    "Explain the solution in one sentence before discussing features."
                ),
                "visual_suggestion": (
                    "Product workflow from user input to reviewed output"
                ),
            },
            {
                "title": "How it works",
                "bullets": [
                    "Capture the user's goal and relevant source material",
                    "Process the request with validation and privacy controls",
                    "Return an editable result and preserve useful history",
                ],
                "key_points": [
                    "Capture the user's goal and relevant source material",
                    "Process the request with validation and privacy controls",
                    "Return an editable result and preserve useful history",
                ],
                "speaker_notes": (
                    "Walk through one realistic example from beginning to end."
                ),
                "visual_suggestion": "Four-step horizontal process diagram",
            },
            {
                "title": "Implementation plan",
                "bullets": [
                    "Begin with one high-value workflow and measurable baseline",
                    "Run a controlled pilot and collect user feedback",
                    "Improve the product before expanding adoption",
                ],
                "key_points": [
                    "Begin with one high-value workflow and measurable baseline",
                    "Run a controlled pilot and collect user feedback",
                    "Improve the product before expanding adoption",
                ],
                "speaker_notes": (
                    "Position the pilot as a low-risk way to validate the proposal."
                ),
                "visual_suggestion": (
                    "Phased roadmap with pilot, evaluation, and expansion"
                ),
            },
            {
                "title": "Expected impact",
                "bullets": [
                    "Less time spent on repetitive work",
                    "More consistent and useful outputs",
                    "Clear evidence for the next investment decision",
                ],
                "key_points": [
                    "Less time spent on repetitive work",
                    "More consistent and useful outputs",
                    "Clear evidence for the next investment decision",
                ],
                "speaker_notes": (
                    "Use measured pilot outcomes instead of unsupported projections."
                ),
                "visual_suggestion": (
                    "Three large outcome metrics with baseline placeholders"
                ),
            },
        ],
        "qa": copy.deepcopy(MOCK_DECK["qa"]),
        "audience_questions": copy.deepcopy(
            MOCK_DECK["audience_questions"]
        ),
    }

    while len(deck["slides"]) < slide_count:
        number = len(deck["slides"]) + 1
        points = [
            "Connect to a measurable user need",
            "Start with a focused pilot",
            "Track adoption, quality, cost, and impact",
        ]
        deck["slides"].append(
            {
                "title": f"Strategic Insight {number}",
                "bullets": points,
                "key_points": points,
                "speaker_notes": "Use one concrete example.",
                "visual_suggestion": (
                    "Metric cards or a before-and-after comparison"
                ),
            }
        )

    deck["slides"] = deck["slides"][:slide_count]
    deck["qa"] = deck["qa"][:qa_count]
    deck["audience_questions"] = deck["audience_questions"][:qa_count]

    return deck


def _generate(
    text: str,
    settings: dict[str, Any],
    input_type: str = "idea",
) -> tuple[Any, bool]:
    """Generate a validated deck or fall back to the demo generator."""
    slide_count = settings.get("slide_count", 6)
    qa_count = settings.get("qa_count", 3)
    audience = settings.get("audience", "General")
    tone = settings.get("tone", "Professional")
    duration_min = settings.get("duration_minutes", 10)

    try:
        from llm.generator import generate_presentation
    except (ImportError, ModuleNotFoundError):
        return _mock_generate(text, slide_count, qa_count), True

    try:
        deck = generate_presentation(
            text=text,
            audience=audience,
            tone=tone,
            duration_min=duration_min,
            num_questions=qa_count,
            input_type=input_type,
        )
        return deck, False
    except Exception as exc:
        # Do not log prompt/text/response contents.
        LOGGER.warning(
            "LLM generation failed; using mock generator (%s).",
            type(exc).__name__,
        )
        return _mock_generate(text, slide_count, qa_count), True


def _save(
    deck: Any,
    settings: dict[str, Any],
    *,
    input_type: str = "idea",
    filename: str | None = None,
    content_hash: str | None = None,
    redacted_text: str | None = None,
) -> bool:
    """Persist a deck and sanitized source metadata, or use session history."""
    try:
        from db.crud import save_presentation
        from schemas import Presentation
    except (ImportError, ModuleNotFoundError):
        _DEMO_HISTORY.insert(
            0,
            {
                "id": len(_DEMO_HISTORY) + 1,
                "title": _as_dict(deck).get(
                    "title",
                    "Untitled presentation",
                ),
                "audience": settings.get("audience", "General"),
                "tone": settings.get("tone", "Professional"),
                "created_at": "Current session",
                "deck": copy.deepcopy(deck),
            },
        )
        return True

    try:
        if isinstance(deck, dict):
            deck_data = _as_dict(deck)
            save_deck = Presentation(
                title=deck_data.get(
                    "title",
                    "Untitled Presentation",
                ),
                audience=settings.get("audience", "General"),
                slides=[
                    {
                        "slide_no": index + 1,
                        "title": slide.get("title", "Slide"),
                        "key_points": slide.get(
                            "key_points",
                            slide.get(
                                "bullets",
                                ["Point 1", "Point 2"],
                            ),
                        ),
                        "speaker_notes": slide.get(
                            "speaker_notes",
                            "Notes here...",
                        ),
                        "visual_suggestion": slide.get(
                            "visual_suggestion",
                            "Visual suggestion",
                        ),
                    }
                    for index, slide in enumerate(
                        deck_data.get("slides", [])
                    )
                ],
                audience_questions=[
                    {
                        "question": question.get(
                            "question",
                            "Question",
                        ),
                        "suggested_answer": question.get(
                            "suggested_answer",
                            question.get("answer", "Answer"),
                        ),
                    }
                    for question in (
                        deck_data.get("audience_questions")
                        or deck_data.get("qa")
                        or []
                    )
                ],
            )
        else:
            save_deck = deck

        save_presentation(
            deck=save_deck,
            tone=settings.get("tone", "Professional"),
            duration_min=settings.get("duration_minutes", 10),
            input_type=input_type,
            filename=filename,
            redacted_text=redacted_text,
            content_hash=content_hash,
        )
        return False

    except Exception as exc:
        # DB failure must never cause raw text to be logged.
        LOGGER.warning(
            "Database save failed; using session history (%s).",
            type(exc).__name__,
        )
        _DEMO_HISTORY.insert(
            0,
            {
                "id": len(_DEMO_HISTORY) + 1,
                "title": _as_dict(deck).get(
                    "title",
                    "Untitled presentation",
                ),
                "audience": settings.get("audience", "General"),
                "tone": settings.get("tone", "Professional"),
                "created_at": "Current session",
                "deck": copy.deepcopy(deck),
            },
        )
        return True


def build_deck(
    *,
    source_type: str,
    idea: str,
    uploaded_bytes: bytes | None,
    uploaded_name: str | None,
    audience: str,
    tone: str,
    duration_minutes: int,
    slide_count: int,
    qa_count: int,
) -> tuple[Any, bool]:
    """Run parse -> redact -> generate -> save.

    For document input:
    - hash is calculated from original upload bytes
    - text is parsed in memory
    - PII is redacted before generation/storage
    """
    settings = {
        "audience": audience,
        "tone": tone,
        "duration_minutes": duration_minutes,
        "slide_count": slide_count,
        "qa_count": qa_count,
    }

    demo = False
    filename: str | None = None
    content_hash: str | None = None

    if source_type == "document":
        if not uploaded_bytes or not uploaded_name:
            raise ValueError("Please upload a source document.")

        from utils.parser import content_hash as calculate_content_hash

        content_hash = calculate_content_hash(
            uploaded_bytes,
            uploaded_name,
        )
        filename = uploaded_name

        text, fallback = _parse_document(
            uploaded_bytes,
            uploaded_name,
        )
        demo |= fallback
    else:
        text = idea.strip()

    if not text:
        raise ValueError("Please provide some source content.")

    # Mandatory security boundary before any LLM or DB operation.
    text, fallback = _redact(text)
    demo |= fallback

    deck, fallback = _generate(
        text,
        settings,
        input_type=source_type,
    )
    demo |= fallback

    demo |= _save(
        deck,
        settings,
        input_type=source_type,
        filename=filename,
        content_hash=content_hash,
        redacted_text=text if source_type == "document" else None,
    )

    LOGGER.info("Presentation pipeline completed")
    return deck, demo


def regenerate_deck_slide(
    *,
    deck: Any,
    slide_index: int,
    settings: dict[str, Any],
    instruction: str = "",
) -> tuple[Any, bool]:
    """Regenerate one slide using the current LLM contract."""
    deck_dict = _as_dict(deck)
    slides = deck_dict.get("slides", [])

    if not 0 <= slide_index < len(slides):
        raise IndexError("Slide index is out of range.")

    try:
        from llm.generator import regenerate_slide
        from schemas import Slide

        current_slide = slides[slide_index]

        if isinstance(current_slide, dict):
            slide_obj = Slide(
                slide_no=slide_index + 1,
                title=current_slide.get(
                    "title",
                    f"Slide {slide_index + 1}",
                ),
                key_points=current_slide.get(
                    "key_points",
                    current_slide.get(
                        "bullets",
                        ["Point 1", "Point 2"],
                    ),
                ),
                speaker_notes=current_slide.get(
                    "speaker_notes",
                    "Notes here...",
                ),
                visual_suggestion=current_slide.get(
                    "visual_suggestion",
                    "Visual suggestion",
                ),
            )
        else:
            slide_obj = current_slide

        replacement = regenerate_slide(
            slide=slide_obj,
            instruction=instruction,
            deck_title=deck_dict.get("title", ""),
            audience=settings.get("audience", "General"),
        )

        if isinstance(deck, dict):
            deck["slides"][slide_index] = _as_dict(replacement)
        else:
            deck.slides[slide_index] = replacement

        return deck, False

    except Exception as exc:
        LOGGER.warning(
            "Slide regeneration failed; using mock fallback (%s).",
            type(exc).__name__,
        )

        if isinstance(deck, dict):
            target = deck["slides"][slide_index]
            points = target.get(
                "key_points",
                target.get("bullets", []),
            )
            target["key_points"] = (
                list(reversed(points))
                if points
                else ["Updated point 1", "Updated point 2"]
            )
            target["bullets"] = target["key_points"]
            target["speaker_notes"] = (
                target.get("speaker_notes", "")
                + " Emphasize the audience outcome."
            ).strip()
        else:
            target = deck.slides[slide_index]
            target.key_points = list(reversed(target.key_points))
            target.speaker_notes = (
                target.speaker_notes
                + " Emphasize the audience outcome."
            ).strip()

        return deck, True


def get_history_items() -> list[Any]:
    try:
        from db.crud import get_history

        items = get_history()
        if items:
            return items
    except Exception as exc:
        LOGGER.warning(
            "Database history unavailable; returning session history (%s).",
            type(exc).__name__,
        )

    return _DEMO_HISTORY


def load_history_deck(presentation_id: Any) -> Any:
    try:
        from db.crud import get_presentation

        presentation = get_presentation(int(presentation_id))
        if presentation:
            return presentation
    except Exception as exc:
        LOGGER.warning(
            "Database presentation unavailable; checking session history (%s).",
            type(exc).__name__,
        )

    for record in _DEMO_HISTORY:
        if str(record["id"]) == str(presentation_id):
            return copy.deepcopy(record["deck"])

    raise KeyError("Presentation not found in history.")


def export_deck(deck: Any) -> tuple[bytes, bool]:
    """Export a deck using the P2 exporter, with a local fallback."""
    try:
        from schemas import Presentation
        from utils.exporter import export_pptx

        if isinstance(deck, dict):
            data = _as_dict(deck)
            presentation = Presentation(
                title=data.get(
                    "title",
                    "Untitled Presentation",
                ),
                audience=data.get(
                    "audience",
                    "General",
                ),
                slides=[
                    {
                        "slide_no": index + 1,
                        "title": slide.get(
                            "title",
                            "Slide",
                        ),
                        "key_points": slide.get(
                            "key_points",
                            slide.get(
                                "bullets",
                                ["Point 1", "Point 2"],
                            ),
                        ),
                        "speaker_notes": slide.get(
                            "speaker_notes",
                            "Notes here...",
                        ),
                        "visual_suggestion": slide.get(
                            "visual_suggestion",
                            "Visual suggestion",
                        ),
                    }
                    for index, slide in enumerate(
                        data.get("slides", [])
                    )
                ],
                audience_questions=[
                    {
                        "question": question.get(
                            "question",
                            "Question",
                        ),
                        "suggested_answer": question.get(
                            "suggested_answer",
                            question.get("answer", "Answer"),
                        ),
                    }
                    for question in (
                        data.get("audience_questions")
                        or data.get("qa")
                        or []
                    )
                ],
            )
        else:
            presentation = deck

        result = export_pptx(presentation)

        if isinstance(result, io.BytesIO):
            return result.getvalue(), False

        if isinstance(result, bytes):
            return result, False

        if isinstance(result, (str, Path)):
            return Path(result).read_bytes(), False

        raise TypeError(
            "export_pptx() must return bytes, BytesIO, or a file path."
        )

    except Exception as exc:
        LOGGER.warning(
            "PPTX export failed; using fallback exporter (%s).",
            type(exc).__name__,
        )
        return _fallback_export(deck), True


def _fallback_export(deck: Any) -> bytes:
    """Small local PPTX fallback used only when the main exporter is unavailable."""
    data = _as_dict(deck)

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)

    navy = RGBColor(15, 23, 42)
    indigo = RGBColor(79, 70, 229)
    violet = RGBColor(124, 58, 237)
    pink = RGBColor(219, 39, 119)
    white = RGBColor(255, 255, 255)
    ink = RGBColor(30, 41, 59)
    muted = RGBColor(100, 116, 139)
    pale = RGBColor(238, 242, 255)
    font = "Aptos"

    def add_text(
        slide,
        value,
        left,
        top,
        width,
        height,
        size,
        color,
        bold=False,
        align=PP_ALIGN.LEFT,
    ):
        box = slide.shapes.add_textbox(
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(height),
        )
        frame = box.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.margin_left = 0
        frame.margin_right = 0
        frame.margin_top = 0
        frame.margin_bottom = 0

        paragraph = frame.paragraphs[0]
        paragraph.text = str(value)
        paragraph.alignment = align
        paragraph.font.name = font
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        return box

    def add_footer(slide, number):
        add_text(
            slide,
            "PITCHPILOT AI",
            0.68,
            7.05,
            2.1,
            0.22,
            8,
            muted,
            True,
        )
        add_text(
            slide,
            f"{number:02d}",
            12.05,
            7.0,
            0.6,
            0.25,
            9,
            muted,
            True,
            PP_ALIGN.RIGHT,
        )

    title_slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )
    title_slide.background.fill.solid()
    title_slide.background.fill.fore_color.rgb = navy

    accent = title_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(0.18),
        Inches(7.5),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = pink
    accent.line.fill.background()

    add_text(
        title_slide,
        "PITCHPILOT AI",
        0.85,
        0.72,
        2.6,
        0.35,
        12,
        RGBColor(196, 181, 253),
        True,
    )

    add_text(
        title_slide,
        data.get("title", "Generated presentation"),
        0.85,
        2.05,
        11.2,
        1.7,
        34,
        white,
        True,
    )

    add_text(
        title_slide,
        data.get(
            "subtitle",
            "Created with PitchPilot AI",
        ),
        0.88,
        4.05,
        9.9,
        0.75,
        18,
        RGBColor(203, 213, 225),
    )

    add_text(
        title_slide,
        "AI-GENERATED PRESENTATION",
        0.88,
        6.55,
        3.6,
        0.3,
        9,
        RGBColor(148, 163, 184),
        True,
    )

    for slide_number, slide_data in enumerate(
        data.get("slides", []),
        start=2,
    ):
        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = white

        band = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(0.18),
            Inches(7.5),
        )
        band.fill.solid()
        band.fill.fore_color.rgb = (
            indigo if slide_number % 2 == 0 else violet
        )
        band.line.fill.background()

        add_text(
            slide,
            slide_data.get("title", "Untitled slide"),
            0.72,
            0.55,
            8.2,
            0.6,
            27,
            ink,
            True,
        )

        add_text(
            slide,
            "KEY POINTS",
            0.75,
            1.55,
            2,
            0.3,
            10,
            indigo,
            True,
        )

        bullets = (
            slide_data.get("key_points")
            or slide_data.get("bullets", [])
        )[:5]

        for bullet_index, bullet in enumerate(bullets):
            y = 2.0 + bullet_index * 1.05

            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(0.78),
                Inches(y + 0.08),
                Inches(0.18),
                Inches(0.18),
            )
            dot.fill.solid()
            dot.fill.fore_color.rgb = pink
            dot.line.fill.background()

            add_text(
                slide,
                bullet,
                1.15,
                y,
                7.15,
                0.72,
                18,
                ink,
            )

        panel = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(9.0),
            Inches(1.48),
            Inches(3.55),
            Inches(4.95),
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = pale
        panel.line.color.rgb = RGBColor(224, 231, 255)

        add_text(
            slide,
            "VISUAL DIRECTION",
            9.38,
            1.9,
            2.8,
            0.3,
            10,
            violet,
            True,
        )

        add_text(
            slide,
            slide_data.get(
                "visual_suggestion",
                "Supporting visual",
            ),
            9.38,
            2.45,
            2.75,
            1.75,
            18,
            ink,
            True,
        )

        add_text(
            slide,
            "Speaker note",
            9.38,
            4.8,
            2.5,
            0.3,
            10,
            muted,
            True,
        )

        note_preview = slide_data.get(
            "speaker_notes",
            "",
        )

        add_text(
            slide,
            note_preview,
            9.38,
            5.18,
            2.72,
            0.86,
            11,
            muted,
        )

        slide.notes_slide.notes_text_frame.text = note_preview
        add_footer(slide, slide_number)

    questions = (
        data.get("audience_questions")
        or data.get("qa")
        or []
    )

    if questions:
        number = len(presentation.slides) + 1

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = navy

        add_text(
            slide,
            "Questions to expect",
            0.8,
            0.65,
            8,
            0.65,
            28,
            white,
            True,
        )

        for index, item in enumerate(questions[:4]):
            y = 1.65 + index * 1.25

            add_text(
                slide,
                f"{index + 1:02d}",
                0.85,
                y,
                0.55,
                0.35,
                13,
                RGBColor(196, 181, 253),
                True,
            )

            add_text(
                slide,
                item.get("question", "Question"),
                1.55,
                y - 0.03,
                10.5,
                0.4,
                17,
                white,
                True,
            )

            add_text(
                slide,
                item.get(
                    "suggested_answer",
                    item.get("answer", ""),
                ),
                1.55,
                y + 0.43,
                10.25,
                0.55,
                12,
                RGBColor(203, 213, 225),
            )

        add_footer(slide, number)

    output = io.BytesIO()
    presentation.save(output)
    return output.getvalue()
