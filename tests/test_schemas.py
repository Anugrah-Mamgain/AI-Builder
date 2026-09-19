"""Offline tests — no API key needed.  Run: pytest -v"""
import pytest
from pydantic import ValidationError

from llm.generator import _extract_json, calc_num_slides
from schemas import Presentation

VALID = {"title": "T", "audience": "x", "slides": [
    {"slide_no": 7, "title": "A", "key_points": ["a", "b"],
     "speaker_notes": "Some speaker notes long enough.", "visual_suggestion": "Bar chart"}]}


def test_slides_renumbered():
    assert Presentation.model_validate(VALID).slides[0].slide_no == 1


def test_rejects_too_few_points():
    bad = {**VALID, "slides": [{**VALID["slides"][0], "key_points": ["only one"]}]}
    with pytest.raises(ValidationError):
        Presentation.model_validate(bad)


def test_extract_json_strips_fences():
    assert _extract_json('```json\n{"a": 1}\n```') == {"a": 1}


def test_slide_count_clamped():
    assert calc_num_slides(1) == 3 and calc_num_slides(100) == 15
