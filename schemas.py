"""
schemas.py — The frozen JSON contract, enforced in code.
Every LLM response must pass these models before it reaches the UI or DB.
"""
from typing import List
from pydantic import BaseModel, Field, field_validator


class Slide(BaseModel):
    slide_no: int = Field(ge=1)
    title: str = Field(min_length=1, max_length=150)
    key_points: List[str] = Field(min_length=2, max_length=6)
    speaker_notes: str = Field(min_length=20)
    visual_suggestion: str = Field(min_length=5)


class AudienceQuestion(BaseModel):
    question: str = Field(min_length=5)
    suggested_answer: str = Field(min_length=10)


class Presentation(BaseModel):
    title: str = Field(min_length=1, max_length=200)
    audience: str
    slides: List[Slide] = Field(min_length=1)
    audience_questions: List[AudienceQuestion] = Field(default_factory=list)

    @field_validator("slides")
    @classmethod
    def renumber_slides(cls, slides: List[Slide]) -> List[Slide]:
        # LLMs sometimes skip or repeat numbers; force a clean 1..N sequence.
        for i, slide in enumerate(slides, start=1):
            slide.slide_no = i
        return slides
