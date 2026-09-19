"""db/models.py — Tables. Portable types only (JSON not JSONB) so Supabase and SQLite both work.  Owner: P2
presentations 1 ─── * slides
presentations 1 ─── * audience_questions
presentations 1 ─── * source_documents
"""
from datetime import datetime, timezone

from sqlalchemy import JSON, Column, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from db.database import Base


def _now():
    return datetime.now(timezone.utc)


class PresentationRow(Base):
    __tablename__ = "presentations"
    id = Column(Integer, primary_key=True)
    title = Column(String(200), nullable=False)
    audience = Column(String(100))
    tone = Column(String(50))
    duration_min = Column(Integer)
    input_type = Column(String(20))          # "idea" | "document"
    created_at = Column(DateTime(timezone=True), default=_now)

    slides = relationship("SlideRow", back_populates="presentation",
                          cascade="all, delete-orphan", order_by="SlideRow.slide_no")
    questions = relationship("QuestionRow", back_populates="presentation",
                             cascade="all, delete-orphan")
    documents = relationship("SourceDocumentRow", back_populates="presentation",
                             cascade="all, delete-orphan")


class SlideRow(Base):
    __tablename__ = "slides"
    id = Column(Integer, primary_key=True)
    presentation_id = Column(Integer, ForeignKey("presentations.id", ondelete="CASCADE"), index=True)
    slide_no = Column(Integer, nullable=False)
    title = Column(String(150), nullable=False)
    key_points = Column(JSON, nullable=False)
    speaker_notes = Column(Text)
    visual_suggestion = Column(Text)
    presentation = relationship("PresentationRow", back_populates="slides")


class QuestionRow(Base):
    __tablename__ = "audience_questions"
    id = Column(Integer, primary_key=True)
    presentation_id = Column(Integer, ForeignKey("presentations.id", ondelete="CASCADE"), index=True)
    question = Column(Text, nullable=False)
    suggested_answer = Column(Text)
    presentation = relationship("PresentationRow", back_populates="questions")


class SourceDocumentRow(Base):
    __tablename__ = "source_documents"
    id = Column(Integer, primary_key=True)
    presentation_id = Column(Integer, ForeignKey("presentations.id", ondelete="CASCADE"), index=True)
    filename = Column(String(255))
    content_hash = Column(String(64), index=True)   # SHA-256 of original text (for caching/dedup)
    redacted_text = Column(Text)                    # ONLY redacted text is ever stored
    presentation = relationship("PresentationRow", back_populates="documents")
