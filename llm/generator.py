"""
generator.py — Public interface for the rest of the team:

    generate_presentation(text, audience, tone, duration_min, num_questions, input_type) -> Presentation
    regenerate_slide(slide, instruction, deck_title, audience) -> Slide

Both return validated Pydantic objects. Use .model_dump() to get a plain dict.
Both raise GenerationError with a user-friendly message if generation fails.
"""
import json
import os
import re
import time
from typing import Type, TypeVar

from dotenv import load_dotenv
from openai import OpenAI
from pydantic import BaseModel, ValidationError

import config as cfg
from schemas import Presentation, Slide
from utils.logger import get_logger
from llm.prompts import (
    SYSTEM_PROMPT, SUMMARY_SYSTEM_PROMPT,
    build_generation_prompt, build_summary_prompt, build_regenerate_prompt,
)

load_dotenv()
log = get_logger(__name__)

# ---------- Step 3: one client for any OpenAI-compatible provider ----------
PROVIDER_URLS = {
    "openai": None,
    "groq": "https://api.groq.com/openai/v1",
    "gemini": "https://generativelanguage.googleapis.com/v1beta/openai/",
}
PROVIDER = os.getenv("LLM_PROVIDER", "groq").lower()
MODEL = os.getenv("LLM_MODEL", "")
API_KEY = os.getenv("LLM_API_KEY", "")

_client = None


def _get_client() -> OpenAI:
    global _client
    if _client is None:
        if not API_KEY or not MODEL:
            raise GenerationError("LLM is not configured. Set LLM_API_KEY and LLM_MODEL in .env.")
        _client = OpenAI(
            api_key=API_KEY,
            base_url=PROVIDER_URLS.get(PROVIDER),
            timeout=cfg.LLM_TIMEOUT_SECONDS,   # never hang the demo
            max_retries=cfg.LLM_SDK_RETRIES,   # SDK auto-retries network errors / rate limits
        )
    return _client


T = TypeVar("T", bound=BaseModel)


class GenerationError(Exception):
    """Raised with a message that is safe to show directly in the UI."""


def _call_llm(system: str, user: str, temperature: float = 0.4, json_mode: bool = True) -> str:
    kwargs = {
        "model": MODEL,
        "messages": [{"role": "system", "content": system},
                     {"role": "user", "content": user}],
        "temperature": temperature,
    }
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    try:
        resp = _get_client().chat.completions.create(**kwargs)
    except GenerationError:
        raise
    except Exception as e:
        log.error(f"LLM call failed: {type(e).__name__}")
        raise GenerationError(f"The AI service is unavailable right now ({type(e).__name__}). Please retry.") from e
    return (resp.choices[0].message.content or "").strip()


# ---------- Step 4: parse + validate + self-correcting retry ----------
def _extract_json(raw: str) -> dict:
    """Strip code fences / stray text and parse the outermost JSON object."""
    raw = re.sub(r"^```(?:json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end <= start:
        raise ValueError("No JSON object found in model output")
    return json.loads(raw[start:end + 1])


def _generate_validated(system: str, user: str, model_cls: Type[T],
                        temperature: float, attempts: int = 2) -> T:
    prompt = user
    last_error = ""
    for _ in range(attempts):
        raw = _call_llm(system, prompt, temperature=temperature)
        try:
            return model_cls.model_validate(_extract_json(raw))
        except (ValueError, ValidationError) as e:
            last_error = str(e)[:800]
            log.warning(f"Invalid LLM output for {model_cls.__name__}, retrying")
            # Feed the exact error back so the model fixes its own output.
            prompt = (f"{user}\n\nYour previous response was INVALID. Error:\n{last_error}\n"
                      f"Return ONLY the corrected JSON object.")
    log.error(f"LLM output invalid after {attempts} attempts")
    raise GenerationError("The AI returned an invalid structure twice. Please try again "
                          "or simplify the input.")


# ---------- Step 5: long documents (map-reduce summarisation) ----------
def _split_chunks(text: str, size: int) -> list[str]:
    chunks, start = [], 0
    while start < len(text):
        end = min(start + size, len(text))
        if end < len(text):                       # prefer to cut at a paragraph/line break
            cut = text.rfind("\n", start + size // 2, end)
            end = cut if cut != -1 else end
        chunks.append(text[start:end].strip())
        start = end
    return [c for c in chunks if c]


def prepare_content(text: str) -> str:
    if len(text) <= cfg.MAX_DIRECT_CHARS:
        return text
    chunks = _split_chunks(text, cfg.CHUNK_CHARS)[:cfg.MAX_CHUNKS]
    log.info(f"Long input ({len(text)} chars): summarising {len(chunks)} chunks")
    summaries = [
        _call_llm(SUMMARY_SYSTEM_PROMPT, build_summary_prompt(c, i, len(chunks)),
                  temperature=cfg.TEMP_SUMMARY, json_mode=False)
        for i, c in enumerate(chunks, start=1)
    ]
    return "\n\n".join(summaries)


def calc_num_slides(duration_min: int) -> int:
    # Rule of thumb: ~1.5 minutes per slide, clamped to a sensible range.
    return max(cfg.MIN_SLIDES, min(cfg.MAX_SLIDES, round(duration_min / cfg.MINUTES_PER_SLIDE)))


# ---------- Step 6: main public function ----------
def generate_presentation(text: str, audience: str = "General", tone: str = "Professional",
                          duration_min: int = 10, num_questions: int = 5,
                          input_type: str = "idea") -> Presentation:
    if not text or not text.strip():
        raise GenerationError("Please enter an idea or upload a document with readable text.")
    if len(text.strip()) < 15:
        raise GenerationError("The input is too short. Add a sentence or two describing your topic.")

    start = time.time()
    num_questions = max(0, min(cfg.MAX_QUESTIONS, int(num_questions)))
    num_slides = calc_num_slides(int(duration_min))
    content = prepare_content(text.strip())

    user_prompt = build_generation_prompt(content, audience, tone, num_slides,
                                          num_questions, input_type)
    deck = _generate_validated(SYSTEM_PROMPT, user_prompt, Presentation,
                               temperature=cfg.TEMP_GENERATE)
    deck.audience = audience   # trust our input, not the model's echo
    # Log metadata only — never the content (PII safety)
    log.info(f"Deck generated: {len(text)} chars in, {len(deck.slides)} slides, "
             f"{len(deck.audience_questions)} questions, {time.time() - start:.1f}s")
    return deck


# ---------- Step 7: regenerate a single slide ----------
def regenerate_slide(slide: Slide, instruction: str, deck_title: str = "",
                     audience: str = "General") -> Slide:
    if not instruction or not instruction.strip():
        instruction = "Improve clarity and impact."
    user_prompt = build_regenerate_prompt(slide.model_dump_json(indent=2), instruction.strip(),
                                          deck_title, audience)
    new_slide = _generate_validated(SYSTEM_PROMPT, user_prompt, Slide,
                                    temperature=cfg.TEMP_REGENERATE)
    log.info(f"Slide {slide.slide_no} regenerated")
    new_slide.slide_no = slide.slide_no   # keep its position in the deck
    return new_slide
