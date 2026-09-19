"""Full pipeline WITHOUT the UI. Run from project root:  python smoke_test.py
If this passes but the app fails, the bug is in app.py."""
import time

from db.database import DB_MODE, init_db
from services.pipeline import build_deck

init_db()
IDEA = ("An AI assistant that turns ideas or documents into ready-to-present slide decks "
        "with speaker notes, visual suggestions and likely audience questions. "
        "Contact: test.user@example.com, +91 98765 43210")   # PII should be redacted

t = time.time()
deck, pres_id = build_deck(idea_text=IDEA, audience="Executives", tone="Persuasive",
                           duration_min=8, num_questions=3)
print(f"OK in {time.time() - t:.1f}s | DB={DB_MODE} | id={pres_id} | "
      f"{len(deck.slides)} slides | {len(deck.audience_questions)} questions")
for s in deck.slides:
    print(f"  {s.slide_no}. {s.title}")
