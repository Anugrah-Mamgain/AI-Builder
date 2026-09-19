# P3 integration contract

`app.py` contains UI only. `services/pipeline.py` owns orchestration and runs with
mock data, session history, and a built-in PPTX generator until P1/P2 merge.

## P1: LLM functions

`llm/generator.py` should expose:

```python
def generate_presentation(source_text: str, **settings): ...
def regenerate_slide(slide, slide_index: int, deck, **settings): ...
```

The deck can be a dictionary or a Pydantic model and needs `title`, `subtitle`,
`slides`, and `qa`. Each slide needs `title`, `bullets`, `speaker_notes`, and
`visual_suggestion`. Each Q&A item needs `question` and `answer`.

## P2: document, privacy, database, and export functions

```python
utils.parser.extract_text(file_bytes, filename)
utils.pii.redact_pii(text)
utils.exporter.export_pptx(deck)
db.crud.save_presentation(deck=deck, **settings)
db.crud.get_history()
db.crud.get_presentation(presentation_id)
```

`export_pptx` may return `bytes`, `io.BytesIO`, or a `.pptx` path.

## Commands

```bash
pip install -r requirements.txt
python smoke_test.py
streamlit run app.py
```
