"""utils/parser.py — Extract text from uploads.  Owner: P2
Use pypdf for PDF, python-docx for DOCX. Enforce config.MAX_UPLOAD_MB and config.ALLOWED_EXTENSIONS.
Raise ValueError with a user-friendly message for bad/empty/scanned files.
"""


def extract_text(uploaded_file) -> str:
    """uploaded_file: Streamlit UploadedFile (has .name, .size, .read()). Return plain text."""
    raise NotImplementedError("P2: implement extract_text")
