"""config.py — All tunable settings in one place (no secrets here; secrets live in .env)."""

# LLM input handling
MAX_DIRECT_CHARS = 12_000   # above this, summarise first (map-reduce)
CHUNK_CHARS = 8_000
MAX_CHUNKS = 6

# Deck shape
MIN_SLIDES, MAX_SLIDES = 3, 15
MINUTES_PER_SLIDE = 1.5
MAX_QUESTIONS = 10

# Temperatures
TEMP_GENERATE = 0.4
TEMP_REGENERATE = 0.6
TEMP_SUMMARY = 0.2

# LLM client
LLM_TIMEOUT_SECONDS = 60
LLM_SDK_RETRIES = 2

# Uploads
MAX_UPLOAD_MB = 5
ALLOWED_EXTENSIONS = {".pdf", ".docx", ".txt"}

# Database
SQLITE_PATH = "data/app.db"
