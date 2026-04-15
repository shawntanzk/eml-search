import os
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
ONTOLOGY_DIR = BASE_DIR / "ontology"
MODELS_DIR = BASE_DIR / "models"

for _d in [DATA_DIR, MODELS_DIR, ONTOLOGY_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

DB_PATH = str(DATA_DIR / "index.db")
GRAPH_DATA_PATH = str(DATA_DIR / "email_data.ttl")
ONTOLOGY_PATH = str(ONTOLOGY_DIR / "email_ontology.ttl")
SETTINGS_PATH = str(DATA_DIR / "settings.json")

# Default email folder — overridden by settings UI or env var
DEFAULT_EMAIL_FOLDER = os.environ.get("EML_FOLDER", str(BASE_DIR / "test_emails"))

# NLP
SPACY_MODEL = "en_core_web_sm"
SENTENCE_TRANSFORMER_MODEL = "all-MiniLM-L6-v2"

# Search
SEMANTIC_TOP_K = 100
MAX_SEARCH_RESULTS = 200

# Seconds between folder polls in watcher
WATCH_POLL_INTERVAL = 10


def load_settings() -> dict:
    if os.path.exists(SETTINGS_PATH):
        with open(SETTINGS_PATH) as f:
            return json.load(f)
    return {"email_folder": DEFAULT_EMAIL_FOLDER}


def save_settings(settings: dict) -> None:
    with open(SETTINGS_PATH, "w") as f:
        json.dump(settings, f, indent=2)
