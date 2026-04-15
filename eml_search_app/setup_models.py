"""
One-time setup script. Run this before launching the app:

    python setup_models.py [--folder /path/to/eml/folder]

It will:
  1. Download the spaCy model  (en_core_web_sm, ~12 MB)
  2. Download the sentence-transformer model  (all-MiniLM-L6-v2, ~90 MB)
  3. Initialise the SQLite database
  4. Index all .eml files in the configured folder
"""
import argparse
import subprocess
import sys
from pathlib import Path

# Ensure the app directory is on the path
sys.path.insert(0, str(Path(__file__).parent))

import config


def download_spacy_model() -> None:
    print(f"[1/4] Downloading spaCy model '{config.SPACY_MODEL}'…")
    try:
        import spacy
        spacy.load(config.SPACY_MODEL)
        print("      already installed.")
    except OSError:
        subprocess.run(
            [sys.executable, "-m", "spacy", "download", config.SPACY_MODEL],
            check=True,
        )


def download_sentence_transformer() -> None:
    print(f"[2/4] Loading sentence-transformer '{config.SENTENCE_TRANSFORMER_MODEL}'…")
    from sentence_transformers import SentenceTransformer
    SentenceTransformer(config.SENTENCE_TRANSFORMER_MODEL)
    print("      model cached.")


def init_database() -> None:
    print("[3/4] Initialising database…")
    from modules.indexer import init_db
    init_db()
    print("      done.")


def index_emails(folder: str) -> None:
    print(f"[4/4] Indexing emails in '{folder}'…")
    from modules.watcher import run_initial_index
    result = run_initial_index(folder)
    print(f"      indexed {result['indexed']} new emails ({result['total']} total in index).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Set up the EML search app.")
    parser.add_argument(
        "--folder",
        default=None,
        help="Path to the folder containing .eml files (overrides config/settings).",
    )
    args = parser.parse_args()

    raw = args.folder or config.load_settings().get("email_folder", config.DEFAULT_EMAIL_FOLDER)
    folder = str(Path(raw).resolve())

    if not Path(folder).exists():
        print(f"Warning: folder '{folder}' does not exist. Skipping email indexing.")
        folder = None

    download_spacy_model()
    download_sentence_transformer()
    init_database()

    if folder:
        # Save the folder choice
        settings = config.load_settings()
        settings["email_folder"] = folder
        config.save_settings(settings)
        index_emails(folder)
    else:
        print("[4/4] Skipped (no valid folder).")

    print("\nSetup complete. Start the app with:  streamlit run app.py")
