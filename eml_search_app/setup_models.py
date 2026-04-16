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
import tarfile
import tempfile
import urllib.request
from pathlib import Path

# Ensure the app directory is on the path
sys.path.insert(0, str(Path(__file__).parent))

import config

# Fallback models hosted on GitHub Releases (bypasses cert-restricted networks)
SPACY_MODEL_FALLBACK_URL = (
    "https://github.com/shawntanzk/eml-search/releases/download/"
    "models-v1/en_core_web_sm-3.8.0.tar.gz"
)
SENTENCE_TRANSFORMER_FALLBACK_URL = (
    "https://github.com/shawntanzk/eml-search/releases/download/"
    "models-v1/all-MiniLM-L6-v2.tar.gz"
)


def _download_tarball(url: str, dest: Path) -> None:
    """Download a tarball from url and extract it into dest."""
    print(f"      downloading from {url} …")
    with tempfile.NamedTemporaryFile(suffix=".tar.gz", delete=False) as tmp:
        urllib.request.urlretrieve(url, tmp.name)
        with tarfile.open(tmp.name, "r:gz") as tf:
            tf.extractall(dest)
    print(f"      installed to {dest}")


def _install_spacy_model_from_url(url: str) -> None:
    """Download a packaged spaCy model tarball and install it into site-packages."""
    import site
    _download_tarball(url, Path(site.getsitepackages()[0]))


def download_spacy_model() -> None:
    print(f"[1/4] Downloading spaCy model '{config.SPACY_MODEL}'…")
    try:
        import spacy
        spacy.load(config.SPACY_MODEL)
        print("      already installed.")
        return
    except OSError:
        pass

    # Try the standard spaCy download first
    result = subprocess.run(
        [sys.executable, "-m", "spacy", "download", config.SPACY_MODEL],
        capture_output=True,
    )
    if result.returncode == 0:
        return

    # Fall back to the GitHub Release bundle (works on cert-restricted machines)
    print("      spaCy download failed — trying GitHub Release fallback…")
    _install_spacy_model_from_url(SPACY_MODEL_FALLBACK_URL)


def download_sentence_transformer() -> None:
    print(f"[2/4] Loading sentence-transformer '{config.SENTENCE_TRANSFORMER_MODEL}'…")
    try:
        from sentence_transformers import SentenceTransformer
    except ImportError:
        print(
            "      sentence-transformers is not installed (no compatible wheels for this "
            "Python version). Semantic search and NLP tag classification will be disabled.\n"
            f"      To install manually later, try:\n"
            f"        pip install sentence-transformers\n"
            f"      or download the pre-packaged model from:\n"
            f"        {SENTENCE_TRANSFORMER_FALLBACK_URL}"
        )
        return

    try:
        SentenceTransformer(config.SENTENCE_TRANSFORMER_MODEL)
        print("      model cached.")
        return
    except Exception:
        pass

    # Fall back to the GitHub Release bundle (works on cert-restricted machines)
    print("      HuggingFace download failed — trying GitHub Release fallback…")
    hf_cache = Path.home() / ".cache" / "huggingface" / "hub"
    hf_cache.mkdir(parents=True, exist_ok=True)
    _download_tarball(SENTENCE_TRANSFORMER_FALLBACK_URL, hf_cache)
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
