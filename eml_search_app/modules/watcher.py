"""Background watcher: polls email folder for new .eml files and indexes them."""
import logging
import threading
from pathlib import Path
from typing import Optional

import config
from modules import eml_parser, indexer, nlp_engine, semantic_search

logger = logging.getLogger(__name__)


class EmailWatcher:
    """
    Runs a background daemon thread that periodically scans the email folder
    for new .eml files and runs the indexing pipeline on them.

    Pipeline per email:
      1. Parse EML → structured dict
      2. Insert into SQLite + FTS5
      3. NER via spaCy → store entities
      4. Compute sentence-transformer embedding → store as BLOB
    """

    def __init__(self, folder: str, interval: int = config.WATCH_POLL_INTERVAL):
        self.folder = folder
        self.interval = interval
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self.status = "idle"
        self.last_indexed = 0

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="EmailWatcher")
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()

    def _run(self) -> None:
        self._scan()
        while not self._stop_event.wait(self.interval):
            self._scan()

    def _scan(self) -> None:
        if not Path(self.folder).exists():
            self.status = f"folder not found: {self.folder}"
            return

        try:
            new_files = indexer.get_unindexed_files(self.folder)
        except Exception as exc:
            self.status = f"scan error: {exc}"
            return

        if not new_files:
            self.status = "watching"
            return

        self.status = f"indexing {len(new_files)} new email(s)…"
        indexed_count = 0

        for path in new_files:
            try:
                parsed = eml_parser.parse_eml(path)
                if parsed is None:
                    continue

                inserted = indexer.insert_email(parsed)
                if not inserted:
                    continue

                text = f"{parsed.get('subject', '')} {parsed.get('body_text', '')}"
                entities = nlp_engine.extract_entities(text)
                entities += nlp_engine.extract_orgs_from_email_addrs(parsed)
                if entities:
                    indexer.insert_entities(parsed["id"], entities)

                embed_text = f"{parsed.get('subject', '')} {parsed.get('body_text', '')[:400]}"
                vec = semantic_search.embed_text(embed_text)
                indexer.insert_embedding(parsed["id"], vec)

                indexed_count += 1

            except Exception as exc:
                logger.warning("Failed to index %s: %s", path, exc)

        self.last_indexed = indexed_count
        self.status = f"watching ({indexed_count} indexed this scan)"


def run_initial_index(folder: str) -> dict:
    """Synchronous one-shot indexing pass (used by setup_models.py and the UI)."""
    indexer.init_db()

    new_files = indexer.get_unindexed_files(folder)
    indexed = 0

    for path in new_files:
        try:
            parsed = eml_parser.parse_eml(path)
            if parsed is None:
                continue
            indexer.insert_email(parsed)

            text = f"{parsed.get('subject', '')} {parsed.get('body_text', '')}"
            entities = nlp_engine.extract_entities(text)
            entities += nlp_engine.extract_orgs_from_email_addrs(parsed)
            if entities:
                indexer.insert_entities(parsed["id"], entities)

            embed_text = f"{parsed.get('subject', '')} {parsed.get('body_text', '')[:400]}"
            vec = semantic_search.embed_text(embed_text)
            indexer.insert_embedding(parsed["id"], vec)

            indexed += 1
        except Exception as exc:
            logger.warning("Failed to index %s: %s", path, exc)

    return {"indexed": indexed, "total": indexer.get_email_count()}
