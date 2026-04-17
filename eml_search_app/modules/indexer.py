"""SQLite + FTS5 index for emails with embeddings, NER entities, and user-defined tags."""
import json
import sqlite3
import threading
from pathlib import Path
from typing import Any, Optional

import numpy as np
import streamlit as st

import config

_local = threading.local()


def _get_conn() -> sqlite3.Connection:
    if not hasattr(_local, "conn") or _local.conn is None:
        conn = sqlite3.connect(config.DB_PATH, check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.row_factory = sqlite3.Row
        _local.conn = conn
    return _local.conn


def init_db() -> None:
    conn = _get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS emails (
            id TEXT PRIMARY KEY,
            file_path TEXT UNIQUE NOT NULL,
            message_id TEXT,
            subject TEXT,
            sender_name TEXT,
            sender_email TEXT,
            recipients TEXT,
            cc TEXT,
            date TEXT,
            body_text TEXT,
            has_attachments INTEGER DEFAULT 0,
            attachment_names TEXT,
            thread_id TEXT,
            in_reply_to TEXT,
            indexed_at TEXT DEFAULT (datetime('now'))
        );

        CREATE INDEX IF NOT EXISTS idx_emails_date ON emails(date);
        CREATE INDEX IF NOT EXISTS idx_emails_sender ON emails(sender_email);
        CREATE INDEX IF NOT EXISTS idx_emails_thread ON emails(thread_id);

        CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
            subject,
            sender_name,
            sender_email,
            body_text,
            content=emails,
            content_rowid=rowid,
            tokenize='porter unicode61'
        );

        CREATE TRIGGER IF NOT EXISTS emails_ai AFTER INSERT ON emails BEGIN
            INSERT INTO emails_fts(rowid, subject, sender_name, sender_email, body_text)
            VALUES (new.rowid, new.subject, new.sender_name, new.sender_email, new.body_text);
        END;

        CREATE TRIGGER IF NOT EXISTS emails_ad AFTER DELETE ON emails BEGIN
            INSERT INTO emails_fts(emails_fts, rowid, subject, sender_name, sender_email, body_text)
            VALUES ('delete', old.rowid, old.subject, old.sender_name, old.sender_email, old.body_text);
        END;

        CREATE TABLE IF NOT EXISTS email_entities (
            email_id TEXT NOT NULL,
            entity_text TEXT NOT NULL,
            entity_label TEXT NOT NULL,
            PRIMARY KEY (email_id, entity_text, entity_label)
        );
        CREATE INDEX IF NOT EXISTS idx_entities_label ON email_entities(entity_label);

        CREATE TABLE IF NOT EXISTS embeddings (
            email_id TEXT PRIMARY KEY,
            vector BLOB NOT NULL
        );

        -- Human-defined tags (the tag library)
        CREATE TABLE IF NOT EXISTS tags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        );

        -- Active tag assignments: source is 'manual' or 'nlp'
        CREATE TABLE IF NOT EXISTS email_tags (
            email_id TEXT NOT NULL,
            tag_id INTEGER NOT NULL,
            source TEXT NOT NULL DEFAULT 'nlp',
            PRIMARY KEY (email_id, tag_id)
        );
        CREATE INDEX IF NOT EXISTS idx_email_tags_tag ON email_tags(tag_id);

        -- NLP is blocked from assigning this tag to this email
        -- (set when a user manually removes a tag)
        CREATE TABLE IF NOT EXISTS email_tag_blocks (
            email_id TEXT NOT NULL,
            tag_id INTEGER NOT NULL,
            PRIMARY KEY (email_id, tag_id)
        );

        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT
        );
    """)
    # Idempotent schema migrations — add per-tag NLP settings columns if absent
    for col, definition in [
        ("nlp_method", "TEXT NOT NULL DEFAULT 'tfidf'"),
        ("nlp_threshold", "REAL NOT NULL DEFAULT 0.15"),
    ]:
        try:
            conn.execute(f"ALTER TABLE tags ADD COLUMN {col} {definition}")
            conn.commit()
        except Exception:
            pass  # column already exists


def is_indexed(file_path: str) -> bool:
    conn = _get_conn()
    row = conn.execute(
        "SELECT 1 FROM emails WHERE file_path = ?", (str(Path(file_path).resolve()),)
    ).fetchone()
    return row is not None


def insert_email(parsed: dict) -> bool:
    """Insert a parsed email. Returns True if inserted, False if already exists."""
    conn = _get_conn()
    try:
        conn.execute(
            """INSERT OR IGNORE INTO emails
               (id, file_path, message_id, subject, sender_name, sender_email,
                recipients, cc, date, body_text, has_attachments, attachment_names,
                thread_id, in_reply_to)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                parsed["id"],
                parsed["file_path"],
                parsed.get("message_id", ""),
                parsed.get("subject", ""),
                parsed.get("sender_name", ""),
                parsed.get("sender_email", ""),
                json.dumps(parsed.get("recipients", [])),
                json.dumps(parsed.get("cc", [])),
                parsed.get("date"),
                parsed.get("body_text", ""),
                1 if parsed.get("has_attachments") else 0,
                json.dumps(parsed.get("attachment_names", [])),
                parsed.get("thread_id", ""),
                parsed.get("in_reply_to", ""),
            ),
        )
        conn.commit()
        return conn.execute("SELECT changes()").fetchone()[0] > 0
    except sqlite3.IntegrityError:
        return False


def insert_entities(email_id: str, entities: list[dict]) -> None:
    conn = _get_conn()
    conn.executemany(
        "INSERT OR IGNORE INTO email_entities (email_id, entity_text, entity_label) VALUES (?,?,?)",
        [(email_id, e["text"], e["label"]) for e in entities],
    )
    conn.commit()


def insert_embedding(email_id: str, vector: np.ndarray) -> None:
    conn = _get_conn()
    conn.execute(
        "INSERT OR REPLACE INTO embeddings (email_id, vector) VALUES (?,?)",
        (email_id, vector.astype(np.float32).tobytes()),
    )
    conn.commit()


def get_meta(key: str) -> Optional[str]:
    conn = _get_conn()
    row = conn.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else None


def set_meta(key: str, value: str) -> None:
    conn = _get_conn()
    conn.execute("INSERT OR REPLACE INTO meta (key, value) VALUES (?,?)", (key, value))
    conn.commit()


def get_email_count() -> int:
    return _get_conn().execute("SELECT COUNT(*) FROM emails").fetchone()[0]


def get_embedding_count() -> int:
    return _get_conn().execute("SELECT COUNT(*) FROM embeddings").fetchone()[0]


def get_all_email_ids() -> list[str]:
    rows = _get_conn().execute("SELECT id FROM emails").fetchall()
    return [r["id"] for r in rows]


def get_emails_without_embeddings() -> list[dict]:
    conn = _get_conn()
    rows = conn.execute(
        "SELECT id, subject, body_text FROM emails WHERE id NOT IN (SELECT email_id FROM embeddings)"
    ).fetchall()
    return [dict(r) for r in rows]


def get_emails_without_entities() -> list[dict]:
    conn = _get_conn()
    rows = conn.execute(
        "SELECT id, body_text, subject FROM emails WHERE id NOT IN (SELECT DISTINCT email_id FROM email_entities)"
    ).fetchall()
    return [dict(r) for r in rows]


def get_all_embeddings() -> tuple[list[str], np.ndarray]:
    """Returns (list_of_email_ids, embedding_matrix)."""
    conn = _get_conn()
    rows = conn.execute("SELECT email_id, vector FROM embeddings").fetchall()
    if not rows:
        return [], np.empty((0, 0), dtype=np.float32)
    n = len(rows)
    dim = len(rows[0]["vector"]) // 4  # float32 = 4 bytes
    matrix = np.empty((n, dim), dtype=np.float32)
    ids = []
    for i, r in enumerate(rows):
        ids.append(r["email_id"])
        matrix[i] = np.frombuffer(r["vector"], dtype=np.float32)
    return ids, matrix


@st.cache_data
def get_cached_embeddings() -> tuple[list[str], np.ndarray]:
    """Cached version of get_all_embeddings for search operations."""
    return get_all_embeddings()


def search_fts(query: str, filters: dict, limit: int = 200) -> list[dict]:
    conn = _get_conn()
    params: list[Any] = []

    fts_query = query.replace('"', '""')
    sql = """
        SELECT e.id, e.subject, e.sender_name, e.sender_email, e.date,
               e.has_attachments, e.thread_id,
               snippet(emails_fts, 3, '<b>', '</b>', '...', 20) AS snippet
        FROM emails_fts
        JOIN emails e ON e.rowid = emails_fts.rowid
        WHERE emails_fts MATCH ?
    """
    params.append(f'"{fts_query}"')

    if filters.get("sender"):
        sql += " AND e.sender_email LIKE ?"
        params.append(f"%{filters['sender'].lower()}%")
    if filters.get("date_from"):
        sql += " AND e.date >= ?"
        params.append(filters["date_from"])
    if filters.get("date_to"):
        sql += " AND e.date <= ?"
        params.append(filters["date_to"])
    if filters.get("has_attachments"):
        sql += " AND e.has_attachments = 1"
    if filters.get("tag_id") is not None:
        sql += " AND e.id IN (SELECT email_id FROM email_tags WHERE tag_id = ?)"
        params.append(filters["tag_id"])

    sql += f" ORDER BY rank LIMIT {limit}"
    try:
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    except sqlite3.OperationalError:
        return []


def get_email_by_id(email_id: str) -> Optional[dict]:
    conn = _get_conn()
    row = conn.execute("SELECT * FROM emails WHERE id = ?", (email_id,)).fetchone()
    if not row:
        return None
    result = dict(row)
    result["recipients"] = json.loads(result.get("recipients") or "[]")
    result["cc"] = json.loads(result.get("cc") or "[]")
    result["attachment_names"] = json.loads(result.get("attachment_names") or "[]")
    return result


def get_email_keywords(email_id: str, limit: int = 10) -> list[str]:
    """Return stored entity texts for an email — no model inference needed."""
    rows = _get_conn().execute(
        "SELECT entity_text FROM email_entities WHERE email_id = ? "
        "AND entity_label IN ('PERSON','ORG','GPE','PRODUCT','LOC') LIMIT ?",
        (email_id, limit),
    ).fetchall()
    return [r["entity_text"] for r in rows]


def get_emails_by_ids(email_ids: list[str]) -> dict[str, dict]:
    """Fetch multiple emails by ID in a single query. Returns {id: email_dict}."""
    if not email_ids:
        return {}
    conn = _get_conn()
    placeholders = ",".join("?" * len(email_ids))
    rows = conn.execute(
        f"SELECT id, subject, sender_name, sender_email, date, has_attachments, thread_id, body_text"
        f" FROM emails WHERE id IN ({placeholders})",
        email_ids,
    ).fetchall()
    return {r["id"]: dict(r) for r in rows}


def list_emails(filters: dict, limit: int = 50, offset: int = 0) -> list[dict]:
    conn = _get_conn()
    sql = "SELECT id, subject, sender_name, sender_email, date, has_attachments, thread_id FROM emails WHERE 1=1"
    params: list[Any] = []

    if filters.get("sender"):
        sql += " AND sender_email LIKE ?"
        params.append(f"%{filters['sender'].lower()}%")
    if filters.get("date_from"):
        sql += " AND date >= ?"
        params.append(filters["date_from"])
    if filters.get("date_to"):
        sql += " AND date <= ?"
        params.append(filters["date_to"])
    if filters.get("has_attachments"):
        sql += " AND has_attachments = 1"
    if filters.get("tag_id") is not None:
        sql += " AND id IN (SELECT email_id FROM email_tags WHERE tag_id = ?)"
        params.append(filters["tag_id"])

    sql += f" ORDER BY date DESC LIMIT {limit} OFFSET {offset}"
    rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def delete_emails_by_paths(paths: list[str]) -> int:
    """
    Delete emails and all related data (entities, embeddings, tags) by file_path.
    The FTS5 delete trigger handles emails_fts automatically.
    Returns the number of emails deleted.
    """
    if not paths:
        return 0
    conn = _get_conn()
    placeholders = ",".join("?" * len(paths))
    rows = conn.execute(
        f"SELECT id FROM emails WHERE file_path IN ({placeholders})", paths
    ).fetchall()
    email_ids = [r["id"] for r in rows]
    if not email_ids:
        return 0
    id_ph = ",".join("?" * len(email_ids))
    conn.execute(f"DELETE FROM email_entities  WHERE email_id IN ({id_ph})", email_ids)
    conn.execute(f"DELETE FROM embeddings      WHERE email_id IN ({id_ph})", email_ids)
    conn.execute(f"DELETE FROM email_tags      WHERE email_id IN ({id_ph})", email_ids)
    conn.execute(f"DELETE FROM email_tag_blocks WHERE email_id IN ({id_ph})", email_ids)
    conn.execute(f"DELETE FROM emails          WHERE id        IN ({id_ph})", email_ids)
    conn.commit()
    return len(email_ids)


def get_indexed_imap_paths(host: str, mailbox: str) -> set[str]:
    """Return the set of imap:// virtual file_paths already indexed for a given host+mailbox."""
    conn = _get_conn()
    prefix = f"imap://{host}/{mailbox}/"
    rows = conn.execute(
        "SELECT file_path FROM emails WHERE file_path LIKE ?", (prefix + "%",)
    ).fetchall()
    return {r["file_path"] for r in rows}


def get_unindexed_files(folder: str) -> list[str]:
    conn = _get_conn()
    indexed = set(r[0] for r in conn.execute("SELECT file_path FROM emails").fetchall())
    return [
        str(p.resolve())
        for p in Path(folder).rglob("*.eml")
        if str(p.resolve()) not in indexed
    ]


def get_tag_nlp_settings(tag_id: int) -> dict:
    """Return the nlp_method and nlp_threshold for a tag (with defaults)."""
    conn = _get_conn()
    row = conn.execute(
        "SELECT nlp_method, nlp_threshold FROM tags WHERE id = ?", (tag_id,)
    ).fetchone()
    if row:
        return {"nlp_method": row["nlp_method"], "nlp_threshold": row["nlp_threshold"]}
    return {"nlp_method": "tfidf", "nlp_threshold": 0.15}


def save_tag_nlp_settings(tag_id: int, method: str, threshold: float) -> None:
    """Persist per-tag NLP classification method and threshold."""
    conn = _get_conn()
    conn.execute(
        "UPDATE tags SET nlp_method = ?, nlp_threshold = ? WHERE id = ?",
        (method, threshold, tag_id),
    )
    conn.commit()
