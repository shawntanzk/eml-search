"""Human-defined tag management + NLP auto-classification.

Tag assignment rules
--------------------
- Manual add   : always persists; clears any NLP block on that tag for that email.
- Manual remove: removes the assignment AND blocks NLP from ever re-adding that tag
                 to that email in future classification runs.
- NLP classify : only adds tags; never removes; never touches manually-blocked pairs.
                 Uses sentence-transformer cosine similarity between tag name and
                 email embedding.
"""
import sqlite3
from collections import defaultdict
from typing import Optional

import numpy as np

from modules import indexer, semantic_search
from modules.tfidf_classifier import TFIDFClassifier


# ── Tag library ──────────────────────────────────────────────────────────────

def get_all_tags() -> list[dict]:
    """Return all defined tags sorted alphabetically, including per-tag NLP settings."""
    conn = indexer._get_conn()
    rows = conn.execute(
        "SELECT id, name, nlp_method, nlp_threshold FROM tags ORDER BY name"
    ).fetchall()
    return [dict(r) for r in rows]


def add_tag(name: str) -> Optional[int]:
    """Create a new tag. Returns its id, or existing id if name already exists."""
    name = name.strip()
    if not name:
        return None
    conn = indexer._get_conn()
    existing = conn.execute("SELECT id FROM tags WHERE name = ?", (name,)).fetchone()
    if existing:
        return existing["id"]
    conn.execute("INSERT INTO tags (name) VALUES (?)", (name,))
    conn.commit()
    return conn.execute("SELECT last_insert_rowid()").fetchone()[0]


def delete_tag(tag_id: int) -> None:
    """Delete a tag and all its assignments and blocks."""
    conn = indexer._get_conn()
    conn.execute("DELETE FROM tags WHERE id = ?", (tag_id,))
    conn.execute("DELETE FROM email_tags WHERE tag_id = ?", (tag_id,))
    conn.execute("DELETE FROM email_tag_blocks WHERE tag_id = ?", (tag_id,))
    conn.commit()


# ── Per-email tag operations ──────────────────────────────────────────────────

def get_email_tags(email_id: str) -> list[dict]:
    """Return tags currently assigned to an email, with their source."""
    conn = indexer._get_conn()
    rows = conn.execute(
        """SELECT t.id, t.name, et.source
           FROM email_tags et
           JOIN tags t ON t.id = et.tag_id
           WHERE et.email_id = ?
           ORDER BY t.name""",
        (email_id,),
    ).fetchall()
    return [dict(r) for r in rows]


def assign_tag_manual(email_id: str, tag_id: int) -> None:
    """
    Manually assign a tag to an email.
    - Upgrades source to 'manual' if already assigned by NLP.
    - Removes any existing NLP block so the tag stays permanently.
    """
    conn = indexer._get_conn()
    # Clear any block the user previously set for this pair
    conn.execute(
        "DELETE FROM email_tag_blocks WHERE email_id = ? AND tag_id = ?",
        (email_id, tag_id),
    )
    # Insert or upgrade source to 'manual'
    conn.execute(
        """INSERT INTO email_tags (email_id, tag_id, source) VALUES (?, ?, 'manual')
           ON CONFLICT(email_id, tag_id) DO UPDATE SET source = 'manual'""",
        (email_id, tag_id),
    )
    conn.commit()


def remove_tag_manual(email_id: str, tag_id: int) -> None:
    """
    Manually remove a tag from an email.
    - Deletes the assignment regardless of source.
    - Blocks NLP from ever re-adding this tag to this email.
    """
    conn = indexer._get_conn()
    conn.execute(
        "DELETE FROM email_tags WHERE email_id = ? AND tag_id = ?",
        (email_id, tag_id),
    )
    conn.execute(
        "INSERT OR IGNORE INTO email_tag_blocks (email_id, tag_id) VALUES (?, ?)",
        (email_id, tag_id),
    )
    conn.commit()


# ── NLP classification ────────────────────────────────────────────────────────

def classify_emails_nlp(threshold: float = 0.25, tags: Optional[list] = None) -> dict:
    """
    Auto-assign tags to emails using sentence-transformer cosine similarity.

    If `tags` is provided, only those tags are classified; otherwise all tags.
    Assigns the tag (source='nlp') when similarity >= threshold, UNLESS:
      - the email already has that tag (any source), OR
      - a manual block exists for that (email, tag) pair.

    Never removes any existing tag assignments.

    Returns a summary dict with 'new_assignments' and 'emails_affected'.
    """
    if not semantic_search.SEMANTIC_AVAILABLE:
        return {"new_assignments": 0, "emails_affected": 0, "unavailable": True}

    if tags is None:
        tags = get_all_tags()
    if not tags:
        return {"new_assignments": 0, "emails_affected": 0}

    email_ids, matrix = indexer.get_all_embeddings()
    if len(email_ids) == 0:
        return {"new_assignments": 0, "emails_affected": 0}

    # Embed all tag names in one batch
    tag_names = [t["name"] for t in tags]
    tag_vecs = semantic_search.embed_batch(tag_names)

    conn = indexer._get_conn()

    # Load existing assignments into a set for O(1) lookup
    existing: set[tuple[str, int]] = set(
        (r[0], r[1])
        for r in conn.execute("SELECT email_id, tag_id FROM email_tags").fetchall()
    )
    # Load blocks
    blocked: set[tuple[str, int]] = set(
        (r[0], r[1])
        for r in conn.execute("SELECT email_id, tag_id FROM email_tag_blocks").fetchall()
    )

    new_rows: list[tuple[str, int, str]] = []
    affected_emails: set[str] = set()

    for tag, tag_vec in zip(tags, tag_vecs):
        # Cosine similarity — embeddings are already l2-normalised
        scores: np.ndarray = matrix @ tag_vec.astype(np.float32)

        for email_id, score in zip(email_ids, scores):
            if float(score) < threshold:
                continue
            key = (email_id, tag["id"])
            if key in existing or key in blocked:
                continue
            new_rows.append((email_id, tag["id"], "nlp"))
            existing.add(key)  # prevent duplicates within this run
            affected_emails.add(email_id)

    if new_rows:
        conn.executemany(
            "INSERT OR IGNORE INTO email_tags (email_id, tag_id, source) VALUES (?,?,?)",
            new_rows,
        )
        conn.commit()

    return {
        "new_assignments": len(new_rows),
        "emails_affected": len(affected_emails),
    }


def classify_emails_tfidf(threshold: float = 0.15, tags: Optional[list] = None) -> dict:
    """
    Auto-assign tags using TF-IDF cosine similarity — no HuggingFace, no spaCy.

    If `tags` is provided, only those tags are classified; otherwise all tags.
    Assigns the tag (source='nlp') when similarity >= threshold, subject to
    the same manual-block rules as classify_emails_nlp.

    Returns a summary dict with 'new_assignments' and 'emails_affected'.
    """
    if tags is None:
        tags = get_all_tags()
    if not tags:
        return {"new_assignments": 0, "emails_affected": 0}

    conn = indexer._get_conn()
    rows = conn.execute(
        "SELECT id, body_text FROM emails WHERE body_text IS NOT NULL AND body_text != ''"
    ).fetchall()
    if not rows:
        return {"new_assignments": 0, "emails_affected": 0}

    email_ids = [r["id"] for r in rows]
    bodies = [r["body_text"] for r in rows]

    classifier = TFIDFClassifier(bodies)

    existing: set[tuple[str, int]] = set(
        (r[0], r[1])
        for r in conn.execute("SELECT email_id, tag_id FROM email_tags").fetchall()
    )
    blocked: set[tuple[str, int]] = set(
        (r[0], r[1])
        for r in conn.execute("SELECT email_id, tag_id FROM email_tag_blocks").fetchall()
    )

    new_rows: list[tuple[str, int, str]] = []
    affected_emails: set[str] = set()

    for tag in tags:
        scores = classifier.score(tag["name"])
        for email_id, score in zip(email_ids, scores):
            if float(score) < threshold:
                continue
            key = (email_id, tag["id"])
            if key in existing or key in blocked:
                continue
            new_rows.append((email_id, tag["id"], "nlp"))
            existing.add(key)
            affected_emails.add(email_id)

    if new_rows:
        conn.executemany(
            "INSERT OR IGNORE INTO email_tags (email_id, tag_id, source) VALUES (?,?,?)",
            new_rows,
        )
        conn.commit()

    return {
        "new_assignments": len(new_rows),
        "emails_affected": len(affected_emails),
    }


def classify_tag(tag_id: int) -> dict:
    """
    Classify emails for a single tag using that tag's saved nlp_method and nlp_threshold.

    Returns the same summary dict as classify_emails_nlp / classify_emails_tfidf.
    """
    conn = indexer._get_conn()
    row = conn.execute(
        "SELECT id, name, nlp_method, nlp_threshold FROM tags WHERE id = ?", (tag_id,)
    ).fetchone()
    if not row:
        return {"new_assignments": 0, "emails_affected": 0}

    tag = dict(row)
    method = tag.get("nlp_method") or "tfidf"
    threshold = tag.get("nlp_threshold") or 0.15

    if method == "semantic":
        return classify_emails_nlp(threshold=threshold, tags=[tag])
    else:
        return classify_emails_tfidf(threshold=threshold, tags=[tag])


def classify_all_tags() -> dict:
    """
    Classify all tags, each using its own saved nlp_method and nlp_threshold.

    Returns a combined summary dict.
    """
    tags = get_all_tags()
    total_new = 0
    total_affected: set[str] = set()

    # Partition tags by method so we can batch TF-IDF (builds corpus once)
    semantic_tags = [t for t in tags if (t.get("nlp_method") or "tfidf") == "semantic"]
    tfidf_tags = [t for t in tags if (t.get("nlp_method") or "tfidf") != "semantic"]

    # Group semantic tags by threshold so each group uses a single classify call
    semantic_by_threshold: dict[float, list] = defaultdict(list)
    for t in semantic_tags:
        semantic_by_threshold[t["nlp_threshold"]].append(t)

    tfidf_by_threshold: dict[float, list] = defaultdict(list)
    for t in tfidf_tags:
        tfidf_by_threshold[t["nlp_threshold"]].append(t)

    for threshold, group in semantic_by_threshold.items():
        r = classify_emails_nlp(threshold=threshold, tags=group)
        total_new += r.get("new_assignments", 0)

    for threshold, group in tfidf_by_threshold.items():
        r = classify_emails_tfidf(threshold=threshold, tags=group)
        total_new += r.get("new_assignments", 0)

    return {"new_assignments": total_new}


def get_emails_by_tag(tag_id: int, limit: int = 200) -> list[dict]:
    """Return emails that have a given tag assigned."""
    conn = indexer._get_conn()
    rows = conn.execute(
        """SELECT e.id, e.subject, e.sender_name, e.sender_email, e.date,
                  e.has_attachments, et.source
           FROM email_tags et
           JOIN emails e ON e.id = et.email_id
           WHERE et.tag_id = ?
           ORDER BY e.date DESC
           LIMIT ?""",
        (tag_id, limit),
    ).fetchall()
    return [dict(r) for r in rows]


def get_tag_counts() -> list[dict]:
    """Return tags with how many emails each has."""
    conn = indexer._get_conn()
    rows = conn.execute(
        """SELECT t.id, t.name, COUNT(et.email_id) AS count
           FROM tags t
           LEFT JOIN email_tags et ON et.tag_id = t.id
           GROUP BY t.id
           ORDER BY t.name""",
    ).fetchall()
    return [dict(r) for r in rows]
