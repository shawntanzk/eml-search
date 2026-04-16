"""Orchestrates FTS, semantic, and hybrid search with Reciprocal Rank Fusion."""
import sqlite3
from typing import Optional

import config
from modules import indexer, semantic_search


def _rrf_merge(
    fts_results: list[dict],
    sem_results: list[tuple[str, float]],
    k: int = 60,
) -> list[dict]:
    """Merge FTS and semantic results using Reciprocal Rank Fusion."""
    fts_ids = [r["id"] for r in fts_results]
    sem_ids = [rid for rid, _ in sem_results]

    rrf: dict[str, float] = {}
    for rank, eid in enumerate(fts_ids):
        rrf[eid] = rrf.get(eid, 0.0) + 1.0 / (k + rank + 1)
    for rank, eid in enumerate(sem_ids):
        rrf[eid] = rrf.get(eid, 0.0) + 1.0 / (k + rank + 1)

    all_ids = sorted(rrf, key=lambda x: rrf[x], reverse=True)

    fts_by_id = {r["id"]: r for r in fts_results}
    rows = []
    for eid in all_ids[: config.MAX_SEARCH_RESULTS]:
        if eid in fts_by_id:
            rows.append(fts_by_id[eid])
        else:
            em = indexer.get_email_by_id(eid)
            if em:
                rows.append(
                    {
                        "id": em["id"],
                        "subject": em["subject"],
                        "sender_name": em["sender_name"],
                        "sender_email": em["sender_email"],
                        "date": em["date"],
                        "has_attachments": em["has_attachments"],
                        "thread_id": em["thread_id"],
                        "snippet": (em.get("body_text") or "")[:200],
                    }
                )
    return rows


def search(
    query: str,
    mode: str = "hybrid",
    filters: Optional[dict] = None,
    limit: int = 50,
) -> list[dict]:
    """
    Search emails.

    Parameters
    ----------
    query   : search string
    mode    : 'fts' | 'semantic' | 'hybrid'
    filters : dict with optional keys: sender, date_from, date_to,
              has_attachments (bool), topic_id (int)
    limit   : max results to return
    """
    if filters is None:
        filters = {}

    fts_results: list[dict] = []
    sem_results: list[tuple[str, float]] = []

    if mode in ("fts", "hybrid") and query.strip():
        fts_results = indexer.search_fts(query, filters, limit=config.MAX_SEARCH_RESULTS)

    if mode in ("semantic", "hybrid") and query.strip() and semantic_search.SEMANTIC_AVAILABLE:
        ids, matrix = indexer.get_all_embeddings()
        if len(ids) > 0:
            try:
                q_vec = semantic_search.embed_text(query)
                sem_results = semantic_search.cosine_search(
                    q_vec, ids, matrix, top_k=config.SEMANTIC_TOP_K
                )
                # Apply filters post-hoc for semantic results
                if any(filters.values()):
                    filtered_ids = _apply_filters_to_ids(
                        [rid for rid, _ in sem_results], filters
                    )
                    sem_results = [(rid, s) for rid, s in sem_results if rid in filtered_ids]
            except Exception:
                pass

    # If semantic unavailable, treat hybrid/semantic as fts
    if not semantic_search.SEMANTIC_AVAILABLE and mode in ("semantic", "hybrid"):
        mode = "fts"

    if mode == "fts" or not query.strip():
        if not query.strip():
            return indexer.list_emails(filters, limit=limit)
        return fts_results[:limit]

    if mode == "semantic":
        if not sem_results:
            return []
        result_ids = [rid for rid, _ in sem_results[:limit]]
        rows = []
        for eid in result_ids:
            em = indexer.get_email_by_id(eid)
            if em:
                rows.append(
                    {
                        "id": em["id"],
                        "subject": em["subject"],
                        "sender_name": em["sender_name"],
                        "sender_email": em["sender_email"],
                        "date": em["date"],
                        "has_attachments": em["has_attachments"],
                        "thread_id": em["thread_id"],
                        "snippet": (em.get("body_text") or "")[:200],
                    }
                )
        return rows

    # Hybrid: RRF merge
    merged = _rrf_merge(fts_results, sem_results)
    return merged[:limit]


def _apply_filters_to_ids(ids: list[str], filters: dict) -> set[str]:
    if not ids:
        return set()
    conn = indexer._get_conn()
    placeholders = ",".join("?" * len(ids))
    sql = f"SELECT id FROM emails WHERE id IN ({placeholders})"
    params: list = list(ids)

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

    rows = conn.execute(sql, params).fetchall()
    return {r["id"] for r in rows}
