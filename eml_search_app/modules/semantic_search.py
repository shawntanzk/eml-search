"""Semantic search using sentence-transformers (offline after first download)."""
from typing import Optional

import numpy as np

import config

_model = None


def _load_model():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        _model = SentenceTransformer(config.SENTENCE_TRANSFORMER_MODEL)
    return _model


def embed_text(text: str) -> np.ndarray:
    model = _load_model()
    if not text or not text.strip():
        text = " "
    return model.encode(text[:512], normalize_embeddings=True, show_progress_bar=False)


def embed_batch(texts: list[str], batch_size: int = 64) -> np.ndarray:
    model = _load_model()
    cleaned = [t[:512] if t and t.strip() else " " for t in texts]
    return model.encode(
        cleaned,
        batch_size=batch_size,
        normalize_embeddings=True,
        show_progress_bar=False,
    )


def cosine_search(
    query_vec: np.ndarray,
    ids: list[str],
    matrix: np.ndarray,
    top_k: int = 50,
) -> list[tuple[str, float]]:
    """Return top_k (email_id, score) pairs sorted by cosine similarity."""
    if matrix.shape[0] == 0:
        return []
    # Embeddings are already l2-normalised, so dot product == cosine similarity
    scores = matrix @ query_vec.astype(np.float32)
    top_k = min(top_k, len(scores))
    top_indices = np.argpartition(scores, -top_k)[-top_k:]
    top_indices = top_indices[np.argsort(scores[top_indices])[::-1]]
    return [(ids[i], float(scores[i])) for i in top_indices]


def is_model_available() -> bool:
    try:
        _load_model()
        return True
    except Exception:
        return False
