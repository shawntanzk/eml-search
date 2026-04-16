"""Semantic search using sentence-transformers (offline after first download).

If sentence-transformers (or its dependencies, e.g. torch) cannot be imported
— e.g. on Python 3.14+ where wheels aren't available yet — SEMANTIC_AVAILABLE
is set to False and all public functions return empty results silently.
"""
from typing import Optional

import numpy as np

import config

try:
    from sentence_transformers import SentenceTransformer as _SentenceTransformer
    SEMANTIC_AVAILABLE = True
except Exception:
    _SentenceTransformer = None  # type: ignore[assignment,misc]
    SEMANTIC_AVAILABLE = False

_model = None
_model_load_attempted = False
_model_error: str | None = None


def _load_model():
    global _model, _model_load_attempted, _model_error
    if not SEMANTIC_AVAILABLE:
        raise RuntimeError("sentence-transformers is not installed")
    if _model_load_attempted:
        if _model is None:
            raise RuntimeError(_model_error or "model failed to load")
        return _model
    _model_load_attempted = True
    try:
        # Prefer local cache to avoid SSL errors on cert-restricted machines
        try:
            _model = _SentenceTransformer(
                config.SENTENCE_TRANSFORMER_MODEL, local_files_only=True
            )
        except Exception:
            _model = _SentenceTransformer(config.SENTENCE_TRANSFORMER_MODEL)
    except Exception as exc:
        _model = None
        _model_error = f"{type(exc).__name__}: {exc}"
        raise
    return _model


def model_status() -> tuple[bool, str | None]:
    """
    Returns (ok, error_message).
    ok=True means the model loaded successfully.
    ok=False means either the package is missing or the model failed to load,
    with error_message explaining why.
    """
    if not SEMANTIC_AVAILABLE:
        return False, "sentence-transformers package is not installed"
    if not _model_load_attempted:
        try:
            _load_model()
        except Exception:
            pass
    if _model is not None:
        return True, None
    return False, _model_error


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
    if not SEMANTIC_AVAILABLE:
        return False
    try:
        _load_model()
        return True
    except Exception:
        return False
