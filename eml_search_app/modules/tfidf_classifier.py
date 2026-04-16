"""Lightweight TF-IDF classifier — no HuggingFace, no spaCy, no torch.

Used as a fallback (or explicit choice) for NLP tag auto-classification when
sentence-transformers is unavailable or undesirable.

Algorithm
---------
1. Tokenise all email bodies into term bags (lower-case alpha, length > 2,
   minus a small English stopword list).
2. Build an IDF table over the corpus.
3. For each tag name, treat its words as a tiny query document.
4. Score every email by the cosine similarity of its TF-IDF vector against
   the tag's TF-IDF vector.
5. Assign the tag to emails whose score meets the threshold.

Only numpy is required (already a core dependency of the app).
"""
import math
import re
from collections import Counter

import numpy as np

# ── Minimal English stopword list ────────────────────────────────────────────
_STOPWORDS = frozenset("""
a about above after again against all also am an and any are aren't as at be
because been before being below between both but by can't cannot could couldn't
did didn't do does doesn't doing don't down during each few for from further get
got had hadn't has hasn't have haven't having he he'd he'll he's her here here's
hers herself him himself his how how's i i'd i'll i'm i've if in into is isn't
it it's its itself let's me more most mustn't my myself no nor not of off on once
only or other ought our ours ourselves out over own same shan't she she'd she'll
she's should shouldn't so some such than that that's the their theirs them
themselves then there there's these they they'd they'll they're they've this
those through to too under until up very was wasn't we we'd we'll we're we've
were weren't what what's when when's where where's which while who who's whom
why why's will with won't would wouldn't you you'd you'll you're you've your
yours yourself yourselves also hi just like one see well
""".split())


def _tokenize(text: str) -> list[str]:
    words = re.findall(r"[a-z]+", text.lower())
    return [w for w in words if w not in _STOPWORDS and len(w) > 2]


class TFIDFClassifier:
    """Fits on a corpus of documents; scores new queries against them."""

    def __init__(self, docs: list[str], max_vocab: int = 8000):
        """
        Parameters
        ----------
        docs      : list of raw text strings (one per email)
        max_vocab : cap the vocabulary at the N most frequent terms
        """
        n = len(docs)

        # Count document frequency for each term
        df: Counter = Counter()
        tok_docs: list[list[str]] = []
        for d in docs:
            toks = _tokenize(d)
            tok_docs.append(toks)
            df.update(set(toks))

        # Keep only the top max_vocab terms (by doc frequency)
        vocab_terms = [term for term, _ in df.most_common(max_vocab)]
        self._vocab: dict[str, int] = {t: i for i, t in enumerate(vocab_terms)}
        v = len(self._vocab)

        # IDF (smoothed)
        self._idf = np.zeros(v, dtype=np.float32)
        for term, idx in self._vocab.items():
            self._idf[idx] = math.log((n + 1) / (df[term] + 1)) + 1.0

        # Build normalised TF-IDF matrix  (n_docs × vocab)
        mat = np.zeros((n, v), dtype=np.float32)
        for i, toks in enumerate(tok_docs):
            if not toks:
                continue
            tf = Counter(toks)
            total = len(toks)
            for term, cnt in tf.items():
                idx = self._vocab.get(term)
                if idx is not None:
                    mat[i, idx] = (cnt / total) * self._idf[idx]

        # L2-normalise rows
        norms = np.linalg.norm(mat, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        self._matrix = mat / norms  # shape: (n_docs, vocab)

    def score(self, query: str) -> np.ndarray:
        """Return cosine similarity of query against every document (1-D array)."""
        toks = _tokenize(query)
        if not toks:
            return np.zeros(self._matrix.shape[0], dtype=np.float32)

        v = len(self._vocab)
        qvec = np.zeros(v, dtype=np.float32)
        tf = Counter(toks)
        total = len(toks)
        for term, cnt in tf.items():
            idx = self._vocab.get(term)
            if idx is not None:
                qvec[idx] = (cnt / total) * self._idf[idx]

        norm = np.linalg.norm(qvec)
        if norm == 0:
            return np.zeros(self._matrix.shape[0], dtype=np.float32)
        qvec /= norm

        return self._matrix @ qvec  # cosine similarities
