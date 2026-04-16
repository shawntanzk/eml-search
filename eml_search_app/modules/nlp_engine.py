"""NLP processing: NER via spaCy and keyword extraction.

If the spaCy model is not installed, NLP_AVAILABLE is set to False and all
public functions return empty results silently rather than raising errors.
"""
import re
import warnings
from collections import Counter

import config

warnings.filterwarnings("ignore", category=UserWarning)

_nlp = None
_load_attempted = False


def _load_spacy():
    global _nlp, _load_attempted
    if _load_attempted:
        return _nlp  # None means unavailable — don't retry
    _load_attempted = True
    try:
        import spacy
        _nlp = spacy.load(config.SPACY_MODEL, disable=["parser"])
    except Exception:
        _nlp = None
    return _nlp


def NLP_AVAILABLE() -> bool:
    return _load_spacy() is not None


def extract_entities(text: str) -> list[dict]:
    """Run NER on text and return list of {text, label} dicts."""
    if not text or not text.strip():
        return []
    nlp = _load_spacy()
    if nlp is None:
        return []
    doc = nlp(text[:100_000])
    seen = set()
    entities = []
    for ent in doc.ents:
        if ent.label_ in ("PERSON", "ORG", "GPE", "LOC", "PRODUCT", "EVENT", "WORK_OF_ART"):
            key = (ent.text.strip(), ent.label_)
            if key not in seen and len(ent.text.strip()) > 1:
                seen.add(key)
                entities.append({"text": ent.text.strip(), "label": ent.label_})
    return entities


def extract_keywords(text: str, top_n: int = 10) -> list[str]:
    """Extract keywords using spaCy noun chunks and content word frequency."""
    if not text or not text.strip():
        return []
    nlp = _load_spacy()
    if nlp is None:
        return []
    doc = nlp(text[:5000])

    candidates: list[str] = []

    # Individual content words (nouns and proper nouns)
    for token in doc:
        if (
            not token.is_stop
            and not token.is_punct
            and not token.is_space
            and token.pos_ in ("NOUN", "PROPN")
            and len(token.text) > 2
        ):
            candidates.append(token.lemma_.lower())

    counts = Counter(candidates)
    return [w for w, _ in counts.most_common(top_n)]
