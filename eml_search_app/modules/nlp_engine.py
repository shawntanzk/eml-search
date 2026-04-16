"""NLP processing: NER via spaCy and keyword extraction.

If the spaCy model is not installed, NLP_AVAILABLE is set to False and all
public functions return empty results silently rather than raising errors.

extract_orgs_from_email_addrs() is always available — it requires no models.
"""
import re
import warnings
from collections import Counter

import config

# ── Email-based organisation extraction ─────────────────────────────────────

# Domains that belong to free/personal email providers — not organisations.
_FREE_EMAIL_DOMAINS = frozenset("""
gmail.com yahoo.com yahoo.co.uk yahoo.com.au hotmail.com hotmail.co.uk
hotmail.fr outlook.com outlook.co.uk live.com live.co.uk icloud.com
me.com mac.com aol.com aol.co.uk protonmail.com proton.me tutanota.com
tutanota.de fastmail.com fastmail.net zoho.com ymail.com msn.com
googlemail.com mail.com inbox.com gmx.com gmx.net gmx.de web.de
""".split())

# Country-code second-level domains used before the TLD (e.g. co.uk, com.au)
_SECOND_LEVEL = frozenset(["co", "com", "org", "net", "gov", "edu", "ac", "sch"])


def _domain_to_org_name(domain: str) -> str | None:
    """
    Convert an email domain to a human-readable organisation name.
    Returns None for free providers, IP addresses, or single-label domains.

    Examples
    --------
    microsoft.com      → Microsoft
    acme-corp.co.uk    → Acme Corp
    sub.bigbank.com    → Bigbank
    gmail.com          → None
    """
    domain = domain.lower().strip()
    if not domain or domain in _FREE_EMAIL_DOMAINS:
        return None

    parts = domain.split(".")
    if len(parts) < 2:
        return None

    # Determine the registered second-level domain
    # e.g. acme.co.uk → parts = ["acme","co","uk"] → sld = "acme"
    if len(parts) >= 3 and parts[-2] in _SECOND_LEVEL:
        sld = parts[-3]
    else:
        sld = parts[-2]

    name = sld.replace("-", " ").replace("_", " ").title()
    return name if len(name) > 1 else None


def extract_orgs_from_email_addrs(parsed: dict) -> list[dict]:
    """
    Extract organisation entities from the sender and recipient email addresses
    in a parsed email dict. Always available — no models required.

    Collects all unique addresses, strips free-provider domains, and returns
    a list of {text, label} dicts (label = 'ORG') ready for indexer.insert_entities().
    """
    addresses: set[str] = set()

    sender = parsed.get("sender_email", "")
    if sender:
        addresses.add(sender.lower())

    for field in ("recipients", "cc"):
        for entry in (parsed.get(field) or []):
            addr = entry.get("email", "") if isinstance(entry, dict) else ""
            if addr:
                addresses.add(addr.lower())

    seen_orgs: set[str] = set()
    entities: list[dict] = []
    for addr in addresses:
        if "@" not in addr:
            continue
        domain = addr.split("@", 1)[1]
        name = _domain_to_org_name(domain)
        if name and name not in seen_orgs:
            seen_orgs.add(name)
            entities.append({"text": name, "label": "ORG"})

    return entities

warnings.filterwarnings("ignore", category=UserWarning)

_nlp = None
_load_attempted = False
_load_error: str | None = None


def _load_spacy():
    global _nlp, _load_attempted, _load_error
    if _load_attempted:
        return _nlp  # None means unavailable — don't retry
    _load_attempted = True
    try:
        import spacy
        _nlp = spacy.load(config.SPACY_MODEL, disable=["parser"])
    except Exception as exc:
        _nlp = None
        _load_error = f"{type(exc).__name__}: {exc}"
    return _nlp


def NLP_AVAILABLE() -> bool:
    return _load_spacy() is not None


def NLP_ERROR() -> str | None:
    """Return the exception message from the last failed spaCy load, or None if loaded OK."""
    _load_spacy()  # ensure load has been attempted
    return _load_error


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
