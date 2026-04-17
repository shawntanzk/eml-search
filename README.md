# EML Search

A fully offline Streamlit app for indexing, searching, and analysing emails. Pull from a local folder of `.eml` backup files **or connect directly to an IMAP server** — no exports needed. No cloud services, no LLMs — everything runs on your machine.

---

## Features

| Capability | Detail |
|---|---|
| **Full-text search** | SQLite FTS5 with Porter stemming — fast queries across thousands of emails |
| **Semantic search** | Sentence-transformer embeddings (all-MiniLM-L6-v2) with cosine similarity |
| **Hybrid search** | Reciprocal Rank Fusion (RRF) merges FTS and semantic rankings |
| **Filters** | Sender, date range, has-attachments, tag |
| **Local folder indexing** | Background thread watches a folder and indexes new `.eml` files as they arrive |
| **IMAP ingestion** | Connect to any IMAP server via the Settings UI — no `.eml` exports needed |
| **Named entity recognition** | Extracts people, organisations, locations via spaCy |
| **Keyword extraction** | Per-email keyword panel using spaCy noun chunks |
| **Tag management** | Human-defined tag library with manual assignment and NLP auto-classification |
| **NLP auto-classification** | Assign tags via **Semantic** (sentence-transformers) or **TF-IDF** (no ML needed) |
| **Knowledge graph** | RDF/OWL graph with SPARQL console and interactive pyvis visualisation |
| **Graceful degradation** | Runs with FTS-only search even if spaCy or sentence-transformers are unavailable |

---

## First-Time Setup

### Step 1 — Clone and install

```bash
git clone https://github.com/shawntanzk/eml-search.git
cd "eml-search/eml_search_app"

python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

### Step 2 — Download NLP models

This is a one-time download (~100 MB total). It also initialises the database.

```bash
python setup_models.py
```

You should see:

```
Downloading spaCy model...       ✓
Downloading sentence-transformer... ✓
Database initialised.
```

If you have a local folder of `.eml` files and want to index them right away:

```bash
python setup_models.py --folder /path/to/your/eml/folder
```

### Step 3 — Start the app

```bash
streamlit run app.py
```

Opens at **http://localhost:8501**.

---

### Step 4 — Connect your email (IMAP)

If you want to pull directly from Gmail, Outlook, or any IMAP server:

**Before you open the app — get an app password:**

- **Gmail:** Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords). You need 2-Step Verification enabled. Generate a password for "Mail". Copy it — you only see it once.
- **Outlook / Microsoft 365:** Go to [account.microsoft.com/security](https://account.microsoft.com/security) → Advanced security → App passwords.
- **iCloud:** Go to [appleid.apple.com](https://appleid.apple.com) → Sign-In and Security → App-Specific Passwords.
- **Other providers:** Use your normal password if your server doesn't require app passwords.

**In the app — Settings tab → IMAP connection:**

| Field | What to enter |
|---|---|
| IMAP host | `imap.gmail.com` / `imap-mail.outlook.com` / `imap.mail.me.com` |
| Email address | Your full email address |
| Password / app password | The app password you just generated |
| Port | `993` (leave as-is for all major providers) |
| Use SSL | Checked (leave as-is) |
| Mailbox | `INBOX` to start — you can change this later |

1. Click **Save IMAP settings**
2. Click **Test connection** — you should see "Connected successfully"
3. Scroll down to **Fetch emails via IMAP**
4. Set **Max emails** to something larger than your inbox (e.g. `20000`)
5. Click **Fetch & index emails**

A live status line will update as it works through your archive. For 15k emails expect this to take **30–90 minutes** — the bottleneck is the NLP pipeline (NER + embedding), not the network. You can leave it running and use the Search tab in parallel once a few thousand emails are indexed.

> Already-indexed emails are always skipped, so re-running is safe and resumes where it left off if interrupted.

---

## IMAP Reference

### Supported servers

| Provider | Host | Port |
|---|---|---|
| Gmail | `imap.gmail.com` | 993 |
| Outlook / Microsoft 365 | `imap-mail.outlook.com` | 993 |
| iCloud | `imap.mail.me.com` | 993 |
| Yahoo | `imap.mail.yahoo.com` | 993 |
| Self-hosted (SSL) | your server | 993 |
| Self-hosted (no SSL) | your server | 143 |

### Fetching multiple mailboxes

After the initial INBOX fetch, go back to Settings → Fetch emails via IMAP, change the **Mailbox** field, and run again. Common values:

- Gmail: `[Gmail]/Sent Mail`, `[Gmail]/All Mail`
- Outlook: `Sent Items`
- Generic: `Sent`, `Archive`

Each run only fetches UIDs not already in the database, so there's no duplication.

### Credential storage

IMAP credentials are saved to `data/settings.json` on your local machine. This file is listed in `.gitignore` and never leaves your machine. The password field in the UI can be left blank when saving other settings — it will keep the previously saved password.

### Advanced: scripted / programmatic use

```python
from modules.imap_connector import IMAPConnector

conn = IMAPConnector("imap.gmail.com", "you@gmail.com", "app-password")

# One-shot — handles 15k+ emails
result = conn.fetch_and_index(mailbox="INBOX")
# {"indexed": 14823, "skipped": 0, "errors": 0, "total_in_db": 14823}

# Multiple mailboxes
for mailbox in ["INBOX", "[Gmail]/Sent Mail"]:
    conn.fetch_and_index(mailbox=mailbox)

# Background polling (keeps index live)
conn.start(mailbox="INBOX", interval=300)
```

---

## Python 3.14 / Restricted Environments

`sentence-transformers` (and `torch`) do not yet have pre-built wheels for Python 3.14. The app handles this gracefully:

- **FTS search** works normally without any ML packages.
- **Semantic / hybrid search** modes are hidden from the UI if `sentence-transformers` is unavailable.
- **NLP auto-classification** falls back to the TF-IDF method (no ML dependencies).
- **Keyword extraction and NER** are silently skipped if the spaCy model is missing.
- The **sidebar** shows a warning for each unavailable component.

### SSL / cert-restricted networks

If `pip install` or model downloads fail due to SSL errors, pre-packaged model bundles are on [GitHub Releases (models-v1)](https://github.com/shawntanzk/eml-search/releases/tag/models-v1). `setup_models.py` will fall back to these automatically.

Manual install:

```bash
# spaCy model (12 MB) — into your venv:
tar -xzf en_core_web_sm-3.8.0.tar.gz -C .venv/lib/python3.*/site-packages/

# Sentence-transformer model (80 MB) — into HuggingFace cache:
tar -xzf all-MiniLM-L6-v2.tar.gz -C ~/.cache/huggingface/hub/
```

---

## Model Locations

| Model | Location |
|---|---|
| spaCy `en_core_web_sm` | `<venv>/lib/python3.*/site-packages/en_core_web_sm/` |
| sentence-transformers `all-MiniLM-L6-v2` | `~/.cache/huggingface/hub/models--sentence-transformers--all-MiniLM-L6-v2/` |

---

## Project Structure

```
eml_search_app/
├── app.py                   # Streamlit UI — 4 tabs: Search, Tags, Knowledge Graph, Settings
├── config.py                # Paths, model names, thresholds
├── setup_models.py          # One-time setup: models → database → index
├── requirements.txt
│
├── modules/
│   ├── eml_parser.py        # .eml file → structured dict (stdlib only)
│   ├── indexer.py           # SQLite/FTS5 CRUD
│   ├── watcher.py           # Background folder watcher + indexing pipeline
│   ├── imap_connector.py    # IMAP ingestion — fetch from live server, same pipeline as watcher
│   ├── nlp_engine.py        # spaCy NER + keyword extraction (optional)
│   ├── semantic_search.py   # Sentence-transformer embed + cosine search (optional)
│   ├── tfidf_classifier.py  # Pure Python/numpy TF-IDF classifier (always available)
│   ├── tagger.py            # Tag library, manual assignment, NLP auto-classification
│   ├── graph_builder.py     # RDF/OWL ABox builder + SPARQL + pyvis
│   └── search_engine.py     # FTS / semantic / hybrid search orchestration
│
├── ontology/
│   └── email_ontology.ttl   # TBox — editable OWL DL ontology schema
│
├── data/                    # Generated at runtime (gitignored)
│   ├── index.db             # SQLite index (WAL mode)
│   ├── email_data.ttl       # ABox — generated RDF instance data
│   ├── graph_preview.html   # Temp file for pyvis graph render
│   └── settings.json        # Persisted settings incl. IMAP credentials
│
└── test_emails/             # 100 synthetic .eml files for development/testing
```

---

## How It Works

### Indexing Pipeline

Every new email — from a local `.eml` file or fetched via IMAP — goes through the same stages:

```
Source A: .eml file          Source B: IMAP server
eml_parser.parse_eml()       imap_connector._parse_message()
         │                              │
         └──────────────┬───────────────┘
                        │
                  structured dict
                  {id, file_path, subject, sender, body_text, …}
                        │
                        ▼
1. indexer.insert_email()
   Writes to `emails` table. FTS5 trigger indexes tokens automatically.
   IMAP emails get a virtual file_path: imap://host/mailbox/uid
    │
    ▼
2. nlp_engine.extract_entities()   [skipped if spaCy model missing]
   en_core_web_sm NER → PERSON, ORG, GPE, LOC stored in `email_entities`
    │
    ▼
3. semantic_search.embed_batch()   [skipped if sentence-transformers missing]
   all-MiniLM-L6-v2 — batched for efficiency (64 emails per model pass)
   384-dim float32 vectors stored as BLOB in `embeddings`
```

**Local folder** — `EmailWatcher` daemon thread, polls every 10 seconds.

**IMAP** — `IMAPConnector` fetches 100 UIDs per round trip, batches embeddings 64 at a time. After the first full run, subsequent polls only ask the server for UIDs newer than the last seen (cached max UID), so polling a large mailbox is cheap.

---

### Search

#### Full-Text Search
SQLite FTS5 + Porter stemmer, BM25 ranking. Always available.

#### Semantic Search
Query embedded by the same model used at index time. Dot product against the embedding matrix, O(n) top-k via `numpy.argpartition`. Requires `sentence-transformers`.

#### Hybrid Search (default)
Reciprocal Rank Fusion merges FTS and semantic rankings:
```
score(d) = Σ  1 / (k + rank_i(d))
```
Falls back to FTS if `sentence-transformers` is unavailable.

---

### Tags

**Manual assignment** — add or remove per email in the Search tab. Manually removed tags are permanently blocked from NLP re-assignment for that email.

**NLP auto-classification** — runs from the Tags tab:

| Method | Dependencies | Default threshold |
|---|---|---|
| **Semantic** | sentence-transformers | 0.25 |
| **TF-IDF** | numpy only (always available) | 0.15 |

Both only add, never remove. Per-tag method and threshold are configurable.

---

### Knowledge Graph (RDF/OWL)

| File | Contents | Editable? |
|---|---|---|
| `ontology/email_ontology.ttl` | TBox — classes, properties, axioms | Yes |
| `data/email_data.ttl` | ABox — individuals from indexed emails | No — rebuilt on demand |

Click **Build / Rebuild Graph** to regenerate the ABox.

**Ontology classes:**
```
owl:Thing
 ├── ont:Email           — one individual per indexed email
 ├── ont:Person          — senders, recipients, NER-extracted people
 ├── ont:Organization    — NER-extracted ORG entities
 ├── ont:Location        — NER-extracted GPE/LOC entities
 ├── ont:Tag             — one individual per defined tag
 └── ont:Thread          — one individual per unique thread root
```

**SPARQL example:**
```sparql
PREFIX ont: <http://emailsearch.local/ontology#>

SELECT ?subject ?sender
WHERE {
  ?email a ont:Email ;
         ont:hasSubject ?subject ;
         ont:hasSender ?person .
  ?person ont:emailAddress ?sender .
  FILTER(CONTAINS(?sender, "example.com"))
}
LIMIT 20
```

---

### Database Schema

| Table | Purpose |
|---|---|
| `emails` | Core email metadata and body text |
| `emails_fts` | FTS5 virtual table (Porter stemmer, auto-synced via triggers) |
| `email_entities` | NER results (email_id, entity_text, entity_label) |
| `embeddings` | float32 sentence-transformer vectors as BLOB |
| `tags` | Tag library (id, name, nlp_method, nlp_threshold) |
| `email_tags` | Tag assignments (email_id, tag_id, source: 'manual'\|'nlp') |
| `email_tag_blocks` | Blocks NLP from re-adding manually removed tags |
| `meta` | Key-value store — incl. cached IMAP max UIDs |

SQLite runs in **WAL mode** so the background indexer can write while Streamlit reads.

---

## Configuration

All tuneable values in `config.py`:

| Variable | Default | Description |
|---|---|---|
| `SPACY_MODEL` | `en_core_web_sm` | spaCy model name |
| `SENTENCE_TRANSFORMER_MODEL` | `all-MiniLM-L6-v2` | Embedding model |
| `SEMANTIC_TOP_K` | `100` | Candidates retrieved by semantic search |
| `MAX_SEARCH_RESULTS` | `200` | Hard cap on returned results |
| `WATCH_POLL_INTERVAL` | `10` | Seconds between folder polls |

---

## Extending the Ontology

Edit `ontology/email_ontology.ttl`, then click **Build / Rebuild Graph**. Example:

```turtle
ont:Project
    a owl:Class ;
    rdfs:label "Project" .

ont:relatedToProject
    a owl:ObjectProperty ;
    rdfs:domain ont:Email ;
    rdfs:range  ont:Project .
```

To populate the new property automatically, extend `build_abox()` in `modules/graph_builder.py`.

---

## Test Data

100 synthetic `.eml` files in `test_emails/` — useful for verifying setup without your real archive:

```bash
python setup_models.py --folder test_emails
```

---

## Requirements

Core (always required):
```
streamlit>=1.32.0
rdflib>=7.0.0
pandas>=2.0.0
pyvis>=0.3.2
numpy>=1.24.0
```

Optional (app degrades gracefully without these):
```
spacy>=3.7.0                 # NER and keyword extraction
sentence-transformers>=2.7.0 # Semantic/hybrid search and semantic tag classification
```
