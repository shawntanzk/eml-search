# EML Search

A fully offline Streamlit app for indexing, searching, and analysing a local folder of `.eml` email backup files. No cloud services, no LLMs — everything runs on your machine.

---

## Features

| Capability | Detail |
|---|---|
| **Full-text search** | SQLite FTS5 with Porter stemming — fast queries across thousands of emails |
| **Semantic search** | Sentence-transformer embeddings (all-MiniLM-L6-v2) with cosine similarity — requires `sentence-transformers` |
| **Hybrid search** | Reciprocal Rank Fusion (RRF) merges FTS and semantic rankings |
| **Filters** | Sender, date range, has-attachments, tag |
| **Auto-indexing** | Background thread watches the folder and indexes new `.eml` files as they arrive |
| **Keyword extraction** | Per-email keyword panel using spaCy noun chunks — requires spaCy model |
| **Named entity recognition** | Extracts people, organisations, locations via spaCy — requires spaCy model |
| **Tag management** | Human-defined tag library with manual assignment and NLP auto-classification |
| **NLP auto-classification** | Assign tags automatically via **Semantic** (sentence-transformers) or **TF-IDF** (no ML dependencies) |
| **Knowledge graph** | RDF/OWL graph with SPARQL console and interactive pyvis visualisation |
| **Graceful degradation** | App runs fully with FTS-only search even if spaCy or sentence-transformers are unavailable |

---

## Quick Start

### Prerequisites

- Python 3.10+ (including 3.14+ — see [Python 3.14 / restricted environments](#python-314--restricted-environments))
- `pip`

### 1. Install dependencies

```bash
cd eml_search_app
pip install -r requirements.txt
```

### 2. Run one-time setup

Downloads NLP models, initialises the database, and indexes your emails.

```bash
python setup_models.py --folder /path/to/your/eml/folder
```

The folder path is saved to `data/settings.json` after the first run.

### 3. Start the app

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`.

---

## Python 3.14 / Restricted Environments

`sentence-transformers` (and its dependency `torch`) do not yet have pre-built wheels for Python 3.14. The app handles this gracefully:

- **FTS search** works normally without any ML packages.
- **Semantic / hybrid search** modes are hidden from the UI if `sentence-transformers` is unavailable.
- **NLP auto-classification** falls back to the TF-IDF method (no ML dependencies).
- **Keyword extraction and NER** are silently skipped if the spaCy model is missing.
- The **sidebar** displays a warning for each unavailable NLP component.

### SSL / cert-restricted networks

If `pip install` or model downloads fail due to SSL certificate errors, pre-packaged model bundles are available on [GitHub Releases (models-v1)](https://github.com/shawntanzk/eml-search/releases/tag/models-v1). `setup_models.py` will automatically fall back to these if the standard download fails.

You can also install them manually:

**spaCy model** (`en_core_web_sm`, 12 MB) — extract into your environment's `site-packages`:
```bash
# With .venv:
tar -xzf en_core_web_sm-3.8.0.tar.gz -C .venv/lib/python3.14/site-packages/

# System Python:
tar -xzf en_core_web_sm-3.8.0.tar.gz -C $(python3 -c "import site; print(site.getsitepackages()[0])")
```

**Sentence-transformer model** (`all-MiniLM-L6-v2`, 80 MB) — extract into the HuggingFace cache (not the venv):
```bash
tar -xzf all-MiniLM-L6-v2.tar.gz -C ~/.cache/huggingface/hub/
```

---

## Model Locations

| Model | Location |
|---|---|
| spaCy `en_core_web_sm` | `<site-packages>/en_core_web_sm/` (inside your venv or system Python) |
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
│   ├── eml_parser.py        # .eml → structured dict (stdlib only)
│   ├── indexer.py           # SQLite/FTS5 CRUD
│   ├── watcher.py           # Background folder watcher + indexing pipeline
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
├── data/                    # Generated at runtime (not in repo)
│   ├── index.db             # SQLite index (WAL mode)
│   ├── email_data.ttl       # ABox — generated RDF instance data
│   ├── graph_preview.html   # Temp file for pyvis graph render
│   └── settings.json        # Persisted settings (email folder path)
│
└── test_emails/             # 100 synthetic .eml files for development/testing
```

---

## How It Works

### Indexing Pipeline

Every new `.eml` file goes through four stages:

```
.eml file
    │
    ▼
1. eml_parser.parse_eml()
   Extracts: subject, sender, recipients, CC, date, body text,
   attachment names, thread ID (via In-Reply-To / References)
    │
    ▼
2. indexer.insert_email()
   Writes to `emails` table. A trigger populates the FTS5 virtual
   table (emails_fts) with Porter-stemmed tokens automatically.
    │
    ▼
3. nlp_engine.extract_entities()   [skipped if spaCy model missing]
   Runs en_core_web_sm NER on subject + body.
   PERSON, ORG, GPE, LOC entities stored in `email_entities`.
    │
    ▼
4. semantic_search.embed_text()    [skipped if sentence-transformers missing]
   Encodes subject + body via all-MiniLM-L6-v2.
   384-dim l2-normalised float32 vector stored as BLOB in `embeddings`.
```

The **background watcher** (`EmailWatcher`) runs this pipeline in a daemon thread, polling the folder every 10 seconds.

---

### Search

#### Full-Text Search (FTS)
SQLite FTS5 with the Porter stemmer. Results ranked by BM25. Always available regardless of NLP dependencies.

#### Semantic Search
The query is embedded by the same sentence-transformer model used at index time. Dot product against stored embeddings gives cosine similarity (`numpy.argpartition` for O(n) top-k). Only available when `sentence-transformers` is installed.

#### Hybrid Search (default)
Combines FTS and semantic rankings with **Reciprocal Rank Fusion**:

```
score(d) = Σ  1 / (k + rank_i(d))
```

Falls back to FTS automatically if `sentence-transformers` is unavailable.

---

### Tags

Tags are human-defined categories. The tag system has two layers:

**Manual assignment** — add or remove tags per email from the Search tab. Manually removed tags are permanently blocked from being re-added by NLP auto-classification for that email.

**NLP auto-classification** — runs from the Tags tab, two methods to choose from:

| Method | Dependencies | Notes |
|---|---|---|
| **Semantic** | sentence-transformers | Embeds tag name and compares to stored email embeddings. More accurate. Default threshold 0.25. |
| **TF-IDF** | numpy only (always available) | Builds a TF-IDF matrix from email bodies, scores each tag name as a keyword query. Default threshold 0.15. |

Both methods only add tags, never remove. Manual blocks are always respected.

---

### Knowledge Graph (RDF/OWL)

The graph is divided into two strictly separate files:

| File | Name | Contents | Editable? |
|---|---|---|---|
| `ontology/email_ontology.ttl` | **TBox** | Classes, properties, domain/range axioms | Yes |
| `data/email_data.ttl` | **ABox** | Named individuals generated from emails | No — rebuilt on demand |

Click **Build / Rebuild Graph** to regenerate the ABox from the current database.

#### Ontology Classes

```
owl:Thing
 ├── ont:Email           — one individual per indexed email
 ├── ont:Person          — senders, recipients, NER-extracted people
 ├── ont:Organization    — NER-extracted ORG entities
 ├── ont:Location        — NER-extracted GPE/LOC entities
 ├── ont:Tag             — one individual per defined tag
 ├── ont:Thread          — one individual per unique thread root
 └── ont:Attachment      — declared in TBox; populated when metadata present
```

#### SPARQL

The Knowledge Graph tab includes a SPARQL console. Queries run against the merged TBox + ABox. Result rows show **Open email** / **Search sender** / **Search** buttons — "Open email" navigates directly to that specific email in the Search tab.

Example:
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
| `embeddings` | float32 sentence-transformer vectors stored as BLOB |
| `tags` | Tag library (id, name) |
| `email_tags` | Tag assignments (email_id, tag_id, source: 'manual'\|'nlp') |
| `email_tag_blocks` | Blocks NLP from re-adding manually removed tags |
| `meta` | Key-value store for app state |

SQLite runs in **WAL mode** so the background indexer can write while Streamlit reads.

---

## Configuration

All tuneable values live in `config.py`:

| Variable | Default | Description |
|---|---|---|
| `SPACY_MODEL` | `en_core_web_sm` | spaCy model name |
| `SENTENCE_TRANSFORMER_MODEL` | `all-MiniLM-L6-v2` | Embedding model |
| `SEMANTIC_TOP_K` | `100` | Candidates retrieved by semantic search |
| `MAX_SEARCH_RESULTS` | `200` | Hard cap on returned results |
| `WATCH_POLL_INTERVAL` | `10` | Seconds between folder polls |

---

## Extending the Ontology

Edit `ontology/email_ontology.ttl` in any text editor, then click **Build / Rebuild Graph** in the app. Example — adding a `Project` class:

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

100 synthetic `.eml` files are in `test_emails/`. Use them to get started without your real archive:

```bash
python setup_models.py --folder test_emails
```

---

## Requirements

Core (always required):
```
streamlit>=1.32.0
watchdog>=4.0.0
rdflib>=7.0.0
pandas>=2.0.0
pyvis>=0.3.2
beautifulsoup4>=4.12.0
numpy>=1.24.0
```

Optional (app degrades gracefully without these):
```
spacy>=3.7.0                 # NER and keyword extraction
sentence-transformers>=2.7.0 # Semantic/hybrid search and semantic tag classification
```
