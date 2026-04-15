# EML Search

A fully offline Streamlit app for indexing, searching, and analysing a local folder of `.eml` email backup files. No cloud services, no LLMs — everything runs on your machine after a one-time model download.

---

## Features

| Capability | Detail |
|---|---|
| **Full-text search** | SQLite FTS5 with Porter stemming — sub-millisecond queries across thousands of emails |
| **Semantic search** | Sentence-transformer embeddings (all-MiniLM-L6-v2) with cosine similarity |
| **Hybrid search** | Reciprocal Rank Fusion (RRF) merges FTS and semantic rankings |
| **Filters** | Sender, date range, has-attachments, topic |
| **Auto-indexing** | Background thread watches the folder and indexes new `.eml` files as they arrive |
| **NLP insights** | Named entity recognition (people, organisations, locations), topic modelling, keyword extraction |
| **Knowledge graph** | RDF/OWL DL-compliant graph with strict TBox/ABox separation; SPARQL query console; interactive visualisation |
| **Analytics dashboard** | Volume over time, top senders, domain breakdown, topic distribution, hourly heatmap |

---

## Quick Start

### Prerequisites

- Python 3.10+
- `pip`

### 1. Install dependencies

```bash
cd eml_search_app
pip install -r requirements.txt
```

### 2. Run one-time setup

This downloads the NLP models (~100 MB total), initialises the database, and indexes your emails.

```bash
python setup_models.py --folder /path/to/your/eml/folder
```

After the initial run the folder path is saved to `data/settings.json` and does not need to be passed again.

> **No internet after setup.** Both models are cached locally:
> - spaCy `en_core_web_sm` → your Python environment's `site-packages`
> - sentence-transformers `all-MiniLM-L6-v2` → `~/.cache/torch/sentence_transformers/`

### 3. Start the app

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Setup Script Options

```
python setup_models.py [--folder PATH]
```

| Argument | Default | Description |
|---|---|---|
| `--folder` | Value saved in `data/settings.json`, else `test_emails/` | Absolute path to the folder containing `.eml` files |

The script runs four steps in order and reports progress:

1. Downloads / verifies the spaCy model
2. Downloads / verifies the sentence-transformer model
3. Initialises the SQLite database and FTS5 index
4. Indexes all `.eml` files found recursively in the folder, then trains the LDA topic model

Re-running the script is safe — already-indexed files are skipped and only new ones are processed.

---

## Project Structure

```
eml_search_app/
├── app.py                  # Streamlit UI (4 tabs)
├── config.py               # Paths, model names, thresholds
├── setup_models.py         # One-time setup script
├── requirements.txt
│
├── modules/
│   ├── eml_parser.py       # .eml → structured dict
│   ├── indexer.py          # SQLite/FTS5 read/write
│   ├── watcher.py          # Background folder watcher + indexing pipeline
│   ├── nlp_engine.py       # spaCy NER + sklearn LDA topics + TF-IDF keywords
│   ├── semantic_search.py  # Sentence-transformer embed + cosine search
│   ├── graph_builder.py    # RDF/OWL ABox builder + SPARQL + pyvis edges
│   ├── search_engine.py    # FTS / semantic / hybrid search orchestration
│   └── insights.py         # Analytics queries → Pandas DataFrames
│
├── ontology/
│   └── email_ontology.ttl  # TBox — editable OWL DL ontology schema
│
├── data/                   # Generated at runtime (not in repo)
│   ├── index.db            # SQLite index
│   ├── email_data.ttl      # ABox — generated RDF instance data
│   └── settings.json       # Persisted settings
│
├── models/                 # Persisted ML models (generated at runtime)
│   ├── topic_model.pkl     # Trained LDA model
│   └── tfidf_vectorizer.pkl
│
└── test_emails/            # 100 synthetic .eml files for testing
```

---

## How It Works

### Indexing Pipeline

Every new `.eml` file goes through five stages, each result stored in SQLite:

```
.eml file
    │
    ▼
1. eml_parser.parse_eml()
   Extracts: subject, sender, recipients, CC, date, body text,
   attachment names, thread ID (via In-Reply-To / References headers)
    │
    ▼
2. indexer.insert_email()
   Writes to the `emails` table. A trigger simultaneously populates
   the `emails_fts` FTS5 virtual table with Porter-stemmed tokens.
    │
    ▼
3. nlp_engine.extract_entities()
   Runs spaCy en_core_web_sm NER on subject + body.
   Extracts PERSON, ORG, GPE (country/city), LOC entities.
   Stored in `email_entities` table.
    │
    ▼
4. semantic_search.embed_text()
   Encodes subject + first 400 chars of body via
   sentence-transformers all-MiniLM-L6-v2.
   384-dimensional l2-normalised float32 vector stored as BLOB
   in the `embeddings` table.
    │
    ▼
5. nlp_engine.assign_topics()
   If an LDA model is trained, transforms the email body through
   the TF-IDF vectorizer and LDA to get a topic probability
   distribution. Top topics stored in `email_topics`.
```

The **background watcher** (`EmailWatcher`) runs this pipeline in a daemon thread, polling the folder every 10 seconds. New files are discovered by comparing resolved file paths against the `emails.file_path` column.

After every 50 new emails the LDA model is **retrained on the full corpus** and all topic assignments are updated in bulk.

---

### Search

#### Full-Text Search (FTS)

Uses SQLite's built-in FTS5 extension with the Porter stemming tokeniser. A query like `"invoice payment"` is tokenised and stemmed, then matched against the `subject`, `sender_name`, `sender_email`, and `body_text` columns. Results are ranked by BM25 (SQLite's built-in `rank`). Filters are applied as `WHERE` clauses joined against the base `emails` table.

#### Semantic Search

The query string is embedded by the same sentence-transformer model used at index time. The resulting 384-dim vector is compared against every stored embedding using a dot product (equivalent to cosine similarity since all vectors are l2-normalised). `numpy.argpartition` is used for an O(n) top-k selection, making this fast even for large corpora without requiring a dedicated vector database.

#### Hybrid Search (default)

Combines both ranking lists using **Reciprocal Rank Fusion**:

```
score(d) = Σ  1 / (k + rank_i(d))
```

where `k = 60` and `rank_i(d)` is the document's position in each ranking list. Documents appearing in both lists are boosted; documents missing from one list still contribute via the other. The merged list is re-sorted by RRF score before returning results.

---

### NLP

#### Named Entity Recognition

spaCy's `en_core_web_sm` model identifies entities in each email. The following entity types are extracted and stored:

| spaCy label | Stored as | Used in graph |
|---|---|---|
| `PERSON` | `email_entities` | `ont:mentionsPerson` |
| `ORG` | `email_entities` | `ont:mentionsOrganization` |
| `GPE` | `email_entities` | `ont:mentionsLocation` |
| `LOC` | `email_entities` | `ont:mentionsLocation` |

#### Topic Modelling (LDA)

Latent Dirichlet Allocation is run via `sklearn.decomposition.LatentDirichletAllocation`:

1. Email bodies are cleaned (URLs, email addresses, and punctuation removed; lowercased)
2. A `TfidfVectorizer` builds a 5,000-feature document-term matrix (`min_df=2`, `max_df=0.85`, English stop words removed)
3. LDA decomposes this into N topics (default 10, capped at `n_emails // 5`)
4. Each topic is labelled with its top 3 keywords (e.g. `"invoice / payment / amount"`)
5. Each email receives a full probability distribution over topics; assignments with score ≥ 0.10 are stored

The model is retrained automatically after every 50 new emails, updating all assignments.

#### Keyword Extraction

Per-email keywords are extracted on demand (in the search results view) by running the stored TF-IDF vectorizer on the single document and returning the top-scoring terms.

---

### Knowledge Graph (RDF/OWL)

The graph layer is divided into two strictly separate files:

| File | Name | Contents | Editable? |
|---|---|---|---|
| `ontology/email_ontology.ttl` | **TBox** | Classes, object properties, data properties, domain/range axioms | Yes — edit freely |
| `data/email_data.ttl` | **ABox** | Named individuals generated from indexed emails | No — regenerated automatically |

This separation means you can refine or extend the ontology schema at any time without touching the data graph. Rebuilding the ABox (via the **Build / Rebuild Graph** button) re-reads the current TBox and regenerates `email_data.ttl` from scratch.

#### Ontology Classes

```
owl:Thing
 ├── ont:Email           — one individual per indexed email
 ├── ont:Person          — sender/recipient (by address) + NER-extracted people (by name)
 ├── ont:Organization    — NER-extracted ORG entities
 ├── ont:Location        — NER-extracted GPE/LOC entities
 ├── ont:Topic           — one individual per LDA topic
 ├── ont:Thread          — one individual per unique thread root
 └── ont:Attachment      — declared in TBox, populated when attachment metadata is present
```

#### Key Object Properties

| Property | Domain | Range | Note |
|---|---|---|---|
| `ont:hasSender` | Email | Person | `owl:FunctionalProperty` |
| `ont:hasRecipient` | Email | Person | |
| `ont:belongsToThread` | Email | Thread | |
| `ont:hasTopic` | Email | Topic | only topics with score ≥ 0.10 |
| `ont:mentionsPerson` | Email | Person | from NER |
| `ont:mentionsOrganization` | Email | Organization | from NER |
| `ont:mentionsLocation` | Email | Location | from NER |
| `ont:worksFor` | Person | Organization | declared in TBox, ready for inference |

#### SPARQL

The **Knowledge Graph** tab includes a SPARQL query console. Queries run against the merged TBox + ABox graph loaded in memory. Example — find all emails from a specific domain that are assigned to a topic:

```sparql
PREFIX ont: <http://emailsearch.local/ontology#>
PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>

SELECT ?subject ?sender ?topicLabel
WHERE {
  ?email rdf:type ont:Email ;
         ont:hasSubject ?subject ;
         ont:hasSender ?person ;
         ont:hasTopic ?topic .
  ?person ont:emailAddress ?sender .
  ?topic rdfs:label ?topicLabel .
  FILTER(CONTAINS(?sender, "example.com"))
}
LIMIT 50
```

---

### Database Schema

```
emails            core email metadata + body text
emails_fts        FTS5 virtual table (Porter stemmer, auto-synced via INSERT/DELETE triggers)
email_entities    (email_id, entity_text, entity_label)
embeddings        (email_id, vector BLOB)  — float32 numpy array, 384 dims
topics            (id, words JSON, label)
email_topics      (email_id, topic_id, score)
meta              key/value store for internal state
```

SQLite runs in **WAL mode** (`PRAGMA journal_mode=WAL`) so the background indexer thread can write while Streamlit reads without locking.

---

## Configuration

All tuneable values live in `config.py`:

| Variable | Default | Description |
|---|---|---|
| `SPACY_MODEL` | `en_core_web_sm` | spaCy model name |
| `SENTENCE_TRANSFORMER_MODEL` | `all-MiniLM-L6-v2` | Embedding model |
| `NUM_TOPICS` | `10` | Max LDA topics |
| `NUM_TOPIC_WORDS` | `8` | Keywords per topic |
| `MIN_EMAILS_FOR_TOPICS` | `10` | Minimum corpus size before LDA trains |
| `TOPIC_RETRAIN_THRESHOLD` | `50` | New emails before LDA auto-retrains |
| `SEMANTIC_TOP_K` | `100` | Candidates retrieved by semantic search |
| `MAX_SEARCH_RESULTS` | `200` | Hard cap on returned results |
| `WATCH_POLL_INTERVAL` | `10` | Seconds between folder polls |

---

## Extending the Ontology

Open `ontology/email_ontology.ttl` in any text editor and add classes or properties using standard Turtle syntax. For example, to add a `Project` class and link emails to it:

```turtle
ont:Project
    a owl:Class ;
    rdfs:label "Project" ;
    rdfs:comment "A work project referenced in emails." .

ont:relatedToProject
    a owl:ObjectProperty ;
    rdfs:label "related to project" ;
    rdfs:domain ont:Email ;
    rdfs:range  ont:Project .
```

After saving the file, click **Build / Rebuild Graph** in the app. The ABox generator in `modules/graph_builder.py` picks up the new schema on the next build. To populate the new property automatically, extend `build_abox()` in `graph_builder.py` with the corresponding logic.

---

## Offline Operation

After running `setup_models.py` once the app requires no network access:

| Component | Cached at |
|---|---|
| spaCy `en_core_web_sm` | Python env `site-packages/en_core_web_sm` |
| sentence-transformers model | `~/.cache/torch/sentence_transformers/` |
| LDA + TF-IDF models | `models/topic_model.pkl`, `models/tfidf_vectorizer.pkl` |
| All email data | `data/index.db` |
| RDF graph | `data/email_data.ttl` |

---

## Test Data

100 synthetic `.eml` files are included in `test_emails/` for development and testing. Run the setup script pointing at this folder to get started immediately without your real email archive:

```bash
python setup_models.py --folder test_emails
```

---

## Requirements

```
streamlit>=1.32.0
watchdog>=4.0.0
rdflib>=7.0.0
spacy>=3.7.0
sentence-transformers>=2.7.0
scikit-learn>=1.4.0
pandas>=2.0.0
plotly>=5.20.0
pyvis>=0.3.2
beautifulsoup4>=4.12.0
numpy>=1.24.0
```
