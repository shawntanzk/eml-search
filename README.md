# EML Search

A fully local Streamlit app for indexing, searching, and analysing emails — with a live calendar view that links events to related emails. Pull from a local folder of `.eml` files **or connect directly to an IMAP server** — no cloud services, no LLMs, everything runs on your machine.

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
| **Calendar tab** | Month + week + list views linked to your email archive — see related emails for any event |
| **Live calendar sync** | Pull events from iCal URL, Apple iCloud, Google Calendar, Outlook.com, or Microsoft 365 Graph API |
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

One-time download (~100 MB). Also initialises the database.

```bash
python setup_models.py
```

Expected output:

```
Downloading spaCy model...       ✓
Downloading sentence-transformer... ✓
Database initialised.
```

To index a local `.eml` folder right away:

```bash
python setup_models.py --folder /path/to/your/eml/folder
```

### Step 3 — Start the app

```bash
streamlit run app.py
```

Opens at **http://localhost:8501**.

---

## Connect Your Email (IMAP)

### Get an app password first

| Provider | Where to get it |
|---|---|
| Gmail | [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) — needs 2-Step Verification |
| Outlook / Microsoft 365 | [account.microsoft.com/security](https://account.microsoft.com/security) → Advanced security → App passwords |
| iCloud | [appleid.apple.com](https://appleid.apple.com) → Sign-In and Security → App-Specific Passwords |
| Other | Use your normal password if no app password required |

### In the app — Settings → IMAP connection

| Field | Value |
|---|---|
| IMAP host | `imap.gmail.com` / `imap-mail.outlook.com` / `imap.mail.me.com` |
| Email address | Your full email |
| Password | The app password |
| Port | `993` |
| Mailbox | `INBOX` |

1. Click **Save IMAP settings**
2. Click **Test connection** — should show "Connected successfully"
3. Scroll to **Fetch emails via IMAP**, set **Max emails**, click **Fetch & index emails**

For 15k emails expect 30–90 minutes — the bottleneck is the NLP pipeline, not the network. Already-indexed emails are always skipped.

### IMAP host reference

| Provider | Host | Port |
|---|---|---|
| Gmail | `imap.gmail.com` | 993 |
| Outlook / Microsoft 365 | `imap-mail.outlook.com` | 993 |
| iCloud | `imap.mail.me.com` | 993 |
| Yahoo | `imap.mail.yahoo.com` | 993 |
| Self-hosted (SSL) | your server | 993 |
| Self-hosted (no SSL) | your server | 143 |

---

## Calendar

The Calendar tab shows a month view, week view, and list view of your events, and automatically finds emails related to any selected event — ranked by attendee overlap, subject match, semantic similarity, and tags.

You can connect as many calendars as you like. Events from all sources are merged into one view and colour-coded by calendar — the same way Apple Calendar works.

> **Offline mode** — only JSON file calendars are active. iCal and Microsoft 365 accounts are silently skipped. Switch to Online mode in Settings to enable them.

---

### How to add a calendar

All calendar management is in **Settings → Calendar accounts**:

1. Click **＋ Add calendar account**
2. Enter a **Name** (e.g. "Google Work"), choose a **Type**, pick a **colour**
3. Fill in the type-specific fields (see guides below)
4. Click **Save**

Repeat for as many calendars as you want. Each one gets its own colour dot in the month, week, and list views. You can enable/disable any account without deleting it.

---

### Connect Google Calendar

**Get the secret iCal link:**

1. Go to [calendar.google.com](https://calendar.google.com) → click the gear icon → **Settings**
2. In the left sidebar, click the calendar you want under **Settings for my calendars**
3. Scroll down to **Integrate calendar**
4. Copy the **Secret address in iCal format** — it looks like:
   ```
   https://calendar.google.com/calendar/ical/youraddress%40gmail.com/private-xxxxxxxx/basic.ics
   ```

> This URL is private and acts as your password — don't share it. No username or password needed.

**Add to the app:**

1. Settings → Calendar accounts → **＋ Add calendar account**
2. Name: `Google Calendar` (or whatever you like)
3. Type: **iCal URL**
4. iCal URL: paste the link from step 4 above
5. Leave username and password blank
6. Click **Test fetch** to verify, then **Save**

To add multiple Google calendars (personal, work, shared), repeat this for each one — each has its own secret iCal link in Google Calendar Settings.

---

### Connect Apple Calendar (iCloud)

**Get the iCloud feed URL:**

1. On your Mac: open **Calendar** → right-click the calendar in the sidebar → **Share Calendar…**
   - Or on [icloud.com](https://www.icloud.com/calendar): click the broadcast icon next to a calendar
2. Tick **Public Calendar** — a `webcal://` URL appears. Copy it.
   (The app automatically converts `webcal://` → `https://`, so you can paste it as-is.)

> **iCloud requires an app-specific password** — the regular iCloud password won't work for URL-based access.
>
> 1. Go to [appleid.apple.com](https://appleid.apple.com) → **Sign-In and Security** → **App-Specific Passwords**
> 2. Click **+** → name it "EML Search" → copy the generated password (shown only once)

**Add to the app:**

1. Settings → Calendar accounts → **＋ Add calendar account**
2. Name: `iCloud Calendar`
3. Type: **iCal URL**
4. iCal URL: paste the `webcal://` or `https://` link
5. Username: your iCloud email (e.g. `you@icloud.com`)
6. App-specific password: the password you generated above
7. Click **Test fetch** to verify, then **Save**

To add multiple iCloud calendars (each calendar has its own share URL), repeat for each one.

---

### Connect Outlook Calendar (Outlook.com / personal Microsoft account)

**Get the ICS link:**

1. Go to [outlook.live.com](https://outlook.live.com) → **Calendar**
2. Click the gear icon → **View all Outlook settings** → **Calendar** → **Shared calendars**
3. Under **Publish a calendar**, choose your calendar, set permission to **Can view all details**, click **Publish**
4. Copy the **ICS** link that appears

**Add to the app:**

1. Settings → Calendar accounts → **＋ Add calendar account**
2. Name: `Outlook Calendar`
3. Type: **iCal URL**
4. iCal URL: paste the ICS link
5. Leave username and password blank
6. Click **Test fetch** to verify, then **Save**

---

### Connect Microsoft 365 / Work Outlook (Graph API)

This uses OAuth2 device-code sign-in — no redirect URL needed. You can reuse the Azure App Registration already set up for IMAP email.

#### Step A — Azure App Registration (skip if already done for IMAP)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `EML Search` (or anything)
3. Supported account types: **Personal Microsoft accounts only** (or multi-tenant if needed)
4. No redirect URI needed — click **Register**
5. Copy the **Application (Client) ID**

#### Step B — Add the Calendar permission

1. In your app registration → **API permissions** → **Add a permission**
2. Choose **Microsoft Graph** → **Delegated permissions**
3. Search **Calendars.Read** → tick it → **Add permissions**
4. Click **Grant admin consent** if your org requires it

> Already set up IMAP? Your app registration already exists — just add `Calendars.Read` to it in step B.

#### Step C — Add to the app

1. Settings → Calendar accounts → **＋ Add calendar account**
2. Name: `Work Outlook`
3. Type: **Microsoft 365**
4. Azure Client ID: paste from Step A
5. Set **Days back** and **Days forward** (e.g. 30 / 90)
6. Click **Sign in with Microsoft** — you'll get a short code
7. Go to [microsoft.com/devicelogin](https://microsoft.com/devicelogin), enter the code, and approve
8. Back in the app, click **Complete sign-in**, then **Save**

Events are cached for the configured refresh interval (default 15 min). Use **🔄 Refresh** in the Calendar tab to force an immediate re-fetch.

---

### Other calendar apps

Any app that publishes an `.ics` URL can be added as an **iCal URL** account:

| App | Where to find the URL |
|---|---|
| Fantastical / BusyCal | Calendar sharing settings → "Subscribe" link |
| Exchange on-premise | OWA → Calendar → Share → Copy ICS link (or ask IT) |
| Nextcloud | Right-click calendar → Copy private link (change `webcal://` → `https://`) |
| Fastmail | Settings → Calendar → Manage → Share → ICS link |
| Any CalDAV server | Check your server's calendar export/subscribe options |

---

### Related email ranking

When you select an event, the app finds the most relevant emails in your archive using:

| Signal | Weight |
|---|---|
| 👥 Attendee/organiser email directly matches sender, recipient, or CC | Highest — 0.50 boost |
| 📝 Subject keyword match (FTS on event title) | High — 2× RRF weight |
| 🔍 Semantic match on invite text | Standard RRF weight |
| 🏷 Named entity overlap (people/orgs in invite vs emails) | Low |
| 🔖 Tag keyword match | Low |

Your own email address is automatically excluded from attendee matching (configure under **Settings → Calendar → My email address**).

Each result shows which signals matched, so you can see exactly why an email was surfaced.

---

## Python 3.14 / Restricted Environments

`sentence-transformers` (and `torch`) do not yet have pre-built wheels for Python 3.14. The app handles this gracefully:

- **FTS search** works normally without any ML packages.
- **Semantic / hybrid search** modes are hidden if `sentence-transformers` is unavailable.
- **NLP auto-classification** falls back to TF-IDF (no ML dependencies).
- **Keyword extraction and NER** are silently skipped if the spaCy model is missing.

### SSL / cert-restricted networks

If `pip install` or model downloads fail due to SSL errors, pre-packaged model bundles are on [GitHub Releases (models-v1)](https://github.com/shawntanzk/eml-search/releases/tag/models-v1). `setup_models.py` will fall back to these automatically.

```bash
# spaCy model (12 MB):
tar -xzf en_core_web_sm-3.8.0.tar.gz -C .venv/lib/python3.*/site-packages/

# Sentence-transformer model (80 MB):
tar -xzf all-MiniLM-L6-v2.tar.gz -C ~/.cache/huggingface/hub/
```

---

## Project Structure

```
eml_search_app/
├── app.py                   # Streamlit UI — Search, Tags, Knowledge Graph, Calendar, Settings
├── config.py                # Paths, model names, thresholds
├── setup_models.py          # One-time setup: models → database → index
├── requirements.txt
│
├── modules/
│   ├── eml_parser.py        # .eml file → structured dict (stdlib only)
│   ├── indexer.py           # SQLite/FTS5 CRUD
│   ├── watcher.py           # Background folder watcher + indexing pipeline
│   ├── imap_connector.py    # IMAP ingestion — fetch from live server
│   ├── calendar_reader.py   # Calendar JSON loader, HTML/Streamlit renderer, email correlator
│   ├── calendar_online.py   # Live calendar fetch — iCal URL + Microsoft Graph API
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
│   └── settings.json        # Persisted settings incl. credentials
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
                        ▼
                  structured dict
                  {id, subject, sender, body_text, …}
                        │
                        ▼
1. indexer.insert_email()          — SQLite + FTS5 trigger
2. nlp_engine.extract_entities()   — NER [optional]
3. semantic_search.embed_batch()   — embeddings [optional]
```

### Search

| Mode | How |
|---|---|
| **Full-text** | SQLite FTS5 + Porter stemmer, BM25 ranking |
| **Semantic** | Query embedded; dot product against embedding matrix |
| **Hybrid** | Reciprocal Rank Fusion of FTS + semantic rankings |

### Calendar

Events from any source (JSON / iCal URL / Graph API) are normalised to the same internal format. Times are stored as UTC and converted to your chosen display timezone. The event correlator uses a multi-signal RRF approach to surface related emails.

---

## Configuration

All tuneable values in `config.py`:

| Variable | Default | Description |
|---|---|---|
| `SPACY_MODEL` | `en_core_web_sm` | spaCy model name |
| `SENTENCE_TRANSFORMER_MODEL` | `all-MiniLM-L6-v2` | Embedding model |
| `SEMANTIC_TOP_K` | `100` | Semantic search candidates |
| `MAX_SEARCH_RESULTS` | `200` | Hard cap on results |
| `WATCH_POLL_INTERVAL` | `10` | Seconds between folder polls |

---

## Requirements

Core (always required):
```
streamlit>=1.32.0
rdflib>=7.0.0
pandas>=2.0.0
pyvis>=0.3.2
numpy>=1.24.0
requests>=2.28.0
msal>=1.20.0
icalendar>=5.0.0
```

Optional (app degrades gracefully without these):
```
spacy>=3.7.0                 # NER and keyword extraction
sentence-transformers>=2.7.0 # Semantic/hybrid search and NLP tag classification
```

---

## Credential Storage

All credentials (IMAP password, OAuth2 tokens, iCal app password, Graph tokens) are saved to `data/settings.json` on your local machine. This file is listed in `.gitignore` and never leaves your machine.
