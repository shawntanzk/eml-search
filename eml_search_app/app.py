"""EML Search — Streamlit app entry point."""
import sys
from pathlib import Path

import threading
from typing import Optional
import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR))

import config
from modules import indexer, nlp_engine, search_engine, semantic_search, tagger
from modules import calendar_reader
from modules.watcher import EmailWatcher
from modules.imap_connector import IMAPConnector, MICROSOFT_AUTHORITY, OUTLOOK_IMAP_SCOPE
from modules.graph_builder import (
    build_abox, save_abox, load_abox, get_merged_graph,
    sparql_query, get_graph_stats, get_all_graph_nodes, get_subgraph,
    get_paths_between_seeds, inline_graph_assets,
)

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="EML Search",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)

indexer.init_db()


@st.cache_resource
def _get_watcher(folder: str) -> EmailWatcher:
    w = EmailWatcher(folder)
    w.start()
    return w


def _current_folder() -> str:
    return config.load_settings().get("email_folder", config.DEFAULT_EMAIL_FOLDER)


@st.cache_resource
def _get_imap_poller(host: str, username: str, interval: int, sync_deletions: bool) -> IMAPConnector:
    """
    Cached IMAP background poller — one instance per (host, username, interval, sync_deletions).
    Starts a daemon thread that checks for new emails every *interval* seconds.
    The token_save_callback persists auto-refreshed OAuth2 tokens to settings.json.
    """
    imap_cfg = config.load_settings().get("imap", {})

    def _save_tokens(new_tokens: dict) -> None:
        s = config.load_settings()
        s.setdefault("imap", {})
        s["imap"]["access_token"] = new_tokens["access_token"]
        s["imap"]["refresh_token"] = new_tokens["refresh_token"]
        config.save_settings(s)

    if imap_cfg.get("use_oauth2"):
        connector = IMAPConnector(
            host=host,
            username=username,
            access_token=imap_cfg.get("access_token", ""),
            refresh_token=imap_cfg.get("refresh_token", ""),
            client_id=imap_cfg.get("client_id", ""),
            token_save_callback=_save_tokens,
        )
    else:
        connector = IMAPConnector(
            host=host,
            username=username,
            password=imap_cfg.get("password", ""),
            port=int(imap_cfg.get("port", 993)),
            use_ssl=bool(imap_cfg.get("use_ssl", True)),
        )

    connector.start(
        mailbox=imap_cfg.get("mailbox", "INBOX"),
        interval=interval,
        sync_deletions=sync_deletions,
    )
    return connector


def _maybe_start_imap_poller() -> Optional[IMAPConnector]:
    """Start the IMAP poller if settings are configured, otherwise return None."""
    imap_cfg = config.load_settings().get("imap", {})
    oauth2_ready = imap_cfg.get("use_oauth2") and bool(imap_cfg.get("access_token"))
    basic_ready = (
        not imap_cfg.get("use_oauth2")
        and all(imap_cfg.get(k) for k in ("host", "username", "password"))
    )
    if not (oauth2_ready or basic_ready):
        return None
    host = imap_cfg.get("host", "")
    username = imap_cfg.get("username", "")
    interval = int(imap_cfg.get("poll_interval", 300))
    sync_deletions = bool(imap_cfg.get("sync_deletions", True))
    return _get_imap_poller(host, username, interval, sync_deletions)


_settings = config.load_settings()
_mode = _settings.get("mode", "offline")   # "offline" | "online"

_watcher = _get_watcher(_current_folder())
_imap_poller = _maybe_start_imap_poller() if _mode == "online" else None

# ── Tab-switch helper (must run before tabs are rendered) ─────────────────────
# When SPARQL navigation buttons set switch_to_search=True, inject JS that
# clicks the Search tab, then clear the flag so it doesn't fire again.
def _tab_switch_js(tab_label: str) -> str:
    return f"""
    <script>
    setTimeout(function () {{
        var tabs = window.parent.document.querySelectorAll('[data-baseweb="tab"]');
        for (var i = 0; i < tabs.length; i++) {{
            if (tabs[i].textContent.trim() === "{tab_label}") {{
                tabs[i].click(); break;
            }}
        }}
    }}, 250);
    </script>
    """

if st.session_state.pop("switch_to_search", False):
    components.html(_tab_switch_js("Search"), height=0)

if st.session_state.pop("switch_to_tags", False):
    components.html(_tab_switch_js("Tags"), height=0)


def _nav_to_search(query: str) -> None:
    """Set the search query and trigger a tab switch to Search."""
    st.session_state["_nav_query"] = query
    st.session_state["switch_to_search"] = True
    st.rerun()


def _nav_to_tags(tag_name: str) -> None:
    """Switch to the Tags tab and pre-select a tag in Browse by tag."""
    st.session_state["_direct_tag_name"] = tag_name
    st.session_state["switch_to_tags"] = True
    st.rerun()


def _nav_to_email(email_id: str) -> None:
    """Navigate directly to a specific email by ID (from Knowledge Graph)."""
    st.session_state["_direct_email_id"] = email_id
    st.session_state["switch_to_search"] = True
    st.rerun()


def _render_email_detail(em: dict, all_tags: list) -> None:
    """Render the full detail view for a single email (body, tags, keywords)."""
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Subject:** {em.get('subject', '')}")
        st.write(
            f"**From:** {em.get('sender_name', '')} "
            f"&lt;{em.get('sender_email', '')}&gt;"
        )
        to_list = ", ".join(
            f"{p['name']} <{p['email']}>"
            for p in (em.get("recipients") or [])
        )
        st.write(f"**To:** {to_list}")
    with c2:
        st.write(f"**Date:** {em.get('date', '')}")
        if em.get("attachment_names"):
            st.write(f"**Attachments:** {', '.join(em['attachment_names'])}")

    body_col, kw_col = st.columns([3, 1])
    with body_col:
        st.text_area(
            "Body",
            value=em.get("body_text", "")[:3000],
            height=200,
            disabled=True,
            key=f"body_{em['id']}",
        )
    with kw_col:
        keywords = nlp_engine.extract_keywords(em.get("body_text", ""))
        if keywords:
            st.write("**Keywords**")
            st.write(", ".join(keywords))
        elif not nlp_engine.NLP_AVAILABLE():
            st.caption("Keywords unavailable — spaCy model not installed.")

    st.divider()
    st.write("**Tags**")

    current_tags = tagger.get_email_tags(em["id"])
    assigned_ids = {t["id"] for t in current_tags}

    tag_cols = st.columns(min(len(current_tags) + 1, 6))
    for i, t in enumerate(current_tags):
        with tag_cols[i % 6]:
            source_icon = "👤" if t["source"] == "manual" else "🤖"
            if st.button(
                f"{source_icon} {t['name']} ✕",
                key=f"rm_{em['id']}_{t['id']}",
                help="Click to remove this tag",
            ):
                tagger.remove_tag_manual(em["id"], t["id"])
                st.rerun()

    if not current_tags:
        st.caption("No tags assigned.")

    available = [t for t in all_tags if t["id"] not in assigned_ids]
    if available:
        add_col, btn_col = st.columns([3, 1])
        with add_col:
            chosen = st.selectbox(
                "Add tag",
                options=[""] + [t["name"] for t in available],
                label_visibility="collapsed",
                key=f"add_tag_select_{em['id']}",
            )
        with btn_col:
            if st.button("Add", key=f"add_tag_btn_{em['id']}") and chosen:
                tag_id = next(
                    t["id"] for t in available if t["name"] == chosen
                )
                tagger.assign_tag_manual(em["id"], tag_id)
                st.rerun()
    elif not all_tags:
        st.caption("Create tags in the **Tags** tab first.")


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📧 EML Search")

    # Mode toggle
    _mode_choice = st.radio(
        "Mode",
        ["Offline", "Online"],
        index=0 if _mode == "offline" else 1,
        horizontal=True,
        key="mode_toggle",
        help="Offline: index local .eml files  |  Online: connect via IMAP",
    )
    if _mode_choice.lower() != _mode:
        _settings["mode"] = _mode_choice.lower()
        config.save_settings(_settings)
        st.rerun()

    st.divider()
    total = indexer.get_email_count()
    st.metric("Emails indexed", total)

    if _mode == "offline":
        st.caption(f"Folder watcher: {_watcher.status}")
    else:
        if _imap_poller:
            _del_hint = f" · {_imap_poller.last_deleted:,} deleted" if _imap_poller.last_deleted else ""
            st.caption(f"IMAP poller: {_imap_poller.status}{_del_hint}")
        else:
            st.caption("IMAP poller: not configured")

    st.divider()
    with st.expander("NLP diagnostics", expanded=False):
        _nlp_ok = nlp_engine.NLP_AVAILABLE()
        _sem_ok, _sem_err = semantic_search.model_status()
        _total = indexer.get_email_count()
        _embedded = indexer.get_embedding_count()

        if _nlp_ok:
            st.success("spaCy model loaded")
        else:
            st.error(f"spaCy: {nlp_engine.NLP_ERROR()}")

        if _sem_ok:
            st.success(f"Sentence-transformer loaded")
            if _embedded < _total:
                st.warning(f"Embeddings: {_embedded}/{_total} — {_total - _embedded} email(s) missing.")
                if st.button("Generate missing embeddings", key="fill_embeddings"):
                    missing = indexer.get_emails_without_embeddings()
                    progress = st.progress(0, text="Embedding emails…")
                    for i, em in enumerate(missing):
                        try:
                            text = f"{em.get('subject', '')} {(em.get('body_text') or '')[:400]}"
                            vec = semantic_search.embed_text(text)
                            indexer.insert_embedding(em["id"], vec)
                        except Exception:
                            pass
                        progress.progress((i + 1) / len(missing), text=f"Embedding {i + 1}/{len(missing)}…")
                    progress.empty()
                    st.success(f"Done — {len(missing)} embedding(s) generated.")
                    st.rerun()
            else:
                st.success(f"Embeddings: {_embedded}/{_total}")
        else:
            st.error(f"Sentence-transformer: {_sem_err}")

    st.divider()
    tag_counts = tagger.get_tag_counts()
    tag_options = {f"{t['name']} ({t['count']})": t["id"] for t in tag_counts}

    st.subheader("Filters")
    f_sender      = st.text_input("Sender contains", key="f_sender")
    f_date_from   = st.date_input("Date from", value=None, key="f_date_from")
    f_date_to     = st.date_input("Date to", value=None, key="f_date_to")
    f_attachments = st.checkbox("Has attachments only", key="f_attachments")
    f_tag_label   = st.selectbox(
        "Tag", options=["(any)"] + list(tag_options.keys()), key="f_tag"
    )

    filters = {
        "sender":          f_sender or None,
        "date_from":       str(f_date_from) if f_date_from else None,
        "date_to":         str(f_date_to) if f_date_to else None,
        "has_attachments": f_attachments or None,
        "tag_id":          tag_options.get(f_tag_label) if f_tag_label != "(any)" else None,
    }

# ── Tabs ─────────────────────────────────────────────────────────────────────
tab_search, tab_tags, tab_graph, tab_calendar, tab_settings = st.tabs(
    ["Search", "Tags", "Knowledge Graph", "Calendar", "Settings"]
)

# ════════════════════════════════════════════════════════════════════════════
# SEARCH TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_search:
    # Direct email view: navigated here from Knowledge Graph "Open email"
    if "_direct_email_id" in st.session_state:
        direct_id = st.session_state["_direct_email_id"]
        if st.button("← Back to search", key="back_to_search"):
            del st.session_state["_direct_email_id"]
            st.rerun()
        em = indexer.get_email_by_id(direct_id)
        if em:
            all_tags = tagger.get_all_tags()
            with st.container(border=True):
                st.subheader(em.get("subject") or "(no subject)")
                st.caption(
                    f"From: {em.get('sender_name', '')} <{em.get('sender_email', '')}>"
                    f"  ·  {(em.get('date') or '')[:16]}"
                    + ("  ·  📎" if em.get("has_attachments") else "")
                )
                st.divider()
                _render_email_detail(em, all_tags)
        else:
            st.warning("Email not found.")
    else:
        # Normal search UI
        if "_nav_query" in st.session_state:
            st.session_state["search_query"] = st.session_state.pop("_nav_query")

        col_q, col_mode = st.columns([5, 1])
        with col_q:
            query = st.text_input(
                "Search emails",
                placeholder="Enter keywords, names, phrases…",
                label_visibility="collapsed",
                key="search_query",
            )
        with col_mode:
            _modes = ["hybrid", "fts", "semantic"] if semantic_search.model_status()[0] else ["fts"]
            mode = st.selectbox(
                "Mode", _modes,
                label_visibility="collapsed",
                key="search_mode",
            )

        results = search_engine.search(query, mode=mode, filters=filters, limit=50)

        if total == 0:
            st.info("No emails indexed yet. Go to **Settings** to point the app at your EML folder.")
        elif not results and query:
            st.warning("No results found.")
        else:
            if query:
                st.caption(f"{len(results)} result(s) — mode: **{mode}**")

            if "open_email" not in st.session_state:
                st.session_state.open_email = None

            all_tags = tagger.get_all_tags()  # for the add-tag dropdown

            for r in results:
                with st.container(border=True):
                    hdr, _ = st.columns([4, 1])
                    with hdr:
                        clicked = st.button(
                            f"**{r.get('subject') or '(no subject)'}**",
                            key=f"btn_{r['id']}",
                            use_container_width=True,
                        )
                        if clicked:
                            st.session_state.open_email = (
                                None if st.session_state.open_email == r["id"] else r["id"]
                            )
                        st.caption(
                            f"From: {r.get('sender_name', '')} <{r.get('sender_email', '')}>"
                            f"  ·  {(r.get('date') or '')[:16]}"
                            + ("  ·  📎" if r.get("has_attachments") else "")
                        )

                    if st.session_state.open_email == r["id"]:
                        em = indexer.get_email_by_id(r["id"])
                        if em:
                            st.divider()
                            _render_email_detail(em, all_tags)


# ════════════════════════════════════════════════════════════════════════════
# TAGS TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_tags:

    # ── Tag Library ──────────────────────────────────────────────────────────
    st.subheader("Tag library")
    st.caption("Tags are human-defined categories. NLP classification will never remove them.")

    all_tags = tagger.get_all_tags()

    if all_tags:
        tag_counts_map = {t["id"]: t["count"] for t in tagger.get_tag_counts()}
        cols = st.columns(min(len(all_tags), 5))
        for i, t in enumerate(all_tags):
            with cols[i % 5]:
                count = tag_counts_map.get(t["id"], 0)
                if st.button(
                    f"🏷 {t['name']}  ({count})  ✕",
                    key=f"del_tag_{t['id']}",
                    help="Click to delete this tag and all its assignments",
                    use_container_width=True,
                ):
                    tagger.delete_tag(t["id"])
                    st.rerun()
    else:
        st.info("No tags defined yet. Add your first tag below.")

    st.divider()
    new_tag_col, add_col = st.columns([4, 1])
    with new_tag_col:
        new_tag_name = st.text_input(
            "New tag name",
            placeholder="e.g. Invoice, Meeting, Urgent, Project Alpha…",
            label_visibility="collapsed",
            key="new_tag_input",
        )
    with add_col:
        if st.button("Add tag", type="primary", key="add_tag_btn"):
            if new_tag_name.strip():
                tagger.add_tag(new_tag_name.strip())
                st.rerun()
            else:
                st.warning("Enter a tag name first.")

    # ── NLP Classification ───────────────────────────────────────────────────
    st.divider()
    st.subheader("NLP auto-classification")
    st.caption(
        "Each tag has its own classification method and threshold. "
        "Only **adds** tags — never removes. Manually removed tags are permanently "
        "blocked from being re-added by NLP for that email."
    )

    if not all_tags:
        st.warning("Add at least one tag above before running classification.")
    else:
        _semantic_ok = semantic_search.model_status()[0]
        _method_options = []
        if _semantic_ok:
            _method_options.append("Semantic (sentence-transformers)")
        _method_options.append("TF-IDF (no ML dependencies)")

        # "Classify all tags" button at the top
        if st.button("Classify all tags", type="primary", key="run_nlp_all"):
            with st.spinner("Classifying all tags…"):
                result = tagger.classify_all_tags()
            st.success(
                f"Done — **{result['new_assignments']}** new tag assignment(s)."
            )
            st.rerun()

        st.write("**Per-tag settings**")

        for _tag in all_tags:
            _tid = _tag["id"]
            _saved_method = _tag.get("nlp_method") or "tfidf"
            _saved_threshold = _tag.get("nlp_threshold") or 0.15

            # Map DB value → display label
            if _saved_method == "semantic" and _semantic_ok:
                _method_idx = 0
            else:
                _method_idx = len(_method_options) - 1  # TF-IDF is always last

            with st.expander(f"🏷 {_tag['name']}", expanded=False):
                _m_col, _t_col, _btn_col = st.columns([3, 3, 2])

                with _m_col:
                    _chosen_method_label = st.radio(
                        "Method",
                        _method_options,
                        index=_method_idx,
                        key=f"tag_method_{_tid}",
                        horizontal=True,
                    )

                _use_tfidf_tag = _chosen_method_label.startswith("TF-IDF")
                _db_method = "tfidf" if _use_tfidf_tag else "semantic"
                _default_thresh = 0.15 if _use_tfidf_tag else 0.25
                _max_thresh = 0.50 if _use_tfidf_tag else 0.60
                # Clamp saved threshold into the current method's range
                _init_thresh = float(
                    max(0.05, min(_saved_threshold, _max_thresh))
                    if _saved_method == _db_method
                    else _default_thresh
                )

                with _t_col:
                    _chosen_threshold = st.slider(
                        "Threshold",
                        min_value=0.05,
                        max_value=_max_thresh,
                        value=_init_thresh,
                        step=0.05,
                        key=f"tag_threshold_{_tid}",
                        help="Higher = fewer but more confident assignments.",
                    )

                # Persist settings whenever they change
                if _db_method != _saved_method or abs(_chosen_threshold - _saved_threshold) > 1e-6:
                    indexer.save_tag_nlp_settings(_tid, _db_method, _chosen_threshold)

                with _btn_col:
                    st.write("")  # vertical alignment nudge
                    if st.button("Classify", key=f"run_nlp_{_tid}"):
                        # Save current UI values before classifying
                        indexer.save_tag_nlp_settings(_tid, _db_method, _chosen_threshold)
                        with st.spinner(f"Classifying '{_tag['name']}'…"):
                            _result = tagger.classify_tag(_tid)
                        st.success(
                            f"**{_result['new_assignments']}** new assignment(s)."
                        )
                        st.rerun()

    # ── Browse by tag ─────────────────────────────────────────────────────────
    if all_tags:
        st.divider()
        st.subheader("Browse by tag")

        tag_name_map = {t["name"]: t["id"] for t in all_tags}

        # Pre-select a tag when navigating here from the Knowledge Graph
        _direct_tag = st.session_state.pop("_direct_tag_name", None)
        _default_tag_idx = 0
        if _direct_tag and _direct_tag in tag_name_map:
            _default_tag_idx = list(tag_name_map.keys()).index(_direct_tag)

        chosen_tag = st.selectbox(
            "Select tag",
            options=list(tag_name_map.keys()),
            index=_default_tag_idx,
            key="browse_tag_select",
        )
        if chosen_tag:
            tag_id = tag_name_map[chosen_tag]
            tagged_emails = tagger.get_emails_by_tag(tag_id)

            st.caption(f"{len(tagged_emails)} email(s) tagged **{chosen_tag}**")

            for em in tagged_emails:
                source_badge = "👤 manual" if em["source"] == "manual" else "🤖 NLP"
                with st.container(border=True):
                    src_col, info_col = st.columns([1, 5])
                    with src_col:
                        st.caption(source_badge)
                    with info_col:
                        st.write(f"**{em.get('subject') or '(no subject)'}**")
                        st.caption(
                            f"From: {em.get('sender_name', '')} <{em.get('sender_email', '')}>"
                            f"  ·  {(em.get('date') or '')[:16]}"
                            + ("  ·  📎" if em.get("has_attachments") else "")
                        )

                    rm_col, _ = st.columns([1, 5])
                    with rm_col:
                        if st.button("Remove tag", key=f"browse_rm_{tag_id}_{em['id']}"):
                            tagger.remove_tag_manual(em["id"], tag_id)
                            st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# KNOWLEDGE GRAPH TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_graph:
    st.subheader("RDF/OWL Knowledge Graph")

    abox_path = Path(config.GRAPH_DATA_PATH)
    stats_col, build_col = st.columns([3, 1])

    with build_col:
        if st.button("Build / Rebuild Graph", type="primary", key="build_graph"):
            with st.spinner("Building ABox from indexed data…"):
                try:
                    all_ids = indexer.get_all_email_ids()
                    emails_data = [indexer.get_email_by_id(eid) for eid in all_ids]
                    emails_data = [e for e in emails_data if e]

                    conn = indexer._get_conn()
                    ent_rows = conn.execute(
                        "SELECT email_id, entity_text, entity_label FROM email_entities"
                    ).fetchall()
                    entities_map: dict[str, list] = {}
                    for row in ent_rows:
                        entities_map.setdefault(row["email_id"], []).append(
                            {"text": row["entity_text"], "label": row["entity_label"]}
                        )

                    tag_rows = conn.execute(
                        "SELECT et.email_id, t.name FROM email_tags et JOIN tags t ON t.id = et.tag_id"
                    ).fetchall()
                    tags_map: dict[str, list] = {}
                    for row in tag_rows:
                        tags_map.setdefault(row["email_id"], []).append(row["name"])

                    g = build_abox(emails_data, entities_map, tags_map)
                    save_abox(g)
                    st.success(f"Graph built — {len(g)} triples.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Graph build failed: {exc}")

    if not abox_path.exists():
        st.info("No graph data yet. Click **Build / Rebuild Graph** to generate it.")
    else:
        try:
            abox = load_abox()
        except Exception as exc:
            st.error(f"Failed to load graph: {exc}")
            abox = None

        if abox is not None:
            stats = get_graph_stats(abox)
            with stats_col:
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Triples",       stats["triples"])
                c2.metric("Emails",        stats["emails"])
                c3.metric("Persons",       stats["persons"])
                c4.metric("Organisations", stats["organizations"])
                c5.metric("Tags",          stats["topics"])

            # ── Node selection ────────────────────────────────────────────
            st.divider()
            st.subheader("Interactive graph")
            st.caption(
                "Search for nodes and add them as seeds one at a time. "
                "The graph shows each seed and its directly connected neighbours."
            )

            all_nodes = get_all_graph_nodes(abox)

            # Persistent seed set: {label: uri}
            if "_graph_seeds" not in st.session_state:
                st.session_state._graph_seeds = {}

            # Type filter checkboxes
            type_names = ["Email", "Person", "Organization", "Tag", "Thread", "Location"]
            type_cols = st.columns(len(type_names))
            allowed_types: set[str] = set()
            for i, tname in enumerate(type_names):
                with type_cols[i]:
                    if st.checkbox(tname, value=True, key=f"gt_{tname}"):
                        allowed_types.add(tname)

            # Search + add
            search_col, add_col = st.columns([4, 1])
            with search_col:
                node_search = st.text_input(
                    "Search nodes",
                    placeholder="Type a name to search…",
                    key="node_search",
                    label_visibility="collapsed",
                )
            filtered = [
                n for n in all_nodes
                if (not node_search or node_search.lower() in n["label"].lower())
                and n["type"] in allowed_types
            ]
            option_labels = [f"[{n['type']}] {n['label']}" for n in filtered]
            label_to_uri  = {f"[{n['type']}] {n['label']}": n["uri"] for n in all_nodes}

            pick = st.selectbox(
                f"Matching nodes ({len(filtered)} found)",
                options=[""] + option_labels,
                key="node_pick",
                label_visibility="collapsed",
            )
            with add_col:
                if st.button("Add seed", key="add_seed", disabled=not pick):
                    st.session_state._graph_seeds[pick] = label_to_uri[pick]
                    st.rerun()

            # Show selected seeds with remove buttons
            seeds: dict = st.session_state._graph_seeds
            if seeds:
                st.write("**Selected seeds:**")
                seed_cols = st.columns(min(len(seeds), 4))
                for i, label in enumerate(list(seeds)):
                    with seed_cols[i % 4]:
                        if st.button(f"✕ {label}", key=f"rm_seed_{i}", help="Remove this seed"):
                            del st.session_state._graph_seeds[label]
                            st.rerun()
                if st.button("Clear all seeds", key="clear_seeds"):
                    st.session_state._graph_seeds = {}
                    st.rerun()
            else:
                st.info("Search for a node above and click **Add seed** to start.")

            selected_labels = list(seeds.keys())

            if not selected_labels:
                pass
            else:
                # ── Render mode ───────────────────────────────────────────
                st.divider()
                _has_multi_seeds = len(selected_labels) >= 2
                render_mode = st.radio(
                    "Render mode",
                    ["Neighbourhood", "Paths between seeds"],
                    key="render_mode",
                    horizontal=True,
                    help=(
                        "**Neighbourhood**: show each seed and its direct neighbours.\n\n"
                        "**Paths between seeds**: show only nodes that lie on a connecting "
                        "path between any two seeds (multi-hop)."
                    ),
                    disabled=not _has_multi_seeds,
                )
                if not _has_multi_seeds and render_mode == "Paths between seeds":
                    render_mode = "Neighbourhood"

                _use_paths = render_mode == "Paths between seeds"
                if _use_paths:
                    max_hops = st.slider(
                        "Max hops",
                        min_value=1,
                        max_value=6,
                        value=3,
                        key="max_hops",
                        help="Maximum path length between any two seeds. Higher values reveal longer chains but may add many nodes.",
                    )
                else:
                    max_hops = 1

                if st.button("Render graph", type="primary", key="render_graph"):
                    seed_uris = list(seeds.values())
                    with st.spinner("Rendering…"):
                        try:
                            from pyvis.network import Network

                            if _use_paths:
                                nodes, edges = get_paths_between_seeds(
                                    abox, seed_uris, allowed_types, max_hops=max_hops
                                )
                            else:
                                nodes, edges = get_subgraph(abox, seed_uris, allowed_types)

                            net = Network(
                                height="620px", width="100%",
                                bgcolor="#ffffff", font_color="#000000",
                            )
                            net.set_options("""{
                              "physics": {"stabilization": {"iterations": 200}},
                              "nodes": {
                                "font": {"size": 13, "color": "#000000"},
                                "borderWidth": 2
                              },
                              "edges": {
                                "font": {"size": 11, "color": "#333333", "align": "middle"},
                                "arrows": {"to": {"enabled": true, "scaleFactor": 0.8}},
                                "smooth": {"type": "curvedCW", "roundness": 0.2}
                              }
                            }""")

                            for n in nodes:
                                is_seed = n["id"] in set(seed_uris)
                                net.add_node(
                                    n["id"],
                                    label=n["label"],
                                    color=n["color"],
                                    title=f"{n['type']}: {n['label']}",
                                    size=20 if is_seed else 14,
                                    borderWidth=3 if is_seed else 2,
                                )
                            for e in edges:
                                net.add_edge(
                                    e["from"], e["to"],
                                    label=e["label"],
                                    title=e["label"],
                                )

                            html_path = str(Path(config.DATA_DIR) / "graph_preview.html")
                            net.save_graph(html_path)
                            html_content = inline_graph_assets(
                                open(html_path, encoding="utf-8").read()
                            )
                            open(html_path, "w", encoding="utf-8").write(html_content)
                            components.html(html_content, height=640, scrolling=False)
                            st.caption(
                                f"{len(nodes)} node(s), {len(edges)} edge(s). "
                                "Seed nodes are shown larger with a thicker border."
                            )
                            st.session_state["_graph_rendered_nodes"] = nodes
                        except ImportError:
                            st.error("pyvis not installed. Run: pip install pyvis")
                        except Exception as exc:
                            st.error(f"Render failed: {exc}")

            # ── Navigate from graph nodes ─────────────────────────────────
            rendered_nodes: list[dict] = st.session_state.get("_graph_rendered_nodes", [])
            if rendered_nodes:
                st.divider()
                with st.expander("Navigate from graph nodes", expanded=False):
                    st.caption(
                        "Jump to any node from the last rendered graph. "
                        "Emails open directly; people, organisations, and locations "
                        "search by name; tags switch to the Tags tab."
                    )
                    by_type: dict[str, list[dict]] = {}
                    for n in rendered_nodes:
                        by_type.setdefault(n["type"], []).append(n)

                    _TYPE_ORDER = ["Email", "Person", "Organization", "Tag", "Location", "Thread"]
                    for tname in _TYPE_ORDER:
                        group = by_type.get(tname, [])
                        if not group:
                            continue
                        st.write(f"**{tname}s**")
                        btn_cols = st.columns(min(len(group), 3))
                        for i, n in enumerate(sorted(group, key=lambda x: x["label"].lower())):
                            with btn_cols[i % 3]:
                                if tname == "Email":
                                    email_id = n["id"].split("#email_", 1)[-1] if "#email_" in n["id"] else None
                                    if email_id and st.button(
                                        n["label"] or email_id,
                                        key=f"gnav_{n['id']}",
                                        use_container_width=True,
                                        help="Open this email",
                                    ):
                                        _nav_to_email(email_id)
                                elif tname == "Tag":
                                    if st.button(
                                        n["label"],
                                        key=f"gnav_{n['id']}",
                                        use_container_width=True,
                                        help="Browse emails with this tag",
                                    ):
                                        _nav_to_tags(n["label"])
                                else:
                                    if st.button(
                                        n["label"],
                                        key=f"gnav_{n['id']}",
                                        use_container_width=True,
                                        help=f"Search for {n['label']}",
                                    ):
                                        _nav_to_search(n["label"])

            # ── SPARQL ────────────────────────────────────────────────────
            st.divider()
            st.subheader("SPARQL query")
            st.caption(
                "Click **Search** / **Open email** on any result row to jump to that "
                "item in the Search tab."
            )

            default_sparql = """PREFIX ont: <http://emailsearch.local/ontology#>
PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>

SELECT ?email ?subject ?sender
WHERE {
  ?email rdf:type ont:Email ;
         ont:hasSubject ?subject ;
         ont:hasSender ?person .
  ?person ont:emailAddress ?sender .
}
LIMIT 20"""
            sparql_input = st.text_area(
                "SPARQL SELECT query", value=default_sparql, height=180, key="sparql_input"
            )

            if st.button("Run query", key="run_sparql"):
                with st.spinner("Querying…"):
                    try:
                        merged = get_merged_graph()
                        rows = sparql_query(merged, sparql_input)
                        st.session_state["sparql_results"] = rows
                    except Exception as exc:
                        st.error(f"Query error: {exc}")
                        st.session_state["sparql_results"] = []

            sparql_rows = st.session_state.get("sparql_results")
            if sparql_rows is not None:
                if not sparql_rows:
                    st.info("No results.")
                else:
                    import re as _re

                    def _email_id_from_uri(val: str) -> str | None:
                        m = _re.search(r"data#email_([a-f0-9]+)", val)
                        return m.group(1) if m else None

                    def _is_email_address(val: str) -> bool:
                        return bool(_re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", val))

                    st.caption(f"{len(sparql_rows)} result(s)")

                    for row_i, row in enumerate(sparql_rows):
                        with st.container(border=True):
                            val_cols = st.columns(len(row))
                            for col_i, (key, val) in enumerate(row.items()):
                                with val_cols[col_i]:
                                    st.caption(key)
                                    st.write(val if len(val) <= 80 else val[:77] + "…")

                                    email_id = _email_id_from_uri(val)
                                    if email_id:
                                        if st.button(
                                            "Open email",
                                            key=f"sparql_nav_{row_i}_{col_i}",
                                        ):
                                            _nav_to_email(email_id)
                                    elif _is_email_address(val):
                                        if st.button(
                                            "Search sender",
                                            key=f"sparql_nav_{row_i}_{col_i}",
                                        ):
                                            _nav_to_search(val)
                                    elif (
                                        not val.startswith("http")
                                        and len(val) > 3
                                    ):
                                        if st.button(
                                            "Search",
                                            key=f"sparql_nav_{row_i}_{col_i}",
                                        ):
                                            _nav_to_search(val)


# ════════════════════════════════════════════════════════════════════════════
# CALENDAR TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_calendar:
    import calendar as _calendar_mod
    from datetime import date as _date, timedelta as _timedelta

    _cal_settings = config.load_settings()
    _cal_json_path = _cal_settings.get("calendar_json_path", "")

    # ── Path setup ────────────────────────────────────────────────────────────
    if not _cal_json_path:
        st.info(
            "📅 **Calendar not configured.**  \n"
            "Go to **Settings → Calendar** and enter the path to your events JSON file."
        )
    else:
        # ── Load events ───────────────────────────────────────────────────────
        _cal_display_tz = _cal_settings.get("calendar_display_tz", "Asia/Singapore")
        _cal_events_raw = calendar_reader.load_events(_cal_json_path)
        _cal_events = calendar_reader.convert_display_tz(_cal_events_raw, _cal_display_tz)
        _today = _date.today()

        if not _cal_events:
            st.warning(
                f"No events loaded from `{_cal_json_path}`. "
                "Check the path and that the file is valid JSON."
            )
        else:
            # ── Summary metrics ───────────────────────────────────────────────
            _week_start = _today - _timedelta(days=_today.weekday())
            _week_end   = _week_start + _timedelta(days=6)
            _this_week  = calendar_reader.events_in_range(_cal_events, _week_start, _week_end)
            _upcoming   = [
                e for e in _cal_events
                if e["start_dt"] and e["start_dt"].date() >= _today
            ]
            _past = [
                e for e in _cal_events
                if e["start_dt"] and e["start_dt"].date() < _today
            ]
            _m1, _m2, _m3, _m4 = st.columns(4)
            _m1.metric("Total events", len(_cal_events))
            _m2.metric("This week",    len(_this_week))
            _m3.metric("Upcoming",     len(_upcoming))
            _m4.metric("Past",         len(_past))
            st.caption(f"🌐 Displaying in **{_cal_display_tz}**  ·  Change in Settings → Calendar")

            st.divider()

            # ── View selector ─────────────────────────────────────────────────
            _view_tab_cal, _view_tab_week, _view_tab_list = st.tabs(["📅 Month", "📆 Week", "📋 List"])

            # ── MONTH VIEW ────────────────────────────────────────────────────
            with _view_tab_cal:
                # Month navigation state
                if "cal_year"  not in st.session_state:
                    st.session_state.cal_year  = _today.year
                if "cal_month" not in st.session_state:
                    st.session_state.cal_month = _today.month

                _nav_l, _nav_title, _nav_today, _nav_r = st.columns([1, 4, 1.2, 1])
                with _nav_l:
                    if st.button("◀", key="cal_prev_month"):
                        if st.session_state.cal_month == 1:
                            st.session_state.cal_month = 12
                            st.session_state.cal_year -= 1
                        else:
                            st.session_state.cal_month -= 1
                        st.rerun()
                with _nav_title:
                    st.subheader(
                        f"{_calendar_mod.month_name[st.session_state.cal_month]} "
                        f"{st.session_state.cal_year}"
                    )
                with _nav_today:
                    if st.button("Today", key="cal_goto_today"):
                        st.session_state.cal_year  = _today.year
                        st.session_state.cal_month = _today.month
                        st.rerun()
                with _nav_r:
                    if st.button("▶", key="cal_next_month"):
                        if st.session_state.cal_month == 12:
                            st.session_state.cal_month = 1
                            st.session_state.cal_year += 1
                        else:
                            st.session_state.cal_month += 1
                        st.rerun()

                # HTML calendar grid (visual only)
                _month_html, _month_height = calendar_reader.render_month_html(
                    st.session_state.cal_year,
                    st.session_state.cal_month,
                    _cal_events,
                )
                components.html(_month_html, height=_month_height, scrolling=False)

                # Event picker for this month
                _month_evs = [
                    e for e in _cal_events
                    if e["start_dt"]
                    and e["start_dt"].year  == st.session_state.cal_year
                    and e["start_dt"].month == st.session_state.cal_month
                ]
                if _month_evs:
                    st.caption(f"{len(_month_evs)} event(s) this month — select one to view details:")
                    _ev_labels_month = [
                        f"{e['start_dt'].strftime('%b %d  %H:%M')}  —  {e['subject']}"
                        for e in _month_evs
                    ]
                    _picked_month = st.selectbox(
                        "Event",
                        options=[""] + _ev_labels_month,
                        label_visibility="collapsed",
                        key="cal_month_pick",
                    )
                    if _picked_month:
                        _idx = _ev_labels_month.index(_picked_month)
                        st.session_state["cal_selected_event"] = _month_evs[_idx]
                        st.session_state.pop("cal_list_pick", None)
                else:
                    st.info("No events in this month.")

            # ── WEEK VIEW ─────────────────────────────────────────────────────
            with _view_tab_week:
                if "cal_week_offset" not in st.session_state:
                    st.session_state.cal_week_offset = 0

                _wk_base = _today - _timedelta(days=_today.weekday())
                _wk_start_w = _wk_base + _timedelta(weeks=st.session_state.cal_week_offset)
                _wk_end_w   = _wk_start_w + _timedelta(days=6)

                _wn_l, _wn_title, _wn_today, _wn_r = st.columns([1, 4, 1.2, 1])
                with _wn_l:
                    if st.button("◀", key="cal_prev_week"):
                        st.session_state.cal_week_offset -= 1
                        st.rerun()
                with _wn_title:
                    st.subheader(
                        f"{_wk_start_w.strftime('%d %b')} – {_wk_end_w.strftime('%d %b %Y')}"
                    )
                with _wn_today:
                    if st.button("Today", key="cal_goto_today_week"):
                        st.session_state.cal_week_offset = 0
                        st.rerun()
                with _wn_r:
                    if st.button("▶", key="cal_next_week"):
                        st.session_state.cal_week_offset += 1
                        st.rerun()

                _week_days = [_wk_start_w + _timedelta(days=i) for i in range(7)]
                _week_cols = st.columns(7)

                # Day headers
                for _wi, _wd in enumerate(_week_days):
                    _is_today_w = _wd == _today
                    _hdr = f"**{'📍 ' if _is_today_w else ''}{_wd.strftime('%a')}**"
                    _num = f"**{_wd.day}**" if _is_today_w else str(_wd.day)
                    _week_cols[_wi].markdown(f"{_hdr}  \n{_num}")

                st.divider()

                # Event cells
                _week_cols2 = st.columns(7)
                for _wi, _wd in enumerate(_week_days):
                    _day_evs_w = [
                        e for e in _cal_events
                        if e["start_dt"] and e["start_dt"].date() == _wd
                    ]
                    with _week_cols2[_wi]:
                        if not _day_evs_w:
                            st.caption("—")
                        for _ev_w in _day_evs_w:
                            _t_w = calendar_reader.fmt_time(_ev_w["start_dt"])
                            _lbl_w = (
                                f"{_t_w}  \n{_ev_w['subject'][:22]}"
                                + ("…" if len(_ev_w["subject"]) > 22 else "")
                            )
                            if st.button(
                                _lbl_w,
                                key=f"cal_week_ev_{_wd}_{_ev_w['id']}",
                                use_container_width=True,
                                help=_ev_w["subject"],
                            ):
                                st.session_state["cal_selected_event"] = _ev_w
                                st.rerun()

            # ── LIST VIEW ─────────────────────────────────────────────────────
            with _view_tab_list:
                _range_options = {
                    "Today":        (_today,                         _today),
                    "This week":    (_week_start,                    _week_end),
                    "Next 7 days":  (_today,                         _today + _timedelta(days=6)),
                    "Next 30 days": (_today,                         _today + _timedelta(days=29)),
                    "Past 7 days":  (_today - _timedelta(days=6),    _today),
                    "Past 30 days": (_today - _timedelta(days=29),   _today),
                    "All events":   (
                        min((e["start_dt"].date() for e in _cal_events if e["start_dt"]),
                            default=_today),
                        max((e["start_dt"].date() for e in _cal_events if e["start_dt"]),
                            default=_today),
                    ),
                }
                _range_choice = st.selectbox(
                    "Show",
                    options=list(_range_options.keys()),
                    index=2,   # "Next 7 days" default
                    key="cal_list_range",
                )
                _r_start, _r_end = _range_options[_range_choice]
                _list_evs = calendar_reader.events_in_range(_cal_events, _r_start, _r_end)

                if not _list_evs:
                    st.info("No events in this range.")
                else:
                    # Group by date
                    _by_date: dict[_date, list] = {}
                    for _ev in _list_evs:
                        _d = _ev["start_dt"].date()
                        _by_date.setdefault(_d, []).append(_ev)

                    for _d in sorted(_by_date):
                        _is_today_d = _d == _today
                        _day_label  = (
                            f"**📌 Today, {_d.strftime('%A %d %B')}**"
                            if _is_today_d
                            else f"**{_d.strftime('%A %d %B')}**"
                        )
                        st.markdown(_day_label)

                        for _ev in _by_date[_d]:
                            with st.container(border=True):
                                _c_time, _c_subj, _c_btn = st.columns([1.8, 5, 1.5])
                                with _c_time:
                                    _t_start = calendar_reader.fmt_time(_ev["start_dt"])
                                    _t_end   = calendar_reader.fmt_time(_ev["end_dt"])
                                    _dur     = calendar_reader.fmt_duration(_ev)
                                    st.markdown(f"`{_t_start}–{_t_end}`")
                                    if _dur:
                                        st.caption(_dur)
                                with _c_subj:
                                    st.markdown(f"**{_ev['subject']}**")
                                    if _ev["organizer"]:
                                        st.caption(f"Organiser: {_ev['organizer']}")
                                with _c_btn:
                                    if st.button(
                                        "View details",
                                        key=f"cal_list_btn_{_ev['id']}_{_d}",
                                        use_container_width=True,
                                    ):
                                        st.session_state["cal_selected_event"] = _ev
                                        st.session_state.pop("cal_month_pick", None)
                                        st.session_state.pop("cal_list_pick", None)
                                        st.rerun()

            # ── EVENT DETAIL PANEL ────────────────────────────────────────────
            _sel_ev: Optional[dict] = st.session_state.get("cal_selected_event")
            if _sel_ev:
                st.divider()
                _dh_col, _close_col = st.columns([6, 1])
                with _dh_col:
                    st.subheader(_sel_ev["subject"])
                with _close_col:
                    if st.button("✕ Close", key="cal_close_detail"):
                        del st.session_state["cal_selected_event"]
                        st.rerun()

                with st.container(border=True):
                    # Time / duration / timezone row
                    _di1, _di2, _di3 = st.columns(3)
                    with _di1:
                        _sd = _sel_ev["start_dt"]
                        _ed = _sel_ev["end_dt"]
                        if _sd:
                            _date_str = _sd.strftime("%A, %d %B %Y")
                            _time_str = (
                                f"{_sd.strftime('%H:%M')} – {_ed.strftime('%H:%M')}"
                                if _ed else _sd.strftime("%H:%M")
                            )
                            st.markdown(f"📅 **{_date_str}**  \n🕐 {_time_str}")
                        _dur_str = calendar_reader.fmt_duration(_sel_ev)
                        if _dur_str:
                            st.caption(f"Duration: {_dur_str}")
                    with _di2:
                        if _sel_ev["organizer"]:
                            st.markdown("**Organiser**")
                            st.markdown(f"`{_sel_ev['organizer']}`")
                            if st.button(
                                "🔍 Search emails from organiser",
                                key="cal_search_organiser",
                            ):
                                _nav_to_search(_sel_ev["organizer"])
                    with _di3:
                        if _sel_ev["time_zone"]:
                            st.caption(f"🌐 {_sel_ev['time_zone']}")

                    # Attendees
                    _att_req = _sel_ev.get("required_attendees", [])
                    _att_opt = _sel_ev.get("optional_attendees", [])
                    if _att_req or _att_opt:
                        st.divider()
                        _a1, _a2 = st.columns(2)
                        with _a1:
                            if _att_req:
                                st.markdown("**👥 Required attendees**")
                                for _addr in _att_req:
                                    _ac, _ab = st.columns([3, 1])
                                    _ac.caption(_addr)
                                    if _ab.button("🔍", key=f"cal_att_req_{_addr}", help=f"Search emails with {_addr}"):
                                        _nav_to_search(_addr)
                        with _a2:
                            if _att_opt:
                                st.markdown("**👤 Optional attendees**")
                                for _addr in _att_opt:
                                    _ac, _ab = st.columns([3, 1])
                                    _ac.caption(_addr)
                                    if _ab.button("🔍", key=f"cal_att_opt_{_addr}", help=f"Search emails with {_addr}"):
                                        _nav_to_search(_addr)

                    # Body
                    if _sel_ev.get("body"):
                        st.divider()
                        with st.expander("📄 Invite body", expanded=False):
                            st.text(_sel_ev["body"][:3000])

                # ── Related emails ─────────────────────────────────────────
                st.markdown("### 📧 Related emails")
                st.caption(
                    "Priority: 👥 Attendee/organiser overlap · 📝 Subject match · "
                    "🔍 Semantic match on invite · 🏷 Entity · 🔖 Tag"
                )

                with st.spinner("Finding related emails…"):
                    _related = calendar_reader.find_related_emails(_sel_ev, limit=15)

                if not _related:
                    if indexer.get_email_count() == 0:
                        st.info("No emails indexed yet — index some emails first.")
                    else:
                        st.info("No strongly related emails found for this event.")
                else:
                    # Tag summary across related emails
                    _rel_ids   = [e["id"] for e in _related]
                    _tag_summ  = calendar_reader.tag_summary(_rel_ids)
                    if _tag_summ:
                        _tag_badges = "  ".join(
                            f"`{t['name']}` ×{t['cnt']}" for t in _tag_summ[:8]
                        )
                        st.caption(f"**Common tags across related emails:** {_tag_badges}")

                    for _rem in _related:
                        with st.container(border=True):
                            _rh_col, _rb_col = st.columns([5, 1])
                            with _rh_col:
                                st.markdown(
                                    f"**{_rem.get('subject') or '(no subject)'}**"
                                )
                                st.caption(
                                    f"From: {_rem.get('sender_name', '')} "
                                    f"<{_rem.get('sender_email', '')}>"
                                    f"  ·  {(_rem.get('date') or '')[:16]}"
                                    + ("  ·  📎" if _rem.get("has_attachments") else "")
                                )
                                # Show why this email was matched
                                _signals = _rem.get("_match_signals", [])
                                if _signals:
                                    st.caption("Matched via: " + "  ·  ".join(_signals))
                                # Show tags on this email
                                _em_tags = tagger.get_email_tags(_rem["id"])
                                if _em_tags:
                                    _tbadges = "  ".join(
                                        f"`{'👤' if t['source']=='manual' else '🤖'} {t['name']}`"
                                        for t in _em_tags
                                    )
                                    st.caption(f"Tags: {_tbadges}")
                            with _rb_col:
                                if st.button(
                                    "Open",
                                    key=f"cal_open_rel_{_sel_ev['id']}_{_rem['id']}",
                                    use_container_width=True,
                                    help="Open this email in the Search tab",
                                ):
                                    _nav_to_email(_rem["id"])


# ════════════════════════════════════════════════════════════════════════════
# SETTINGS TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_settings:
    st.subheader("Settings")
    settings = config.load_settings()

    if _mode == "offline":
        # ── Offline: local .eml folder ────────────────────────────────────────
        new_folder = st.text_input(
            "EML folder path",
            value=settings.get("email_folder", config.DEFAULT_EMAIL_FOLDER),
            key="settings_folder",
        )
        if st.button("Save folder", key="save_folder"):
            settings["email_folder"] = new_folder
            config.save_settings(settings)
            st.success("Saved. Restart the app to pick up the new folder.")

        st.divider()
        st.subheader("Indexing")
        if st.button("Index new emails now", type="primary", key="index_now"):
            with st.spinner("Scanning for new emails…"):
                folder = settings.get("email_folder", config.DEFAULT_EMAIL_FOLDER)
                new_files = indexer.get_unindexed_files(folder)
                if not new_files:
                    st.info("No new .eml files found.")
                else:
                    from modules.watcher import run_initial_index
                    result = run_initial_index(folder)
                    st.success(f"Indexed {result['indexed']} new emails ({result['total']} total).")
                    st.rerun()

        if st.button("Backfill organisation entities from email addresses", key="backfill_orgs"):
            with st.spinner("Extracting organisations from all indexed emails…"):
                all_ids = indexer.get_all_email_ids()
                count = 0
                for eid in all_ids:
                    em = indexer.get_email_by_id(eid)
                    if em:
                        orgs = nlp_engine.extract_orgs_from_email_addrs(em)
                        if orgs:
                            indexer.insert_entities(eid, orgs)
                            count += 1
            st.success(f"Done — extracted organisation entities for {count} email(s). Rebuild the graph to see them.")

    if _mode == "online":
        # ── IMAP connection ───────────────────────────────────────────────────
        st.divider()
        st.subheader("IMAP connection")
        st.caption(
            "Connect directly to a mail server. "
            "Credentials are saved to `data/settings.json` on this machine only."
        )

        _imap_cfg = settings.get("imap", {})

        # Auth method selector
        _auth_options = ["Password (Gmail, iCloud, custom server)", "Microsoft OAuth2 (Outlook.com / Hotmail / Live)"]
        _auth_idx = 1 if _imap_cfg.get("use_oauth2") else 0
        _auth_method = st.radio(
            "Authentication method",
            _auth_options,
            index=_auth_idx,
            horizontal=True,
            key="imap_auth_method",
        )
        _use_oauth2 = _auth_method.startswith("Microsoft")

        # Common fields
        _imap_user = st.text_input(
            "Email address",
            value=_imap_cfg.get("username", ""),
            placeholder="you@outlook.com",
            key="imap_user",
        )
        _imap_mailbox = st.text_input(
            "Mailbox",
            value=_imap_cfg.get("mailbox", "INBOX"),
            help='e.g. INBOX, Sent Items, "[Gmail]/All Mail"',
            key="imap_mailbox",
        )

        if not _use_oauth2:
            # ── Password / basic auth ─────────────────────────────────────
            _col1, _col2 = st.columns(2)
            with _col1:
                _imap_host = st.text_input(
                    "IMAP host",
                    value=_imap_cfg.get("host", ""),
                    placeholder="imap.gmail.com",
                    key="imap_host",
                )
                _imap_pass = st.text_input(
                    "Password / app password",
                    value=_imap_cfg.get("password", ""),
                    type="password",
                    placeholder="Leave blank to keep saved password",
                    key="imap_pass",
                )
            with _col2:
                _imap_port = st.number_input(
                    "Port", value=int(_imap_cfg.get("port", 993)),
                    min_value=1, max_value=65535, step=1, key="imap_port",
                )
                _imap_ssl = st.checkbox(
                    "Use SSL", value=bool(_imap_cfg.get("use_ssl", True)), key="imap_ssl"
                )

            _save_col, _test_col = st.columns(2)
            with _save_col:
                if st.button("Save IMAP settings", key="imap_save"):
                    _new_pass = _imap_pass.strip() or _imap_cfg.get("password", "")
                    settings["imap"] = {
                        "use_oauth2": False,
                        "host": _imap_host.strip(),
                        "username": _imap_user.strip(),
                        "password": _new_pass,
                        "port": int(_imap_port),
                        "use_ssl": _imap_ssl,
                        "mailbox": _imap_mailbox.strip(),
                    }
                    config.save_settings(settings)
                    st.success("IMAP settings saved.")
            with _test_col:
                if st.button("Test connection", key="imap_test"):
                    _h = _imap_host.strip()
                    _u = _imap_user.strip()
                    _p = _imap_pass.strip() or _imap_cfg.get("password", "")
                    if not (_h and _u and _p):
                        st.warning("Fill in host, email address, and password first.")
                    else:
                        with st.spinner("Connecting…"):
                            try:
                                _tc = IMAPConnector(_h, _u, _p, int(_imap_port), _imap_ssl)
                                _ic = _tc._connect()
                                _ic.noop()
                                _ic.logout()
                                st.success(f"Connected to {_h}.")
                            except Exception as _exc:
                                st.error(f"Connection failed: {_exc}")

        else:
            # ── Microsoft OAuth2 ──────────────────────────────────────────
            try:
                import msal as _msal
                _msal_available = True
            except ImportError:
                _msal_available = False
                st.error("msal not installed. Run: `pip install msal` then restart the app.")

            if _msal_available:
                _client_id_input = st.text_input(
                    "Azure Application (Client) ID",
                    value=_imap_cfg.get("client_id", ""),
                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                    help="From Azure Portal → App registrations → your app → Overview",
                    key="imap_client_id",
                )

                _has_token = bool(_imap_cfg.get("access_token"))

                if _has_token:
                    # Already authenticated — show status and management buttons
                    st.success(
                        f"✓ Signed in as **{_imap_cfg.get('username', '')}** "
                        f"— OAuth2 tokens stored"
                    )
                    _sc, _rc, _tc = st.columns(3)
                    with _sc:
                        if st.button("Save mailbox", key="imap_save_oauth"):
                            settings.setdefault("imap", {})
                            settings["imap"]["mailbox"] = _imap_mailbox.strip()
                            settings["imap"]["username"] = _imap_user.strip()
                            settings["imap"]["client_id"] = _client_id_input.strip()
                            config.save_settings(settings)
                            st.success("Saved.")
                    with _rc:
                        if st.button("Re-authenticate", key="imap_reauth"):
                            settings.setdefault("imap", {})
                            settings["imap"].pop("access_token", None)
                            settings["imap"].pop("refresh_token", None)
                            config.save_settings(settings)
                            for _k in ["_oauth_flow", "_oauth_client_id", "_oauth_imap_user"]:
                                st.session_state.pop(_k, None)
                            st.rerun()
                    with _tc:
                        if st.button("Test connection", key="imap_test_oauth"):
                            with st.spinner("Connecting…"):
                                try:
                                    _tc_conn = IMAPConnector(
                                        host=_imap_cfg.get("host", "imap-mail.outlook.com"),
                                        username=_imap_cfg.get("username", ""),
                                        access_token=_imap_cfg.get("access_token", ""),
                                        refresh_token=_imap_cfg.get("refresh_token", ""),
                                        client_id=_imap_cfg.get("client_id", ""),
                                    )
                                    _ic = _tc_conn._connect()
                                    _ic.noop()
                                    _ic.logout()
                                    if _tc_conn.new_tokens:
                                        settings["imap"]["access_token"] = _tc_conn.new_tokens["access_token"]
                                        settings["imap"]["refresh_token"] = _tc_conn.new_tokens["refresh_token"]
                                        config.save_settings(settings)
                                    st.success("Connected successfully.")
                                except Exception as _exc:
                                    st.error(f"Connection failed: {_exc}")
                else:
                    # Not yet authenticated — device code flow
                    if "_oauth_flow" not in st.session_state:
                        st.caption(
                            "Click **Start Microsoft sign-in** — you'll get a short code to "
                            "enter at microsoft.com. No redirect URL or server setup needed."
                        )
                        if st.button(
                            "Start Microsoft sign-in",
                            type="primary",
                            key="imap_start_oauth",
                            disabled=not (_client_id_input.strip() and _imap_user.strip()),
                            help="Fill in your Client ID and email address first.",
                        ):
                            with st.spinner("Contacting Microsoft…"):
                                _msal_app = _msal.PublicClientApplication(
                                    _client_id_input.strip(),
                                    authority=MICROSOFT_AUTHORITY,
                                )
                                _flow = _msal_app.initiate_device_flow(scopes=OUTLOOK_IMAP_SCOPE)
                            if "user_code" in _flow:
                                st.session_state["_oauth_flow"] = _flow
                                st.session_state["_oauth_client_id"] = _client_id_input.strip()
                                st.session_state["_oauth_imap_user"] = _imap_user.strip()
                                # Pre-save non-secret fields so they survive the rerun
                                settings.setdefault("imap", {})
                                settings["imap"].update({
                                    "use_oauth2": True,
                                    "client_id": _client_id_input.strip(),
                                    "username": _imap_user.strip(),
                                    "host": "imap-mail.outlook.com",
                                    "port": 993,
                                    "use_ssl": True,
                                    "mailbox": _imap_mailbox.strip(),
                                })
                                config.save_settings(settings)
                                st.rerun()
                            else:
                                st.error(
                                    f"Could not start sign-in: "
                                    f"{_flow.get('error_description', str(_flow))}"
                                )
                    else:
                        # Show the device code to the user
                        _flow = st.session_state["_oauth_flow"]
                        st.info(
                            f"**Step 1** — Open this URL in your browser:  \n"
                            f"### {_flow['verification_uri']}\n\n"
                            f"**Step 2** — Enter this code when prompted:  \n"
                            f"# `{_flow['user_code']}`\n\n"
                            f"**Step 3** — Sign in with your Microsoft / Outlook account\n\n"
                            f"**Step 4** — Come back here and click **I've signed in**"
                        )
                        _done_col, _cancel_col = st.columns([2, 1])
                        with _done_col:
                            if st.button("I've signed in ✓", type="primary", key="imap_complete_oauth"):
                                with st.spinner("Checking with Microsoft… (up to 30 s)"):
                                    import threading as _th
                                    _cid = st.session_state.get("_oauth_client_id", "")
                                    _msal_app = _msal.PublicClientApplication(
                                        _cid, authority=MICROSOFT_AUTHORITY
                                    )
                                    _tok_holder: dict = {}
                                    _done_evt = _th.Event()

                                    def _acquire():
                                        _tok_holder["r"] = _msal_app.acquire_token_by_device_flow(_flow)
                                        _done_evt.set()

                                    _th.Thread(target=_acquire, daemon=True).start()
                                    _done_evt.wait(timeout=30)

                                if "access_token" in _tok_holder.get("r", {}):
                                    _tok = _tok_holder["r"]
                                    settings.setdefault("imap", {})
                                    settings["imap"].update({
                                        "access_token": _tok["access_token"],
                                        "refresh_token": _tok.get("refresh_token", ""),
                                    })
                                    config.save_settings(settings)
                                    for _k in ["_oauth_flow", "_oauth_client_id", "_oauth_imap_user"]:
                                        st.session_state.pop(_k, None)
                                    st.success("Signed in! You can now fetch your emails below.")
                                    st.rerun()
                                elif "r" in _tok_holder:
                                    st.error(
                                        f"Sign-in failed: "
                                        f"{_tok_holder['r'].get('error_description', str(_tok_holder['r']))}"
                                    )
                                else:
                                    st.warning(
                                        "Timed out — make sure you completed sign-in in your browser, "
                                        "then click **I've signed in** again."
                                    )
                        with _cancel_col:
                            if st.button("Cancel", key="imap_cancel_oauth"):
                                for _k in ["_oauth_flow", "_oauth_client_id", "_oauth_imap_user"]:
                                    st.session_state.pop(_k, None)
                                st.rerun()

        # ── Fetch emails via IMAP ─────────────────────────────────────────────────
        st.divider()
        st.subheader("Fetch emails via IMAP")

        _saved_imap = settings.get("imap", {})
        _oauth2_ready = _saved_imap.get("use_oauth2") and bool(_saved_imap.get("access_token"))
        _basic_ready = (
            not _saved_imap.get("use_oauth2")
            and all(_saved_imap.get(k) for k in ("host", "username", "password"))
        )
        _imap_ready = _oauth2_ready or _basic_ready

        if not _imap_ready:
            st.info("Set up and authenticate your IMAP connection above first.")
        else:
            _fetch_col, _batch_col = st.columns([3, 1])
            with _fetch_col:
                _fetch_mailbox = st.text_input(
                    "Mailbox to fetch",
                    value=_saved_imap.get("mailbox", "INBOX"),
                    key="imap_fetch_mailbox",
                )
            with _batch_col:
                _fetch_limit = st.number_input(
                    "Max emails",
                    value=20000,
                    min_value=1,
                    step=1000,
                    help="Maximum new emails per run. Already-indexed emails are always skipped.",
                    key="imap_fetch_limit",
                )

            if st.button("Fetch & index emails", type="primary", key="imap_fetch"):
                if _saved_imap.get("use_oauth2"):
                    _fc = IMAPConnector(
                        host=_saved_imap.get("host", "imap-mail.outlook.com"),
                        username=_saved_imap.get("username", ""),
                        access_token=_saved_imap.get("access_token", ""),
                        refresh_token=_saved_imap.get("refresh_token", ""),
                        client_id=_saved_imap.get("client_id", ""),
                    )
                else:
                    _fc = IMAPConnector(
                        host=_saved_imap["host"],
                        username=_saved_imap["username"],
                        password=_saved_imap["password"],
                        port=int(_saved_imap.get("port", 993)),
                        use_ssl=bool(_saved_imap.get("use_ssl", True)),
                    )

                _progress = st.progress(0, text="Connecting to IMAP server…")
                _status_box = st.empty()

                _result_holder: dict = {}
                _done_flag = threading.Event()

                def _run_fetch():
                    _result_holder["result"] = _fc.fetch_and_index(
                        mailbox=_fetch_mailbox.strip(),
                        batch_size=int(_fetch_limit),
                    )
                    _done_flag.set()

                threading.Thread(target=_run_fetch, daemon=True).start()
                while not _done_flag.wait(timeout=0.5):
                    _status_box.caption(f"⏳ {_fc.status}")

                _result = _result_holder.get("result")
                _progress.progress(1.0, text="Done!")
                _status_box.empty()

                if _result:
                    # Persist any auto-refreshed OAuth2 tokens
                    if _result.get("new_tokens"):
                        settings["imap"]["access_token"] = _result["new_tokens"]["access_token"]
                        settings["imap"]["refresh_token"] = _result["new_tokens"]["refresh_token"]
                        config.save_settings(settings)

                    _del_msg = f" Deleted: {_result['deleted']:,}." if _result.get("deleted") else ""
                    st.success(
                        f"Fetched **{_result['indexed']:,}** new email(s).{_del_msg} "
                        f"Skipped: {_result['skipped']:,}. "
                        f"Errors: {_result['errors']:,}. "
                        f"Total in DB: **{_result['total_in_db']:,}**."
                    )
                    st.rerun()
                else:
                    st.error("Fetch failed — check the connection settings above.")

        # ── Background polling ────────────────────────────────────────────────────
        st.divider()
        st.subheader("Background polling")
        st.caption(
            "Automatically check for new emails in the background while the app is running. "
            "Status is shown in the sidebar."
        )

        _saved_imap = settings.get("imap", {})
        _poll_col, _sync_col = st.columns(2)
        with _poll_col:
            _poll_minutes = st.number_input(
                "Check for new emails every (minutes)",
                min_value=1,
                max_value=1440,
                value=int(_saved_imap.get("poll_interval", 300)) // 60,
                step=1,
                key="imap_poll_interval",
            )
        with _sync_col:
            _sync_del = st.checkbox(
                "Sync deletions",
                value=bool(_saved_imap.get("sync_deletions", True)),
                key="imap_sync_deletions",
                help="Remove emails from the local index when they are deleted on the server.",
            )

        if st.button("Save polling settings", key="save_poll_interval"):
            settings.setdefault("imap", {})
            settings["imap"]["poll_interval"] = int(_poll_minutes) * 60
            settings["imap"]["sync_deletions"] = _sync_del
            config.save_settings(settings)
            st.success(
                f"Saved — polling every {_poll_minutes} min, "
                f"deletion sync {'on' if _sync_del else 'off'}. "
                "Restart the app for changes to take effect."
            )

        if _imap_poller:
            _del_info = ", deletions synced" if _saved_imap.get("sync_deletions", True) else ""
            st.info(
                f"Poller running — checking **{_saved_imap.get('mailbox', 'INBOX')}** "
                f"every **{int(_saved_imap.get('poll_interval', 300)) // 60} min**{_del_info}. "
                f"Last run: **{_imap_poller.last_indexed:,}** indexed"
                + (f", **{_imap_poller.last_deleted:,}** deleted" if _imap_poller.last_deleted else "") + "."
            )
        else:
            st.warning("Poller not running — configure and authenticate IMAP above first.")

    # ── Calendar (always shown) ───────────────────────────────────────────────
    st.divider()
    st.subheader("Calendar")
    st.caption(
        "Path to the JSON file produced by your calendar automation. "
        "The file is re-read automatically whenever it changes (no restart needed)."
    )
    _cal_path_current = settings.get("calendar_json_path", "")
    _cal_path_input = st.text_input(
        "Events JSON file path",
        value=_cal_path_current,
        placeholder="/Users/you/calendar_events.json",
        key="cal_json_path_input",
    )
    if st.button("Save calendar path", key="cal_save_path"):
        settings["calendar_json_path"] = _cal_path_input.strip()
        config.save_settings(settings)
        st.success("Calendar path saved — switch to the **Calendar** tab to view events.")
    if _cal_path_current:
        from pathlib import Path as _P
        _exists = _P(_cal_path_current).exists()
        if _exists:
            _events_count = len(calendar_reader.load_events(_cal_path_current))
            st.caption(f"✓ File found — {_events_count} event(s) loaded.")
        else:
            st.warning(f"File not found: `{_cal_path_current}`")

    st.markdown("**Display timezone**")
    _tz_options = [
        "Asia/Singapore",
        "UTC",
        "Asia/Jakarta",
        "Asia/Bangkok",
        "Asia/Kuala_Lumpur",
        "Asia/Tokyo",
        "Asia/Shanghai",
        "Asia/Hong_Kong",
        "Asia/Seoul",
        "Europe/London",
        "Europe/Berlin",
        "Europe/Paris",
        "America/New_York",
        "America/Chicago",
        "America/Los_Angeles",
        "Australia/Sydney",
    ]
    _current_tz = settings.get("calendar_display_tz", "Asia/Singapore")
    _tz_idx = _tz_options.index(_current_tz) if _current_tz in _tz_options else 0
    _tz_input = st.selectbox(
        "Timezone",
        options=_tz_options,
        index=_tz_idx,
        key="cal_tz_input",
        label_visibility="collapsed",
    )
    st.caption("Event times from the JSON are assumed to be UTC and converted to this timezone for display.")
    if st.button("Save timezone", key="cal_save_tz"):
        settings["calendar_display_tz"] = _tz_input
        config.save_settings(settings)
        st.success(f"Timezone saved — events will display in {_tz_input}.")

    # ── Database (always shown) ───────────────────────────────────────────────
    st.divider()
    st.subheader("Database")
    db_size    = Path(config.DB_PATH).stat().st_size / 1024 if Path(config.DB_PATH).exists() else 0
    graph_size = Path(config.GRAPH_DATA_PATH).stat().st_size / 1024 if Path(config.GRAPH_DATA_PATH).exists() else 0
    st.write(f"Index DB: **{db_size:.1f} KB** | Graph ABox: **{graph_size:.1f} KB**")

    ont_path = Path(config.ONTOLOGY_PATH)
    if ont_path.exists():
        with st.expander("View ontology (TBox — read-only here)"):
            st.code(ont_path.read_text(), language="turtle")

    st.divider()
    st.caption(
        "Ontology: `ontology/email_ontology.ttl` — edit in any text editor, then rebuild the graph.\n"
        "Data: `data/email_data.ttl` — auto-generated, do not edit manually."
    )
