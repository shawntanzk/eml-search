"""EML Search — Streamlit app entry point."""
import sys
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR))

import config
from modules import indexer, nlp_engine, search_engine, semantic_search, tagger
from modules.watcher import EmailWatcher
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


_watcher = _get_watcher(_current_folder())

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
    total = indexer.get_email_count()
    st.metric("Emails indexed", total)
    st.caption(f"Watcher: {_watcher.status}")

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
tab_search, tab_tags, tab_graph, tab_settings = st.tabs(
    ["Search", "Tags", "Knowledge Graph", "Settings"]
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
# SETTINGS TAB
# ════════════════════════════════════════════════════════════════════════════
with tab_settings:
    st.subheader("Settings")
    settings = config.load_settings()

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
