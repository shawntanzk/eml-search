"""RDF/OWL DL-compliant knowledge graph builder.

TBox (ontology schema) lives in ontology/email_ontology.ttl — editable by the user.
ABox (email instance data) is generated here and saved to data/email_data.ttl.
The two are always kept separate so the ontology can be modified without
touching the data graph.
"""
import hashlib
import json
import re
from pathlib import Path
from typing import Optional

from rdflib import Graph, Namespace, URIRef, Literal, RDF, RDFS, OWL, XSD

import config

DATA = Namespace("http://emailsearch.local/data#")
ONT  = Namespace("http://emailsearch.local/ontology#")

# Ordered list used for type detection (first match wins)
_NODE_TYPE_MAP: list[tuple[str, str, str]] = [
    (str(ONT.Email),        "Email",        "#1a73e8"),
    (str(ONT.Person),       "Person",       "#e67e22"),
    (str(ONT.Organization), "Organization", "#27ae60"),
    (str(ONT.Topic),        "Tag",          "#8e44ad"),
    (str(ONT.Thread),       "Thread",       "#c0392b"),
    (str(ONT.Location),     "Location",     "#00897b"),
]

INTERESTING_PROPS = frozenset({
    ONT.hasSender, ONT.hasRecipient, ONT.belongsToThread,
    ONT.hasTopic, ONT.mentionsPerson, ONT.mentionsOrganization,
    ONT.mentionsLocation,
})


def _safe_uri(text: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_\-]", "_", text.strip())
    return text[:80] or "unknown"


def _get_label(g: Graph, uri: URIRef) -> str:
    """Best human-readable label for a node."""
    for _, p, o in g.triples((uri, None, None)):
        if p in (ONT.personName, ONT.emailAddress, ONT.organizationName,
                 ONT.hasSubject, RDFS.label):
            return str(o)[:60]
    return str(uri).split("#")[-1][:60]


def _get_node_type_and_color(g: Graph, uri: URIRef) -> tuple[str, str]:
    types = {str(o) for _, _, o in g.triples((uri, RDF.type, None))}
    for type_uri, tname, tcolor in _NODE_TYPE_MAP:
        if type_uri in types:
            return tname, tcolor
    return "Other", "#95a5a6"


def _email_uri(email_id: str) -> URIRef:
    return DATA[f"email_{email_id}"]

def _person_uri(email_addr: str) -> URIRef:
    key = hashlib.md5(email_addr.lower().encode()).hexdigest()[:12]
    return DATA[f"person_{key}"]

def _org_uri(name: str) -> URIRef:
    return DATA[f"org_{_safe_uri(name.lower())}"]

def _tag_uri(tag_name: str) -> URIRef:
    return DATA[f"tag_{_safe_uri(tag_name.lower())}"]

def _thread_uri(thread_id: str) -> URIRef:
    key = hashlib.md5(thread_id.encode()).hexdigest()[:12]
    return DATA[f"thread_{key}"]

def _loc_uri(name: str) -> URIRef:
    return DATA[f"loc_{_safe_uri(name.lower())}"]


def load_tbox() -> Graph:
    """Load the user-editable TBox ontology."""
    g = Graph()
    if Path(config.ONTOLOGY_PATH).exists():
        g.parse(config.ONTOLOGY_PATH, format="turtle")
    return g


def build_abox(
    emails: list[dict],
    entities_map: dict[str, list[dict]],
    tags_map: dict[str, list[str]],
) -> Graph:
    """
    Build the ABox (instance data) graph.

    Parameters
    ----------
    emails       : list of email dicts from indexer.get_email_by_id
    entities_map : {email_id: [{text, label}, ...]}
    tags_map     : {email_id: [tag_name, ...]}
    """
    g = Graph()
    g.bind("data", DATA)
    g.bind("ont", ONT)
    g.bind("owl", OWL)
    g.bind("xsd", XSD)

    # Declare tag individuals (one per unique tag name)
    all_tag_names: set[str] = set()
    for names in tags_map.values():
        all_tag_names.update(names)
    for tag_name in all_tag_names:
        t_uri = _tag_uri(tag_name)
        g.add((t_uri, RDF.type, ONT.Topic))
        g.add((t_uri, RDF.type, OWL.NamedIndividual))
        g.add((t_uri, RDFS.label, Literal(tag_name)))

    known_persons: dict[str, URIRef] = {}

    def _ensure_person(addr: str, name: str) -> URIRef:
        addr = addr.lower()
        if addr not in known_persons:
            uri = _person_uri(addr)
            g.add((uri, RDF.type, ONT.Person))
            g.add((uri, RDF.type, OWL.NamedIndividual))
            if name:
                g.add((uri, ONT.personName, Literal(name)))
            if addr:
                g.add((uri, ONT.emailAddress, Literal(addr)))
            known_persons[addr] = uri
        return known_persons[addr]

    for em in emails:
        eid = em["id"]
        e_uri = _email_uri(eid)

        g.add((e_uri, RDF.type, ONT.Email))
        g.add((e_uri, RDF.type, OWL.NamedIndividual))

        if em.get("subject"):
            g.add((e_uri, ONT.hasSubject, Literal(em["subject"])))
        if em.get("message_id"):
            g.add((e_uri, ONT.hasMessageId, Literal(em["message_id"])))
        if em.get("date"):
            try:
                g.add((e_uri, ONT.hasSentDate, Literal(em["date"], datatype=XSD.dateTime)))
            except Exception:
                pass

        if em.get("sender_email"):
            sender_uri = _ensure_person(em["sender_email"], em.get("sender_name", ""))
            g.add((e_uri, ONT.hasSender, sender_uri))

        recipients = em.get("recipients", [])
        if isinstance(recipients, str):
            try:
                recipients = json.loads(recipients)
            except Exception:
                recipients = []
        for r in recipients:
            if r.get("email"):
                r_uri = _ensure_person(r["email"], r.get("name", ""))
                g.add((e_uri, ONT.hasRecipient, r_uri))

        if em.get("thread_id"):
            th_uri = _thread_uri(em["thread_id"])
            g.add((th_uri, RDF.type, ONT.Thread))
            g.add((th_uri, RDF.type, OWL.NamedIndividual))
            g.add((e_uri, ONT.belongsToThread, th_uri))

        for tag_name in tags_map.get(eid, []):
            g.add((e_uri, ONT.hasTopic, _tag_uri(tag_name)))

        for ent in entities_map.get(eid, []):
            label = ent["label"]
            text = ent["text"]
            if label == "PERSON":
                ent_uri = DATA[f"person_name_{_safe_uri(text.lower())}"]
                g.add((ent_uri, RDF.type, ONT.Person))
                g.add((ent_uri, RDF.type, OWL.NamedIndividual))
                g.add((ent_uri, ONT.personName, Literal(text)))
                g.add((e_uri, ONT.mentionsPerson, ent_uri))
            elif label == "ORG":
                org_uri = _org_uri(text)
                g.add((org_uri, RDF.type, ONT.Organization))
                g.add((org_uri, RDF.type, OWL.NamedIndividual))
                g.add((org_uri, ONT.organizationName, Literal(text)))
                g.add((e_uri, ONT.mentionsOrganization, org_uri))
            elif label in ("GPE", "LOC"):
                loc_uri = _loc_uri(text)
                g.add((loc_uri, RDF.type, ONT.Location))
                g.add((loc_uri, RDF.type, OWL.NamedIndividual))
                g.add((loc_uri, RDFS.label, Literal(text)))
                g.add((e_uri, ONT.mentionsLocation, loc_uri))

    return g


def save_abox(g: Graph) -> None:
    g.serialize(destination=config.GRAPH_DATA_PATH, format="turtle")


def load_abox() -> Graph:
    g = Graph()
    if Path(config.GRAPH_DATA_PATH).exists():
        g.parse(config.GRAPH_DATA_PATH, format="turtle")
    return g


def get_merged_graph() -> Graph:
    tbox = load_tbox()
    abox = load_abox()
    merged = Graph()
    for triple in tbox:
        merged.add(triple)
    for triple in abox:
        merged.add(triple)
    return merged


def sparql_query(graph: Graph, query_string: str) -> list[dict]:
    try:
        results = graph.query(query_string)
        return [{str(k): str(v) for k, v in zip(results.vars, row)} for row in results]
    except Exception as exc:
        return [{"error": str(exc)}]


def get_graph_stats(g: Optional[Graph] = None) -> dict:
    if g is None:
        g = load_abox()
    return {
        "triples": len(g),
        "emails": len(set(g.subjects(RDF.type, ONT.Email))),
        "persons": len(set(g.subjects(RDF.type, ONT.Person))),
        "organizations": len(set(g.subjects(RDF.type, ONT.Organization))),
        "topics": len(set(g.subjects(RDF.type, ONT.Topic))),
        "threads": len(set(g.subjects(RDF.type, ONT.Thread))),
    }


def get_all_graph_nodes(g: Optional[Graph] = None) -> list[dict]:
    """Return every named individual with its type, label, and color — for the selection UI."""
    if g is None:
        g = load_abox()
    result = []
    seen: set[str] = set()
    for uri in g.subjects(RDF.type, OWL.NamedIndividual):
        uid = str(uri)
        if uid in seen:
            continue
        seen.add(uid)
        tname, tcolor = _get_node_type_and_color(g, uri)
        result.append({
            "uri":   uid,
            "label": _get_label(g, uri),
            "type":  tname,
            "color": tcolor,
        })
    return sorted(result, key=lambda x: (x["type"], x["label"].lower()))


def get_subgraph(
    g: Graph,
    seed_uris: list[str],
    allowed_types: set[str],
) -> tuple[list[dict], list[dict]]:
    """
    Return (nodes, edges) for the given seed nodes plus their 1-hop neighbours,
    restricted to node types in allowed_types.
    Seeds are always included regardless of type filter.
    """
    seed_set = set(seed_uris)
    include_uris: set[str] = set(seed_set)

    for seed_str in seed_uris:
        seed_ref = URIRef(seed_str)
        for _, p, o in g.triples((seed_ref, None, None)):
            if p in INTERESTING_PROPS and isinstance(o, URIRef):
                tname, _ = _get_node_type_and_color(g, o)
                if tname in allowed_types:
                    include_uris.add(str(o))
        for s, p, _ in g.triples((None, None, seed_ref)):
            if p in INTERESTING_PROPS and isinstance(s, URIRef):
                tname, _ = _get_node_type_and_color(g, s)
                if tname in allowed_types:
                    include_uris.add(str(s))

    nodes: list[dict] = []
    for uid in include_uris:
        uri = URIRef(uid)
        tname, tcolor = _get_node_type_and_color(g, uri)
        nodes.append({
            "id":    uid,
            "label": _get_label(g, uri),
            "color": tcolor,
            "type":  tname,
        })

    edges: list[dict] = []
    for s, p, o in g:
        if p not in INTERESTING_PROPS or not isinstance(o, URIRef):
            continue
        sid, oid = str(s), str(o)
        if sid in include_uris and oid in include_uris:
            edges.append({
                "from":  sid,
                "to":    oid,
                "label": str(p).split("#")[-1],
            })

    return nodes, edges
