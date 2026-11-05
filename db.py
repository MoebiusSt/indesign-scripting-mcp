"""
InDesign DOM Database Query Layer.

Provides query functions for MCP tools with multi-source support.
"""

import re
import sqlite3
from pathlib import Path

DEFAULT_DB = Path(__file__).parent / "extendscript.db"

RUNTIMES = ("extendscript", "uxp")
CLASS_TYPE_FILTERS = {
    "all": "",
    "class": " AND c.is_enum = 0",
    "enum": " AND c.kind = 'enum'",
    "object": " AND c.kind = 'object'",
    "collection": " AND c.kind = 'collection'",
}


def _connect(db_path: str | None = None) -> sqlite3.Connection:
    """Open a read-only connection."""
    path = db_path or str(DEFAULT_DB)
    conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _source_clause(source: str | None) -> tuple[str, tuple]:
    """Build SQL source filter clause."""
    if source:
        return " AND src.key = ?", (source,)
    return "", ()


def _with_runtime(item: dict, row: sqlite3.Row, runtime: str | None = None, uxp_fallback: str | None = None) -> dict:
    """Attach runtime-specific fields of a row.

    uxp_note / uxp_type are dropped for runtime='extendscript', extendscript_note for runtime='uxp';
    the general note is always kept.
    """
    keys = row.keys()

    def value(key: str) -> str | None:
        return row[key] if key in keys else None

    if runtime != "extendscript":
        uxp_note = value("uxp_note") or uxp_fallback
        if uxp_note:
            item["uxp_note"] = uxp_note
        if value("uxp_type"):
            item["uxp_type"] = value("uxp_type")
    if runtime != "uxp" and value("es_note"):
        item["extendscript_note"] = value("es_note")
    if value("note"):
        item["note"] = value("note")
    return item


def _class_rows(conn: sqlite3.Connection, name: str, source: str | None = None):
    """Fetch class rows by class name and optional source."""
    clause, params = _source_clause(source)
    return conn.execute(
        f"""SELECT c.id, c.name, c.kind, c.is_enum, c.description, c.superclass_name, c.collection_of,
                  c.doc_url, c.uxp_note, c.es_note, c.note, src.key AS source, src.uxp_note AS source_uxp_note
           FROM classes c
           JOIN sources src ON c.source_id = src.id
           WHERE c.name = ?{clause}
           ORDER BY src.key""",
        (name, *params),
    ).fetchall()


def lookup_class(name: str, source: str | None = None, db_path: str | None = None) -> dict | list[dict] | None:
    """Full class info for one class name, optionally source-filtered."""
    conn = _connect(db_path)
    rows = _class_rows(conn, name=name, source=source)
    if not rows:
        conn.close()
        return None

    result = []
    for row in rows:
        class_id = row["id"]
        (prop_count,) = conn.execute("SELECT COUNT(*) FROM properties WHERE class_id = ?", (class_id,)).fetchone()
        (method_count,) = conn.execute("SELECT COUNT(*) FROM methods WHERE class_id = ?", (class_id,)).fetchone()
        subclasses = [
            r["name"]
            for r in conn.execute(
                """SELECT c2.name
                   FROM classes c2
                   WHERE c2.superclass_name = ?
                     AND c2.source_id = (SELECT source_id FROM classes WHERE id = ?)
                   ORDER BY c2.name""",
                (row["name"], class_id),
            ).fetchall()
        ]
        events = [
            {"name": e["name"], "description": e["description"] or ""}
            for e in conn.execute(
                "SELECT name, description FROM events WHERE class_id = ? ORDER BY name", (class_id,)
            ).fetchall()
        ]
        item = {
            "name": row["name"],
            "source": row["source"],
            "kind": row["kind"],
            "is_enum": bool(row["is_enum"]),
            "description": row["description"] or "",
            "superclass": row["superclass_name"],
            "property_count": prop_count,
            "method_count": method_count,
            "direct_subclasses": subclasses,
            "doc_url": row["doc_url"],
        }
        if row["collection_of"]:
            item["collection_of"] = row["collection_of"]
        if events:
            item["events"] = events
        _with_runtime(item, row, uxp_fallback=row["source_uxp_note"])
        result.append(item)
    conn.close()
    return result[0] if len(result) == 1 else result


def get_properties(
    class_name: str,
    source: str | None = None,
    filter_text: str | None = None,
    include_inherited: bool = False,
    runtime: str | None = None,
    db_path: str | None = None,
) -> list[dict]:
    """Properties of one class, optionally source-filtered and inherited."""
    conn = _connect(db_path)

    class_rows = _class_rows(conn, class_name, source)
    class_ids = [row["id"] for row in class_rows]
    if include_inherited:
        inherited_ids = []
        for row in class_rows:
            inherited_ids.extend(_get_ancestor_chain_ids(conn, row["id"]))
        class_ids = list(dict.fromkeys(inherited_ids))
    if not class_ids:
        conn.close()
        return []

    placeholders = ",".join("?" for _ in class_ids)
    query = f"""
        SELECT p.name, p.description, p.data_type, p.is_array, p.is_readonly,
               p.data_type_ref, p.element_type, p.default_value, p.min_value, p.max_value,
               p.measurement_unit, p.es_only, p.uxp_type, p.uxp_note, p.es_note, p.note,
               c.name AS class_name, src.key AS source
        FROM properties p
        JOIN classes c ON p.class_id = c.id
        JOIN sources src ON c.source_id = src.id
        WHERE c.id IN ({placeholders})
    """
    params: list = list(class_ids)

    if filter_text:
        query += " AND (p.name LIKE ? OR p.description LIKE ?)"
        like = f"%{filter_text}%"
        params.extend([like, like])
    if runtime == "uxp":
        query += " AND p.es_only = 0"

    query += " ORDER BY src.key, c.name, p.element_type, p.name"

    rows = conn.execute(query, params).fetchall()
    conn.close()

    result = []
    for r in rows:
        item = {
            "name": r["name"],
            "description": r["description"] or "",
            "data_type": r["data_type"],
            "data_type_ref": r["data_type_ref"],
            "is_array": bool(r["is_array"]),
            "is_readonly": bool(r["is_readonly"]),
            "element_type": r["element_type"],
            "default_value": r["default_value"],
            "min_value": r["min_value"],
            "max_value": r["max_value"],
            "measurement_unit": bool(r["measurement_unit"]),
            "defined_in": r["class_name"],
            "source": r["source"],
        }
        if r["es_only"]:
            item["extendscript_only"] = True
        result.append(_with_runtime(item, r, runtime))
    return result


def get_methods(
    class_name: str,
    source: str | None = None,
    filter_text: str | None = None,
    include_inherited: bool = False,
    runtime: str | None = None,
    db_path: str | None = None,
) -> list[dict]:
    """Methods of one class with short signatures.

    runtime='uxp' drops methods that only exist in ExtendScript (e.g. the [] index accessor).
    """
    conn = _connect(db_path)

    class_rows = _class_rows(conn, class_name, source)
    class_ids = [row["id"] for row in class_rows]
    if include_inherited:
        inherited_ids = []
        for row in class_rows:
            inherited_ids.extend(_get_ancestor_chain_ids(conn, row["id"]))
        class_ids = list(dict.fromkeys(inherited_ids))
    if not class_ids:
        conn.close()
        return []

    placeholders = ",".join("?" for _ in class_ids)
    query = f"""
        SELECT m.id, m.name, m.description, m.return_type, m.return_is_array,
               m.return_type_ref, m.element_type, m.es_only, m.uxp_type, m.uxp_note, m.es_note, m.note,
               c.name AS class_name, src.key AS source
        FROM methods m
        JOIN classes c ON m.class_id = c.id
        JOIN sources src ON c.source_id = src.id
        WHERE c.id IN ({placeholders})
    """
    params: list = list(class_ids)

    if filter_text:
        query += " AND (m.name LIKE ? OR m.description LIKE ?)"
        like = f"%{filter_text}%"
        params.extend([like, like])
    if runtime == "uxp":
        query += " AND m.es_only = 0"

    query += " ORDER BY src.key, c.name, m.element_type, m.name"

    rows = conn.execute(query, params).fetchall()
    result = []

    for r in rows:
        param_rows = conn.execute(
            "SELECT name, data_type, is_optional FROM parameters WHERE method_id = ? ORDER BY sort_order",
            (r["id"],),
        ).fetchall()

        sig_parts = []
        for p in param_rows:
            ptype = p["data_type"] or "any"
            opt = "?" if p["is_optional"] else ""
            sig_parts.append(f"{p['name']}{opt}: {ptype}")

        ret = r["return_type"] or "void"
        signature = f"({', '.join(sig_parts)}) -> {ret}"

        item = {
            "name": r["name"],
            "description": r["description"] or "",
            "signature": signature,
            "return_type": r["return_type"],
            "return_type_ref": r["return_type_ref"],
            "return_is_array": bool(r["return_is_array"]),
            "element_type": r["element_type"],
            "defined_in": r["class_name"],
            "source": r["source"],
        }
        if r["es_only"]:
            item["extendscript_only"] = True
        result.append(_with_runtime(item, r, runtime))

    conn.close()
    return result


def get_method_detail(
    class_name: str,
    method_name: str,
    source: str | None = None,
    db_path: str | None = None,
) -> dict | list[dict] | None:
    """Full detail for a single method including all parameters."""
    conn = _connect(db_path)

    clause, params = _source_clause(source)
    rows = conn.execute(
        f"""SELECT m.id, m.name, m.description, m.return_type, m.return_type_ref, m.return_is_array,
                  m.element_type, m.es_only, m.uxp_type, m.uxp_note, m.es_note, m.note, src.key AS source
           FROM methods m
           JOIN classes c ON m.class_id = c.id
           JOIN sources src ON c.source_id = src.id
           WHERE c.name = ? AND m.name = ?{clause}
           ORDER BY src.key""",
        (class_name, method_name, *params),
    ).fetchall()

    if not rows:
        conn.close()
        return None

    result = []
    for row in rows:
        param_rows = conn.execute(
            """SELECT name, description, data_type, data_type_ref, is_array, is_optional, default_value,
                      min_value, max_value, measurement_unit
               FROM parameters
               WHERE method_id = ?
               ORDER BY sort_order""",
            (row["id"],),
        ).fetchall()
        signature_parts = [
            f"{p['name']}{'?' if p['is_optional'] else ''}: {p['data_type'] or 'any'}" for p in param_rows
        ]
        item = {
            "name": row["name"],
            "class_name": class_name,
            "source": row["source"],
            "description": row["description"] or "",
            "signature": f"({', '.join(signature_parts)}) -> {row['return_type'] or 'void'}",
            "return_type": row["return_type"],
            "return_type_ref": row["return_type_ref"],
            "return_is_array": bool(row["return_is_array"]),
            "element_type": row["element_type"],
            "parameters": [
                {
                    "name": p["name"],
                    "description": p["description"] or "",
                    "data_type": p["data_type"],
                    "data_type_ref": p["data_type_ref"],
                    "is_array": bool(p["is_array"]),
                    "is_optional": bool(p["is_optional"]),
                    "default_value": p["default_value"],
                    "min_value": p["min_value"],
                    "max_value": p["max_value"],
                    "measurement_unit": bool(p["measurement_unit"]),
                }
                for p in param_rows
            ],
        }
        if row["es_only"]:
            item["extendscript_only"] = True
        result.append(_with_runtime(item, row))
    conn.close()
    return result[0] if len(result) == 1 else result


def get_enum_values(enum_name: str, source: str | None = None, db_path: str | None = None) -> dict | list[dict] | None:
    """Enum values for an enumeration class."""
    conn = _connect(db_path)
    clause, params = _source_clause(source)
    rows = conn.execute(
        f"""SELECT c.id, c.name, c.description, c.uxp_note, c.es_note, c.note, src.key AS source
            FROM classes c
            JOIN sources src ON c.source_id = src.id
            WHERE c.name = ? AND c.is_enum = 1{clause}
            ORDER BY src.key""",
        (enum_name, *params),
    ).fetchall()
    if not rows:
        conn.close()
        return None

    result = []
    for row in rows:
        values = conn.execute(
            """SELECT name, description, default_value AS numeric_value
               FROM properties
               WHERE class_id = ? AND element_type = 'class'
               ORDER BY name""",
            (row["id"],),
        ).fetchall()
        item = {
            "name": row["name"],
            "source": row["source"],
            "description": row["description"] or "",
            "values": [
                {
                    "name": v["name"],
                    "description": v["description"] or "",
                    "numeric_value": v["numeric_value"],
                }
                for v in values
            ],
        }
        result.append(_with_runtime(item, row))
    conn.close()
    return result[0] if len(result) == 1 else result


def get_hierarchy(class_name: str, source: str | None = None, db_path: str | None = None) -> dict | list[dict] | None:
    """Full inheritance chain + direct subclasses."""
    conn = _connect(db_path)
    rows = _class_rows(conn, class_name, source)
    if not rows:
        conn.close()
        return None

    result = []
    for row in rows:
        ancestors = _get_ancestor_chain_names(conn, row["id"])
        subclasses = [
            r["name"]
            for r in conn.execute(
                """SELECT name FROM classes
                   WHERE superclass_name = ?
                     AND source_id = (SELECT source_id FROM classes WHERE id = ?)
                   ORDER BY name""",
                (class_name, row["id"]),
            ).fetchall()
        ]
        item = {
            "class_name": class_name,
            "source": row["source"],
            "ancestors": ancestors,
            "direct_subclasses": subclasses,
        }
        if row["collection_of"]:
            item["collection_of"] = row["collection_of"]
        result.append(item)
    conn.close()
    return result[0] if len(result) == 1 else result


def search_dom(
    query: str,
    source: str | None = None,
    max_results: int = 20,
    runtime: str | None = None,
    db_path: str | None = None,
) -> list[dict]:
    """Full-text search across all entities.

    runtime='uxp' limits results to the InDesign DOM and drops the [] index accessor.
    """
    conn = _connect(db_path)
    terms = [t.replace('"', "") for t in query.split()]
    fts_query = " ".join(f'"{t}"*' for t in terms if t)
    if not fts_query:
        conn.close()
        return []

    sql = """SELECT entity_type, entity_name, parent_name, description, source, rank
             FROM dom_search
             WHERE dom_search MATCH ?"""
    params: list = [fts_query]
    if source:
        sql += " AND source = ?"
        params.append(source)
    elif runtime == "uxp":
        sql += " AND source = 'dom'"
    if runtime == "uxp":
        sql += " AND NOT (entity_type = 'method' AND entity_name = '[]')"
    sql += " ORDER BY rank LIMIT ?"
    params.append(max_results)

    rows = conn.execute(sql, params).fetchall()
    if not rows and re.search(r"[^\w\s]", query):
        # The FTS tokenizer drops symbols such as "$"; fall back to an exact name match.
        sql = "SELECT entity_type, entity_name, parent_name, description, source FROM dom_search WHERE entity_name = ?"
        params = [query.strip()]
        if source:
            sql += " AND source = ?"
            params.append(source)
        sql += " LIMIT ?"
        params.append(max_results)
        rows = conn.execute(sql, params).fetchall()
    conn.close()

    return [
        {
            "entity_type": r["entity_type"],
            "entity_name": r["entity_name"],
            "parent_name": r["parent_name"],
            "description": r["description"] or "",
            "source": r["source"],
        }
        for r in rows
    ]


def list_classes(
    type_filter: str = "all",
    source: str | None = None,
    filter_text: str | None = None,
    db_path: str | None = None,
) -> list[dict]:
    """Class overview, filterable by kind, source and name substring."""
    conn = _connect(db_path)

    query = """
        SELECT c.name, c.kind, c.description, c.collection_of, src.key AS source
        FROM classes c
        JOIN sources src ON c.source_id = src.id
        WHERE 1=1
    """
    params: list = []

    if source:
        query += " AND src.key = ?"
        params.append(source)
    query += CLASS_TYPE_FILTERS.get(type_filter, "")
    if filter_text:
        query += " AND c.name LIKE ?"
        params.append(f"%{filter_text}%")

    query += " ORDER BY src.key, c.name"
    rows = conn.execute(query, params).fetchall()
    conn.close()

    result = []
    for r in rows:
        item = {
            "name": r["name"],
            "kind": r["kind"],
            "description": (r["description"] or "")[:120],
            "source": r["source"],
        }
        if r["collection_of"]:
            item["collection_of"] = r["collection_of"]
        result.append(item)
    return result


def dom_info(db_path: str | None = None) -> dict:
    """DB metadata and statistics."""
    conn = _connect(db_path)
    meta = {row["key"]: row["value"] for row in conn.execute("SELECT key, value FROM db_meta").fetchall()}

    counts = {}
    for table in ("classes", "properties", "methods", "parameters", "events"):
        (counts[table],) = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    for kind in ("object", "collection", "enum"):
        (counts[f"{kind}s"],) = conn.execute("SELECT COUNT(*) FROM classes WHERE kind = ?", (kind,)).fetchone()
    source_rows = conn.execute(
        """SELECT src.key AS source, src.version, src.generated, COUNT(c.id) AS class_count
           FROM sources src
           LEFT JOIN classes c ON c.source_id = src.id
           GROUP BY src.id
           ORDER BY src.key"""
    ).fetchall()
    conn.close()

    return {
        "dom_version": meta.get("dom_version", ""),
        "dom_title": meta.get("dom_title", ""),
        "dom_target": meta.get("dom_target", ""),
        "dom_generated": meta.get("dom_generated", ""),
        "source_origin": meta.get("source_origin", ""),
        "source_files": meta.get("source_files", ""),
        "source_keys": meta.get("source_keys", ""),
        "build_timestamp": meta.get("build_timestamp", ""),
        "parser_version": meta.get("parser_version", ""),
        "sources": [
            {"source": r["source"], "version": r["version"], "generated": r["generated"], "class_count": r["class_count"]}
            for r in source_rows
        ],
        "counts": counts,
    }


def list_sources(db_path: str | None = None) -> list[dict]:
    """List available sources with entity counts."""
    conn = _connect(db_path)
    rows = conn.execute(
        """SELECT src.key AS source, src.label, src.url, src.version, src.title, src.generated, src.uxp_note,
                  src.runtimes, src.uxp_guide_url,
                  (SELECT COUNT(*) FROM classes c WHERE c.source_id = src.id) AS classes,
                  (SELECT COUNT(*) FROM properties p JOIN classes c ON p.class_id = c.id
                    WHERE c.source_id = src.id) AS properties,
                  (SELECT COUNT(*) FROM methods m JOIN classes c ON m.class_id = c.id
                    WHERE c.source_id = src.id) AS methods
           FROM sources src
           ORDER BY src.key"""
    ).fetchall()
    conn.close()
    result = []
    for r in rows:
        item = {
            "source": r["source"],
            "label": r["label"],
            "url": r["url"] or "",
            "version": r["title"] or r["version"] or "",
            "generated": r["generated"] or "",
            "runtimes": (r["runtimes"] or "extendscript").split(","),
            "counts": {
                "classes": r["classes"],
                "properties": r["properties"],
                "methods": r["methods"],
            },
        }
        if r["uxp_guide_url"]:
            item["uxp_guide_url"] = r["uxp_guide_url"]
        result.append(_with_runtime(item, r))
    return result


def knowledge_overview(db_path: str | None = None) -> dict:
    """Return capability overview for multi-source scripting lookups."""
    info = dom_info(db_path=db_path)
    return {
        "dom_version": info["dom_title"] or info["dom_version"],
        "dom_generated": info["dom_generated"],
        "source_origin": info["source_origin"],
        "sources": list_sources(db_path=db_path),
        "extendscript_specials": ["$", "UnitValue", "File", "Folder", "Socket", "XML", "XMLList", "RegExp"],
        "scriptui_note": (
            "ScriptUI is legacy technology. Prefer UXP for new UI development. "
            "ScriptUI documentation remains useful for small dialogs and maintaining existing scripts."
        ),
        "runtime_note": (
            "The InDesign DOM is shared by ExtendScript and UXP. Entries carrying 'uxp_note' behave differently "
            "in UXP; 'uxp_type' is the type under UXP (e.g. Promise<File> for File/Folder properties); "
            "'extendscript_only' members (e.g. the [] index accessor) do not exist in UXP. "
            "Pass runtime='uxp' to get_properties/get_methods/search_dom to hide them. "
            "General UXP rules (require, item(), equals(), script.args, script.setResult()) are at "
            "uxp_guide_url of the 'dom' source; TypeScript types for UXP: npm package indesign-uxp-types. "
            "Sources 'javascript' and 'scriptui' are ExtendScript only."
        ),
        "online_docs": (
            "Every class has a Markdown page at doc_url (indesignjs.de by Gregor Fellenz) for long-form reading."
        ),
        "lookup_order": ["knowledge_overview", "search_dom(source=...)", "lookup_class(source=...)", "indesign-exec.run_jsx"],
        "known_name_collisions": ["Window", "Button", "Event", "Events", "Group", "ListBox", "Panel", "RadioButton",
                                  "StaticText"],
    }


def _get_ancestor_chain_ids(conn: sqlite3.Connection, class_id: int) -> list[int]:
    """Walk superclass chain upward by IDs within the same source."""
    chain = []
    current_id = class_id
    seen = set()

    while current_id and current_id not in seen:
        chain.append(current_id)
        seen.add(current_id)
        row = conn.execute(
            """SELECT parent.id AS parent_id
               FROM classes child
               LEFT JOIN classes parent
                 ON parent.name = child.superclass_name
                AND parent.source_id = child.source_id
               WHERE child.id = ?""",
            (current_id,),
        ).fetchone()
        current_id = row["parent_id"] if row and row["parent_id"] else None
    return chain


def _get_ancestor_chain_names(conn: sqlite3.Connection, class_id: int) -> list[str]:
    """Walk superclass chain upward by class names."""
    names = []
    current_id = class_id
    seen = set()

    while current_id and current_id not in seen:
        seen.add(current_id)
        row = conn.execute(
            """SELECT child.name AS name, parent.id AS parent_id
               FROM classes child
               LEFT JOIN classes parent
                 ON parent.name = child.superclass_name
                AND parent.source_id = child.source_id
               WHERE child.id = ?""",
            (current_id,),
        ).fetchone()
        if not row:
            break
        names.append(row["name"])
        current_id = row["parent_id"]
    return names
