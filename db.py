"""
InDesign DOM Database Query Layer.

Provides query functions for MCP tools with multi-source support.
"""

import sqlite3
from pathlib import Path

DEFAULT_DB = Path(__file__).parent / "extendscript.db"


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


def _class_rows(conn: sqlite3.Connection, name: str, source: str | None = None):
    """Fetch class rows by class name and optional source."""
    clause, params = _source_clause(source)
    return conn.execute(
        f"""SELECT c.id, c.name, c.is_enum, c.is_dynamic, c.description, c.description_long,
                  c.superclass_name, s.name AS suite_name, src.key AS source
           FROM classes c
           LEFT JOIN suites s ON c.suite_id = s.id
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
                (name, class_id),
            ).fetchall()
        ]
        result.append(
            {
                "name": row["name"],
                "source": row["source"],
                "suite": row["suite_name"] or "",
                "is_enum": bool(row["is_enum"]),
                "is_dynamic": bool(row["is_dynamic"]),
                "description": row["description"] or "",
                "description_long": row["description_long"] or "",
                "superclass": row["superclass_name"],
                "property_count": prop_count,
                "method_count": method_count,
                "direct_subclasses": subclasses,
            }
        )
    conn.close()
    return result[0] if len(result) == 1 else result


def get_properties(
    class_name: str,
    source: str | None = None,
    filter_text: str | None = None,
    include_inherited: bool = False,
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

    query += " ORDER BY src.key, c.name, p.element_type, p.name"

    rows = conn.execute(query, params).fetchall()
    conn.close()

    return [
        {
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
            "defined_in": r["class_name"],
            "source": r["source"],
        }
        for r in rows
    ]


def get_methods(
    class_name: str,
    source: str | None = None,
    filter_text: str | None = None,
    include_inherited: bool = False,
    db_path: str | None = None,
) -> list[dict]:
    """Methods of one class with short signatures."""
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
               m.return_type_ref, m.element_type, c.name AS class_name, src.key AS source
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
            sig_parts.append(f"{p['name']}: {ptype}{opt}")

        ret = r["return_type"] or "void"
        if r["return_is_array"]:
            ret += "[]"
        signature = f"({', '.join(sig_parts)}) -> {ret}"

        result.append(
            {
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
        )

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
                  m.element_type, src.key AS source
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
            """SELECT name, description, data_type, data_type_ref, is_array, is_optional, default_value
               FROM parameters
               WHERE method_id = ?
               ORDER BY sort_order""",
            (row["id"],),
        ).fetchall()
        result.append(
            {
                "name": row["name"],
                "class_name": class_name,
                "source": row["source"],
                "description": row["description"] or "",
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
                    }
                    for p in param_rows
                ],
            }
        )
    conn.close()
    return result[0] if len(result) == 1 else result


def get_enum_values(enum_name: str, source: str | None = None, db_path: str | None = None) -> dict | list[dict] | None:
    """Enum values for an enumeration class."""
    conn = _connect(db_path)
    clause, params = _source_clause(source)
    rows = conn.execute(
        f"""SELECT c.id, c.name, c.description, src.key AS source
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
        result.append(
            {
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
        )
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
        result.append(
            {
                "class_name": class_name,
                "source": row["source"],
                "ancestors": ancestors,
                "direct_subclasses": subclasses,
            }
        )
    conn.close()
    return result[0] if len(result) == 1 else result


def search_dom(
    query: str,
    source: str | None = None,
    max_results: int = 20,
    db_path: str | None = None,
) -> list[dict]:
    """Full-text search across all entities."""
    conn = _connect(db_path)
    fts_query = " ".join(f"{term}*" for term in query.split())

    if source:
        rows = conn.execute(
            """SELECT entity_type, entity_name, parent_name, description, source, rank
               FROM dom_search
               WHERE dom_search MATCH ?
                 AND source = ?
               ORDER BY rank
               LIMIT ?""",
            (fts_query, source, max_results),
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT entity_type, entity_name, parent_name, description, source, rank
               FROM dom_search
               WHERE dom_search MATCH ?
               ORDER BY rank
               LIMIT ?""",
            (fts_query, max_results),
        ).fetchall()

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
    suite: str | None = None,
    type_filter: str = "all",
    source: str | None = None,
    db_path: str | None = None,
) -> list[dict]:
    """Class overview, filterable by suite/type/source."""
    conn = _connect(db_path)

    query = """
        SELECT c.name, c.is_enum, c.description, s.name AS suite_name, src.key AS source
        FROM classes c
        LEFT JOIN suites s ON c.suite_id = s.id
        JOIN sources src ON c.source_id = src.id
        WHERE 1=1
    """
    params: list = []

    if suite:
        query += " AND s.name = ?"
        params.append(suite)
    if source:
        query += " AND src.key = ?"
        params.append(source)

    if type_filter == "class":
        query += " AND c.is_enum = 0"
    elif type_filter == "enum":
        query += " AND c.is_enum = 1"

    query += " ORDER BY src.key, c.name"
    rows = conn.execute(query, params).fetchall()
    conn.close()

    return [
        {
            "name": r["name"],
            "suite": r["suite_name"] or "",
            "is_enum": bool(r["is_enum"]),
            "description": (r["description"] or "")[:120],
            "source": r["source"],
        }
        for r in rows
    ]


def dom_info(db_path: str | None = None) -> dict:
    """DB metadata and statistics."""
    conn = _connect(db_path)
    meta = {row["key"]: row["value"] for row in conn.execute("SELECT key, value FROM db_meta").fetchall()}

    (suite_count,) = conn.execute("SELECT COUNT(*) FROM suites").fetchone()
    (class_count,) = conn.execute("SELECT COUNT(*) FROM classes").fetchone()
    (enum_count,) = conn.execute("SELECT COUNT(*) FROM classes WHERE is_enum = 1").fetchone()
    (prop_count,) = conn.execute("SELECT COUNT(*) FROM properties").fetchone()
    (method_count,) = conn.execute("SELECT COUNT(*) FROM methods").fetchone()
    (param_count,) = conn.execute("SELECT COUNT(*) FROM parameters").fetchone()
    source_rows = conn.execute(
        """SELECT src.key AS source, COUNT(c.id) AS class_count
           FROM sources src
           LEFT JOIN classes c ON c.source_id = src.id
           GROUP BY src.key
           ORDER BY src.key"""
    ).fetchall()
    conn.close()

    return {
        "dom_version": meta.get("dom_version", ""),
        "dom_title": meta.get("dom_title", ""),
        "source_file": meta.get("source_file", ""),
        "source_files": meta.get("source_files", ""),
        "source_keys": meta.get("source_keys", ""),
        "build_timestamp": meta.get("build_timestamp", ""),
        "parser_version": meta.get("parser_version", ""),
        "sources": [{"source": r["source"], "class_count": r["class_count"]} for r in source_rows],
        "counts": {
            "suites": suite_count,
            "classes": class_count,
            "enums": enum_count,
            "regular_classes": class_count - enum_count,
            "properties": prop_count,
            "methods": method_count,
            "parameters": param_count,
        },
    }


def list_sources(db_path: str | None = None) -> list[dict]:
    """List available sources with entity counts."""
    conn = _connect(db_path)
    rows = conn.execute(
        """SELECT src.key AS source, src.label, src.file,
                  COUNT(DISTINCT c.id) AS classes,
                  COUNT(DISTINCT p.id) AS properties,
                  COUNT(DISTINCT m.id) AS methods
           FROM sources src
           LEFT JOIN classes c ON c.source_id = src.id
           LEFT JOIN properties p ON p.class_id = c.id
           LEFT JOIN methods m ON m.class_id = c.id
           GROUP BY src.id
           ORDER BY src.key"""
    ).fetchall()
    conn.close()
    return [
        {
            "source": r["source"],
            "label": r["label"],
            "file": r["file"] or "",
            "counts": {
                "classes": r["classes"],
                "properties": r["properties"],
                "methods": r["methods"],
            },
        }
        for r in rows
    ]


def knowledge_overview(db_path: str | None = None) -> dict:
    """Return capability overview for multi-source scripting lookups."""
    return {
        "sources": list_sources(db_path=db_path),
        "extendscript_specials": ["$", "UnitValue", "File", "Folder", "Socket", "XML", "XMLList", "RegExp"],
        "scriptui_note": (
            "ScriptUI is legacy technology. Prefer UXP for new UI development. "
            "ScriptUI documentation remains useful for small dialogs and maintaining existing scripts."
        ),
        "lookup_order": ["knowledge_overview", "search_dom(source=...)", "lookup_class(source=...)", "indesign-exec.run_jsx"],
        "known_name_collisions": ["Window", "Group", "Panel", "Event"],
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
