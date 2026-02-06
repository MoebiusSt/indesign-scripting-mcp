"""
InDesign DOM Database Query Layer.

Provides query functions for all 9 MCP tools:
  1. lookup_class
  2. get_properties
  3. get_methods
  4. get_method_detail
  5. get_enum_values
  6. get_hierarchy
  7. search_dom
  8. list_classes
  9. dom_info
"""

import sqlite3
from pathlib import Path

DEFAULT_DB = Path(__file__).parent / "indesign_dom.db"


def _connect(db_path: str | None = None) -> sqlite3.Connection:
    """Open a read-only connection to the DOM database."""
    path = db_path or str(DEFAULT_DB)
    conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


# ---------------------------------------------------------------------------
# 1. lookup_class
# ---------------------------------------------------------------------------

def lookup_class(name: str, db_path: str | None = None) -> dict | None:
    """Full class info: suite, superclass, description, counts, subclasses."""
    conn = _connect(db_path)
    row = conn.execute(
        """SELECT c.id, c.name, c.is_enum, c.is_dynamic, c.description,
                  c.superclass_name, s.name AS suite_name
           FROM classes c
           LEFT JOIN suites s ON c.suite_id = s.id
           WHERE c.name = ?""",
        (name,),
    ).fetchone()

    if not row:
        conn.close()
        return None

    class_id = row["id"]

    (prop_count,) = conn.execute(
        "SELECT COUNT(*) FROM properties WHERE class_id = ?", (class_id,)
    ).fetchone()

    (method_count,) = conn.execute(
        "SELECT COUNT(*) FROM methods WHERE class_id = ?", (class_id,)
    ).fetchone()

    subclasses = [
        r["name"]
        for r in conn.execute(
            "SELECT name FROM classes WHERE superclass_name = ? ORDER BY name",
            (name,),
        ).fetchall()
    ]

    conn.close()

    return {
        "name": row["name"],
        "suite": row["suite_name"] or "",
        "is_enum": bool(row["is_enum"]),
        "is_dynamic": bool(row["is_dynamic"]),
        "description": row["description"] or "",
        "superclass": row["superclass_name"],
        "property_count": prop_count,
        "method_count": method_count,
        "direct_subclasses": subclasses,
    }


# ---------------------------------------------------------------------------
# 2. get_properties
# ---------------------------------------------------------------------------

def get_properties(
    class_name: str,
    filter_text: str | None = None,
    include_inherited: bool = False,
    db_path: str | None = None,
) -> list[dict]:
    """Properties of a class, optionally filtered and with inheritance."""
    conn = _connect(db_path)

    class_names = [class_name]
    if include_inherited:
        class_names = _get_ancestor_chain(conn, class_name)

    placeholders = ",".join("?" for _ in class_names)
    query = f"""
        SELECT p.name, p.description, p.data_type, p.is_array, p.is_readonly,
               p.element_type, p.default_value, p.min_value, p.max_value,
               c.name AS class_name
        FROM properties p
        JOIN classes c ON p.class_id = c.id
        WHERE c.name IN ({placeholders})
    """
    params: list = list(class_names)

    if filter_text:
        query += " AND (p.name LIKE ? OR p.description LIKE ?)"
        like = f"%{filter_text}%"
        params.extend([like, like])

    query += " ORDER BY c.name, p.element_type, p.name"

    rows = conn.execute(query, params).fetchall()
    conn.close()

    return [
        {
            "name": r["name"],
            "description": r["description"] or "",
            "data_type": r["data_type"],
            "is_array": bool(r["is_array"]),
            "is_readonly": bool(r["is_readonly"]),
            "element_type": r["element_type"],
            "default_value": r["default_value"],
            "min_value": r["min_value"],
            "max_value": r["max_value"],
            "defined_in": r["class_name"],
        }
        for r in rows
    ]


# ---------------------------------------------------------------------------
# 3. get_methods
# ---------------------------------------------------------------------------

def get_methods(
    class_name: str,
    filter_text: str | None = None,
    include_inherited: bool = False,
    db_path: str | None = None,
) -> list[dict]:
    """Methods of a class with short signatures."""
    conn = _connect(db_path)

    class_names = [class_name]
    if include_inherited:
        class_names = _get_ancestor_chain(conn, class_name)

    placeholders = ",".join("?" for _ in class_names)
    query = f"""
        SELECT m.id, m.name, m.description, m.return_type, m.return_is_array,
               m.element_type, c.name AS class_name
        FROM methods m
        JOIN classes c ON m.class_id = c.id
        WHERE c.name IN ({placeholders})
    """
    params: list = list(class_names)

    if filter_text:
        query += " AND (m.name LIKE ? OR m.description LIKE ?)"
        like = f"%{filter_text}%"
        params.extend([like, like])

    query += " ORDER BY c.name, m.element_type, m.name"

    rows = conn.execute(query, params).fetchall()
    result = []

    for r in rows:
        # Build short signature
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

        result.append({
            "name": r["name"],
            "description": r["description"] or "",
            "signature": signature,
            "return_type": r["return_type"],
            "return_is_array": bool(r["return_is_array"]),
            "element_type": r["element_type"],
            "defined_in": r["class_name"],
        })

    conn.close()
    return result


# ---------------------------------------------------------------------------
# 4. get_method_detail
# ---------------------------------------------------------------------------

def get_method_detail(
    class_name: str,
    method_name: str,
    db_path: str | None = None,
) -> dict | None:
    """Full detail for a single method including all parameters."""
    conn = _connect(db_path)

    row = conn.execute(
        """SELECT m.id, m.name, m.description, m.return_type, m.return_is_array,
                  m.element_type
           FROM methods m
           JOIN classes c ON m.class_id = c.id
           WHERE c.name = ? AND m.name = ?""",
        (class_name, method_name),
    ).fetchone()

    if not row:
        conn.close()
        return None

    params = conn.execute(
        """SELECT name, description, data_type, is_array, is_optional, default_value
           FROM parameters
           WHERE method_id = ?
           ORDER BY sort_order""",
        (row["id"],),
    ).fetchall()

    conn.close()

    return {
        "name": row["name"],
        "class_name": class_name,
        "description": row["description"] or "",
        "return_type": row["return_type"],
        "return_is_array": bool(row["return_is_array"]),
        "element_type": row["element_type"],
        "parameters": [
            {
                "name": p["name"],
                "description": p["description"] or "",
                "data_type": p["data_type"],
                "is_array": bool(p["is_array"]),
                "is_optional": bool(p["is_optional"]),
                "default_value": p["default_value"],
            }
            for p in params
        ],
    }


# ---------------------------------------------------------------------------
# 5. get_enum_values
# ---------------------------------------------------------------------------

def get_enum_values(enum_name: str, db_path: str | None = None) -> dict | None:
    """Enum values for an enumeration class."""
    conn = _connect(db_path)

    row = conn.execute(
        "SELECT id, name, description FROM classes WHERE name = ? AND is_enum = 1",
        (enum_name,),
    ).fetchone()

    if not row:
        conn.close()
        return None

    values = conn.execute(
        """SELECT name, description, default_value AS numeric_value
           FROM properties
           WHERE class_id = ? AND element_type = 'class'
           ORDER BY name""",
        (row["id"],),
    ).fetchall()

    conn.close()

    return {
        "name": row["name"],
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


# ---------------------------------------------------------------------------
# 6. get_hierarchy
# ---------------------------------------------------------------------------

def get_hierarchy(class_name: str, db_path: str | None = None) -> dict | None:
    """Full inheritance chain + direct subclasses."""
    conn = _connect(db_path)

    # Check class exists
    row = conn.execute(
        "SELECT name, superclass_name FROM classes WHERE name = ?", (class_name,)
    ).fetchone()

    if not row:
        conn.close()
        return None

    # Walk up the ancestor chain
    ancestors = _get_ancestor_chain(conn, class_name)

    # Direct subclasses
    subclasses = [
        r["name"]
        for r in conn.execute(
            "SELECT name FROM classes WHERE superclass_name = ? ORDER BY name",
            (class_name,),
        ).fetchall()
    ]

    conn.close()

    return {
        "class_name": class_name,
        "ancestors": ancestors,  # [self, parent, grandparent, ...]
        "direct_subclasses": subclasses,
    }


# ---------------------------------------------------------------------------
# 7. search_dom
# ---------------------------------------------------------------------------

def search_dom(
    query: str,
    max_results: int = 20,
    db_path: str | None = None,
) -> list[dict]:
    """Full-text search across all entities."""
    conn = _connect(db_path)

    # Prepare FTS5 query - add * for prefix matching
    fts_query = " ".join(f"{term}*" for term in query.split())

    rows = conn.execute(
        """SELECT entity_type, entity_name, parent_name, description,
                  rank
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
        }
        for r in rows
    ]


# ---------------------------------------------------------------------------
# 8. list_classes
# ---------------------------------------------------------------------------

def list_classes(
    suite: str | None = None,
    type_filter: str = "all",
    db_path: str | None = None,
) -> list[dict]:
    """Class overview, filterable by suite and type."""
    conn = _connect(db_path)

    query = """
        SELECT c.name, c.is_enum, c.description, s.name AS suite_name
        FROM classes c
        LEFT JOIN suites s ON c.suite_id = s.id
        WHERE 1=1
    """
    params: list = []

    if suite:
        query += " AND s.name = ?"
        params.append(suite)

    if type_filter == "class":
        query += " AND c.is_enum = 0"
    elif type_filter == "enum":
        query += " AND c.is_enum = 1"

    query += " ORDER BY c.name"

    rows = conn.execute(query, params).fetchall()
    conn.close()

    return [
        {
            "name": r["name"],
            "suite": r["suite_name"] or "",
            "is_enum": bool(r["is_enum"]),
            "description": (r["description"] or "")[:120],
        }
        for r in rows
    ]


# ---------------------------------------------------------------------------
# 9. dom_info
# ---------------------------------------------------------------------------

def dom_info(db_path: str | None = None) -> dict:
    """DB metadata and statistics."""
    conn = _connect(db_path)

    meta = {}
    for row in conn.execute("SELECT key, value FROM db_meta").fetchall():
        meta[row["key"]] = row["value"]

    (suite_count,) = conn.execute("SELECT COUNT(*) FROM suites").fetchone()
    (class_count,) = conn.execute("SELECT COUNT(*) FROM classes").fetchone()
    (enum_count,) = conn.execute("SELECT COUNT(*) FROM classes WHERE is_enum = 1").fetchone()
    (prop_count,) = conn.execute("SELECT COUNT(*) FROM properties").fetchone()
    (method_count,) = conn.execute("SELECT COUNT(*) FROM methods").fetchone()
    (param_count,) = conn.execute("SELECT COUNT(*) FROM parameters").fetchone()

    conn.close()

    return {
        "dom_version": meta.get("dom_version", ""),
        "dom_title": meta.get("dom_title", ""),
        "source_file": meta.get("source_file", ""),
        "build_timestamp": meta.get("build_timestamp", ""),
        "parser_version": meta.get("parser_version", ""),
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_ancestor_chain(conn: sqlite3.Connection, class_name: str) -> list[str]:
    """Walk superclass chain upward. Returns [self, parent, grandparent, ...]."""
    chain = []
    current = class_name
    seen = set()

    while current and current not in seen:
        chain.append(current)
        seen.add(current)
        row = conn.execute(
            "SELECT superclass_name FROM classes WHERE name = ?", (current,)
        ).fetchone()
        if row and row["superclass_name"]:
            current = row["superclass_name"]
        else:
            break

    return chain
