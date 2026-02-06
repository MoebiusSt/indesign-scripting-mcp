"""
InDesign DOM Parser – Parses OMV-XML files into SQLite database.

Handles:
- Suite navigation from <map> element
- Class definitions (regular + enum)
- Properties, methods, parameters
- FTS5 full-text search index
- Analysis reports and validation
"""

import xml.etree.ElementTree as ET
import sqlite3
import os
import re
from datetime import datetime
from pathlib import Path

PARSER_VERSION = "1.0.0"


# ---------------------------------------------------------------------------
# XML Parsing
# ---------------------------------------------------------------------------

def parse_xml(xml_path: str) -> dict:
    """Parse an OMV-XML file and return structured data.

    Returns dict with keys:
        version, title, timestamp, suites, classes
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Extract metadata from <map>
    map_el = root.find("map")
    if map_el is None:
        raise ValueError("No <map> element found in XML")

    version = map_el.get("name", "")
    title = map_el.get("title", "")
    timestamp = map_el.get("time", "")

    # Parse suite navigation
    suites = _parse_suites(map_el)

    # Build reverse lookup: class_name -> suite_name
    class_to_suite = {}
    for suite_name, class_names in suites.items():
        for cn in class_names:
            class_to_suite[cn] = suite_name

    # Parse all class definitions
    package = root.find("package")
    if package is None:
        raise ValueError("No <package> element found in XML")

    classes = []
    for classdef in package.findall("classdef"):
        cls = _parse_classdef(classdef)
        cls["suite"] = class_to_suite.get(cls["name"], "")
        classes.append(cls)

    return {
        "version": version,
        "title": title,
        "timestamp": timestamp,
        "suites": suites,
        "classes": classes,
    }


def _parse_suites(map_el) -> dict:
    """Parse suite navigation from <map> topicrefs.

    Returns dict: suite_name -> [class_name, ...]
    """
    suites = {}
    for suite_ref in map_el.findall("topicref"):
        suite_name = suite_ref.get("navtitle", "")
        class_names = []
        for class_ref in suite_ref.findall("topicref"):
            href = class_ref.get("href", "")
            # href format: #/ClassName
            if href.startswith("#/"):
                class_names.append(href[2:])
        if suite_name:
            suites[suite_name] = class_names
    return suites


def _parse_classdef(classdef) -> dict:
    """Parse a single <classdef> element."""
    name = classdef.get("name", "")
    is_enum = classdef.get("enumeration") == "true"
    is_dynamic = classdef.get("dynamic") == "true"

    shortdesc_el = classdef.find("shortdesc")
    description = shortdesc_el.text.strip() if shortdesc_el is not None and shortdesc_el.text else ""

    superclass_el = classdef.find("superclass")
    superclass_name = superclass_el.text.strip() if superclass_el is not None and superclass_el.text else None

    properties = []
    methods = []

    for elements in classdef.findall("elements"):
        element_type = elements.get("type", "instance")  # 'class' or 'instance'

        for prop in elements.findall("property"):
            properties.append(_parse_property(prop, element_type))

        for meth in elements.findall("method"):
            methods.append(_parse_method(meth, element_type))

    return {
        "name": name,
        "is_enum": is_enum,
        "is_dynamic": is_dynamic,
        "description": description,
        "superclass_name": superclass_name,
        "properties": properties,
        "methods": methods,
    }


def _parse_property(prop, element_type: str) -> dict:
    """Parse a single <property> element."""
    name = prop.get("name", "")
    rwaccess = prop.get("rwaccess", "")
    is_readonly = rwaccess == "readonly"

    shortdesc_el = prop.find("shortdesc")
    description = shortdesc_el.text.strip() if shortdesc_el is not None and shortdesc_el.text else ""

    datatype = prop.find("datatype")
    data_type, is_array, default_value, min_value, max_value = _parse_datatype(datatype)

    return {
        "name": name,
        "description": description,
        "data_type": data_type,
        "is_array": is_array,
        "is_readonly": is_readonly,
        "element_type": element_type,
        "default_value": default_value,
        "min_value": min_value,
        "max_value": max_value,
    }


def _parse_method(meth, element_type: str) -> dict:
    """Parse a single <method> element."""
    name = meth.get("name", "")

    shortdesc_el = meth.find("shortdesc")
    description = shortdesc_el.text.strip() if shortdesc_el is not None and shortdesc_el.text else ""

    # Return type
    datatype = meth.find("datatype")
    return_type = None
    return_is_array = False
    if datatype is not None:
        return_type, return_is_array, _, _, _ = _parse_datatype(datatype)

    # Parameters
    parameters = []
    params_el = meth.find("parameters")
    if params_el is not None:
        for idx, param in enumerate(params_el.findall("parameter")):
            parameters.append(_parse_parameter(param, idx))

    return {
        "name": name,
        "description": description,
        "return_type": return_type,
        "return_is_array": return_is_array,
        "element_type": element_type,
        "parameters": parameters,
    }


def _parse_parameter(param, sort_order: int) -> dict:
    """Parse a single <parameter> element."""
    name = param.get("name", "")
    is_optional = param.get("optional") == "true"

    shortdesc_el = param.find("shortdesc")
    description = shortdesc_el.text.strip() if shortdesc_el is not None and shortdesc_el.text else ""

    # Check description for "(Optional)" hint even when attribute is missing
    if not is_optional and description and "(Optional)" in description:
        is_optional = True

    datatype = param.find("datatype")
    data_type, is_array, default_value, _, _ = _parse_datatype(datatype)

    return {
        "name": name,
        "description": description,
        "data_type": data_type,
        "is_array": is_array,
        "is_optional": is_optional,
        "default_value": default_value,
        "sort_order": sort_order,
    }


def _parse_datatype(datatype):
    """Parse a <datatype> element.

    Returns (type_str, is_array, default_value, min_value, max_value).
    """
    if datatype is None:
        return None, False, None, None, None

    type_el = datatype.find("type")
    type_str = None
    if type_el is not None:
        # Handle varies=any attribute
        varies = type_el.get("varies")
        if varies:
            type_str = f"varies={varies}"
        elif type_el.text:
            type_str = type_el.text.strip()

    is_array = datatype.find("array") is not None

    value_el = datatype.find("value")
    default_value = value_el.text.strip() if value_el is not None and value_el.text else None

    min_el = datatype.find("min")
    min_value = min_el.text.strip() if min_el is not None and min_el.text else None

    max_el = datatype.find("max")
    max_value = max_el.text.strip() if max_el is not None and max_el.text else None

    return type_str, is_array, default_value, min_value, max_value


# ---------------------------------------------------------------------------
# Analysis Report
# ---------------------------------------------------------------------------

def analyze(data: dict, xml_path: str) -> dict:
    """Generate an analysis report from parsed data.

    Returns a stats dict and prints a formatted report.
    """
    classes = data["classes"]
    regular = [c for c in classes if not c["is_enum"]]
    enums = [c for c in classes if c["is_enum"]]

    total_properties = sum(len(c["properties"]) for c in classes)
    total_methods = sum(len(c["methods"]) for c in classes)
    total_parameters = sum(
        len(m["parameters"])
        for c in classes
        for m in c["methods"]
    )
    superclass_count = sum(1 for c in classes if c["superclass_name"])
    polymorphic_count = sum(
        1
        for c in classes
        for p in c["properties"]
        if p["data_type"] and "varies" in str(p["data_type"])
    )

    # Top classes by size
    class_sizes = []
    for c in regular:
        n_props = len(c["properties"])
        n_meths = len(c["methods"])
        class_sizes.append((c["name"], n_props, n_meths))
    class_sizes.sort(key=lambda x: x[1] + x[2], reverse=True)

    stats = {
        "source_file": os.path.basename(xml_path),
        "version": data["version"],
        "title": data["title"],
        "suite_count": len(data["suites"]),
        "class_count": len(classes),
        "regular_count": len(regular),
        "enum_count": len(enums),
        "property_count": total_properties,
        "method_count": total_methods,
        "parameter_count": total_parameters,
        "superclass_count": superclass_count,
        "polymorphic_count": polymorphic_count,
        "top_classes": class_sizes[:5],
    }

    return stats


def print_report(stats: dict):
    """Print a formatted analysis report."""
    sep = "=" * 55
    print(sep)
    print(f"  InDesign DOM Analysis Report")
    print(f"  Source: {stats['source_file']}")
    print(f"  Version: {stats['version']} | Title: {stats['title']}")
    print(sep)
    print(f"  Suites:              {stats['suite_count']:>5}")
    print(f"  Classes (total):     {stats['class_count']:>5}")
    print(f"    Regular classes:   {stats['regular_count']:>5}")
    print(f"    Enumerations:      {stats['enum_count']:>5}")
    print(f"  Properties:          {stats['property_count']:>5}")
    print(f"  Methods:             {stats['method_count']:>5}")
    print(f"  Parameters:          {stats['parameter_count']:>5}")
    print(f"  Superclass refs:     {stats['superclass_count']:>5}")
    print(f"  Polymorphic types:   {stats['polymorphic_count']:>5}")
    print(sep)
    print(f"  Top classes by size:")
    for name, props, meths in stats["top_classes"]:
        print(f"    {name:<20} {props:>4} properties, {meths:>4} methods")
    print(sep)


# ---------------------------------------------------------------------------
# Database Generation
# ---------------------------------------------------------------------------

DB_SCHEMA = """
-- Metadata
CREATE TABLE IF NOT EXISTS db_meta (
    key         TEXT PRIMARY KEY,
    value       TEXT NOT NULL
);

-- Suites
CREATE TABLE IF NOT EXISTS suites (
    id          INTEGER PRIMARY KEY,
    name        TEXT NOT NULL UNIQUE
);

-- Classes
CREATE TABLE IF NOT EXISTS classes (
    id              INTEGER PRIMARY KEY,
    name            TEXT NOT NULL UNIQUE,
    suite_id        INTEGER REFERENCES suites(id),
    is_enum         BOOLEAN NOT NULL DEFAULT 0,
    is_dynamic      BOOLEAN NOT NULL DEFAULT 0,
    description     TEXT,
    superclass_name TEXT
);

-- Properties
CREATE TABLE IF NOT EXISTS properties (
    id              INTEGER PRIMARY KEY,
    class_id        INTEGER NOT NULL REFERENCES classes(id),
    name            TEXT NOT NULL,
    description     TEXT,
    data_type       TEXT,
    is_array        BOOLEAN NOT NULL DEFAULT 0,
    is_readonly     BOOLEAN NOT NULL DEFAULT 0,
    element_type    TEXT NOT NULL DEFAULT 'instance',
    default_value   TEXT,
    min_value       TEXT,
    max_value       TEXT
);

-- Methods
CREATE TABLE IF NOT EXISTS methods (
    id              INTEGER PRIMARY KEY,
    class_id        INTEGER NOT NULL REFERENCES classes(id),
    name            TEXT NOT NULL,
    description     TEXT,
    return_type     TEXT,
    return_is_array BOOLEAN NOT NULL DEFAULT 0,
    element_type    TEXT NOT NULL DEFAULT 'instance'
);

-- Parameters
CREATE TABLE IF NOT EXISTS parameters (
    id              INTEGER PRIMARY KEY,
    method_id       INTEGER NOT NULL REFERENCES methods(id),
    name            TEXT NOT NULL,
    description     TEXT,
    data_type       TEXT,
    is_array        BOOLEAN NOT NULL DEFAULT 0,
    is_optional     BOOLEAN NOT NULL DEFAULT 0,
    default_value   TEXT,
    sort_order      INTEGER NOT NULL
);

-- Full-text search
CREATE VIRTUAL TABLE IF NOT EXISTS dom_search USING fts5(
    entity_type,
    entity_name,
    parent_name,
    description
);

-- Indices
CREATE INDEX IF NOT EXISTS idx_properties_class ON properties(class_id);
CREATE INDEX IF NOT EXISTS idx_methods_class ON methods(class_id);
CREATE INDEX IF NOT EXISTS idx_parameters_method ON parameters(method_id);
CREATE INDEX IF NOT EXISTS idx_classes_suite ON classes(suite_id);
CREATE INDEX IF NOT EXISTS idx_classes_superclass ON classes(superclass_name);
"""


def build_database(data: dict, db_path: str, xml_path: str) -> dict:
    """Build SQLite database from parsed data.

    Returns validation stats.
    """
    # Remove existing DB
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    # Create schema
    conn.executescript(DB_SCHEMA)

    # Insert metadata
    build_ts = datetime.now().isoformat(timespec="seconds")
    meta = {
        "source_file": os.path.basename(xml_path),
        "dom_version": data["version"],
        "dom_title": data["title"],
        "build_timestamp": build_ts,
        "parser_version": PARSER_VERSION,
    }

    # Counts will be added after insert
    for key, value in meta.items():
        conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", (key, value))

    # Insert suites
    suite_ids = {}
    for suite_name in sorted(data["suites"].keys()):
        cur = conn.execute("INSERT INTO suites (name) VALUES (?)", (suite_name,))
        suite_ids[suite_name] = cur.lastrowid

    # Insert classes, properties, methods, parameters
    total_props = 0
    total_meths = 0
    total_params = 0
    fts_rows = []

    for cls in data["classes"]:
        suite_id = suite_ids.get(cls["suite"])
        cur = conn.execute(
            """INSERT INTO classes (name, suite_id, is_enum, is_dynamic, description, superclass_name)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                cls["name"],
                suite_id,
                int(cls["is_enum"]),
                int(cls["is_dynamic"]),
                cls["description"],
                cls["superclass_name"],
            ),
        )
        class_id = cur.lastrowid

        # FTS entry for class
        entity_type = "enum" if cls["is_enum"] else "class"
        fts_rows.append((entity_type, cls["name"], "", cls["description"]))

        # Insert properties
        for prop in cls["properties"]:
            conn.execute(
                """INSERT INTO properties
                   (class_id, name, description, data_type, is_array, is_readonly,
                    element_type, default_value, min_value, max_value)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    class_id,
                    prop["name"],
                    prop["description"],
                    prop["data_type"],
                    int(prop["is_array"]),
                    int(prop["is_readonly"]),
                    prop["element_type"],
                    prop["default_value"],
                    prop["min_value"],
                    prop["max_value"],
                ),
            )
            total_props += 1
            fts_rows.append(("property", prop["name"], cls["name"], prop["description"]))

        # Insert methods
        for meth in cls["methods"]:
            cur = conn.execute(
                """INSERT INTO methods
                   (class_id, name, description, return_type, return_is_array, element_type)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (
                    class_id,
                    meth["name"],
                    meth["description"],
                    meth["return_type"],
                    int(meth["return_is_array"]),
                    meth["element_type"],
                ),
            )
            method_id = cur.lastrowid
            total_meths += 1
            fts_rows.append(("method", meth["name"], cls["name"], meth["description"]))

            # Insert parameters
            for param in meth["parameters"]:
                conn.execute(
                    """INSERT INTO parameters
                       (method_id, name, description, data_type, is_array,
                        is_optional, default_value, sort_order)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        method_id,
                        param["name"],
                        param["description"],
                        param["data_type"],
                        int(param["is_array"]),
                        int(param["is_optional"]),
                        param["default_value"],
                        param["sort_order"],
                    ),
                )
                total_params += 1
                fts_rows.append(("parameter", param["name"], cls["name"], param["description"]))

    # Bulk insert FTS
    conn.executemany(
        "INSERT INTO dom_search (entity_type, entity_name, parent_name, description) VALUES (?, ?, ?, ?)",
        fts_rows,
    )

    # Update metadata with counts
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("class_count", str(len(data["classes"]))))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("property_count", str(total_props)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("method_count", str(total_meths)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("parameter_count", str(total_params)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("suite_count", str(len(data["suites"]))))

    conn.commit()
    conn.close()

    return {
        "class_count": len(data["classes"]),
        "property_count": total_props,
        "method_count": total_meths,
        "parameter_count": total_params,
        "suite_count": len(data["suites"]),
        "fts_rows": len(fts_rows),
        "db_path": db_path,
        "build_timestamp": build_ts,
    }


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate(data: dict, db_path: str) -> tuple[bool, list[str]]:
    """Validate database against parsed XML data.

    Returns (passed: bool, messages: list[str]).
    """
    messages = []
    passed = True

    if not os.path.exists(db_path):
        return False, ["Database file does not exist"]

    conn = sqlite3.connect(db_path)

    # Check counts
    checks = [
        ("classes", len(data["classes"])),
        ("suites", len(data["suites"])),
        ("properties", sum(len(c["properties"]) for c in data["classes"])),
        ("methods", sum(len(c["methods"]) for c in data["classes"])),
        ("parameters", sum(len(m["parameters"]) for c in data["classes"] for m in c["methods"])),
    ]

    for table, expected in checks:
        (actual,) = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
        if actual != expected:
            messages.append(f"FAIL: {table} count mismatch: DB={actual}, XML={expected}")
            passed = False
        else:
            messages.append(f"  OK: {table} count = {actual}")

    # Spot checks: verify a few specific classes exist
    spot_checks = ["Application", "Document", "TextFrame", "Story", "PageItem"]
    for class_name in spot_checks:
        row = conn.execute("SELECT id, description FROM classes WHERE name = ?", (class_name,)).fetchone()
        if row:
            messages.append(f"  OK: {class_name} found (id={row[0]})")
        else:
            # Class might not exist in all versions - warn but don't fail
            xml_names = {c["name"] for c in data["classes"]}
            if class_name in xml_names:
                messages.append(f"FAIL: {class_name} in XML but not in DB")
                passed = False
            else:
                messages.append(f"SKIP: {class_name} not in XML source")

    # Check FTS index
    (fts_count,) = conn.execute("SELECT COUNT(*) FROM dom_search").fetchone()
    expected_fts = (
        len(data["classes"])
        + sum(len(c["properties"]) for c in data["classes"])
        + sum(len(c["methods"]) for c in data["classes"])
        + sum(len(m["parameters"]) for c in data["classes"] for m in c["methods"])
    )
    if fts_count == expected_fts:
        messages.append(f"  OK: FTS index rows = {fts_count}")
    else:
        messages.append(f"FAIL: FTS index count mismatch: DB={fts_count}, expected={expected_fts}")
        passed = False

    # Check metadata
    meta_keys = ["source_file", "dom_version", "dom_title", "build_timestamp", "parser_version"]
    for key in meta_keys:
        row = conn.execute("SELECT value FROM db_meta WHERE key = ?", (key,)).fetchone()
        if row:
            messages.append(f"  OK: db_meta[{key}] = {row[0]}")
        else:
            messages.append(f"FAIL: db_meta[{key}] missing")
            passed = False

    conn.close()
    return passed, messages


def print_validation(passed: bool, messages: list[str]):
    """Print validation results."""
    sep = "=" * 55
    print(sep)
    print("  Database Validation Report")
    print(sep)
    for msg in messages:
        print(f"  {msg}")
    print(sep)
    status = "PASSED" if passed else "FAILED"
    try:
        symbol = "\u2713" if passed else "\u2717"
        print(f"  Structure validation: {symbol} {status}")
    except UnicodeEncodeError:
        symbol = "[OK]" if passed else "[FAIL]"
        print(f"  Structure validation: {symbol} {status}")
    print(sep)
