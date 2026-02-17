"""
InDesign DOM Parser - Parses OMV XML sources into SQLite database.
"""

import json
import os
import re
import sqlite3
import xml.etree.ElementTree as ET
from datetime import datetime

PARSER_VERSION = "2.0.0"


# ---------------------------------------------------------------------------
# XML Parsing
# ---------------------------------------------------------------------------

def parse_xml(xml_path: str, source_key: str = "dom") -> dict:
    """Parse one OMV XML file and return structured data."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    _strip_namespace(root)

    map_el = root.find("map")
    if map_el is None:
        raise ValueError("No <map> element found in XML")

    version = map_el.get("name", "")
    title = map_el.get("title", "")
    timestamp = map_el.get("time", "")

    suites = _parse_suites(map_el)

    class_to_suite = {}
    for suite_name, class_names in suites.items():
        for cn in class_names:
            class_to_suite[cn] = suite_name

    package = root.find("package")
    if package is None:
        raise ValueError("No <package> element found in XML")

    classes = []
    for classdef in package.findall("./classdef"):
        cls = _parse_classdef(classdef, source_key=source_key)
        cls["suite"] = class_to_suite.get(cls["name"], "")
        classes.append(cls)

    return {
        "source_key": source_key,
        "source_file": os.path.basename(xml_path),
        "version": version,
        "title": title,
        "timestamp": timestamp,
        "suites": suites,
        "classes": classes,
    }


def parse_sources(sources: list[tuple[str, str]]) -> list[dict]:
    """Parse multiple XML sources.

    sources: [(source_key, xml_path), ...]
    """
    parsed = []
    for source_key, xml_path in sources:
        parsed.append(parse_xml(xml_path, source_key=source_key))
    return parsed


def _strip_namespace(root: ET.Element) -> None:
    """Normalize namespaced XML tags to local names in-place."""
    for elem in root.iter():
        if "}" in elem.tag:
            elem.tag = elem.tag.split("}", 1)[1]


def _parse_suites(map_el) -> dict:
    """Parse suite navigation from map/topicref nodes."""
    suites = {}
    for suite_ref in map_el.findall("./topicref"):
        suite_name = suite_ref.get("navtitle", "")
        class_names = []
        for class_ref in suite_ref.findall("./topicref"):
            href = class_ref.get("href", "")
            if href.startswith("#/"):
                class_names.append(href[2:])
        if suite_name:
            suites[suite_name] = class_names
    return suites


def _parse_classdef(classdef, source_key: str) -> dict:
    """Parse one classdef."""
    name = classdef.get("name", "")
    is_enum = classdef.get("enumeration") == "true"
    is_dynamic = classdef.get("dynamic") == "true"

    short_description = _extract_text(classdef.find("shortdesc"))
    long_description = _extract_text(classdef.find("description"))
    description = _merge_descriptions(short_description, long_description)

    superclass_el = classdef.find("superclass")
    superclass_name = _extract_text(superclass_el) or None

    properties = []
    methods = []

    for elements in classdef.findall("./elements"):
        element_type = elements.get("type", "instance")

        for prop in elements.findall("./property"):
            properties.append(_parse_property(prop, element_type))

        for meth in elements.findall("./method"):
            methods.append(_parse_method(meth, element_type))

    return {
        "name": name,
        "source_key": source_key,
        "is_enum": is_enum,
        "is_dynamic": is_dynamic,
        "description": description,
        "description_long": long_description or None,
        "superclass_name": superclass_name,
        "properties": properties,
        "methods": methods,
    }


def _parse_property(prop, element_type: str) -> dict:
    """Parse one property."""
    name = prop.get("name", "")
    rwaccess = prop.get("rwaccess", "")
    is_readonly = rwaccess == "readonly"

    short_description = _extract_text(prop.find("shortdesc"))
    long_description = _extract_text(prop.find("description"))
    description = _merge_descriptions(short_description, long_description)
    data_type, data_type_ref, is_array, default_value, min_value, max_value = _parse_datatypes(prop)

    return {
        "name": name,
        "description": description,
        "data_type": data_type,
        "data_type_ref": data_type_ref,
        "is_array": is_array,
        "is_readonly": is_readonly,
        "element_type": element_type,
        "default_value": default_value,
        "min_value": min_value,
        "max_value": max_value,
    }


def _parse_method(meth, element_type: str) -> dict:
    """Parse one method."""
    name = meth.get("name", "")

    short_description = _extract_text(meth.find("shortdesc"))
    long_description = _extract_text(meth.find("description"))
    description = _merge_descriptions(short_description, long_description)

    return_type, return_type_ref, return_is_array, _, _, _ = _parse_datatypes(meth)

    parameters = []
    params_el = meth.find("parameters")
    if params_el is not None:
        for idx, param in enumerate(params_el.findall("./parameter")):
            parameters.append(_parse_parameter(param, idx))

    return {
        "name": name,
        "description": description,
        "return_type": return_type,
        "return_type_ref": return_type_ref,
        "return_is_array": return_is_array,
        "element_type": element_type,
        "parameters": parameters,
    }


def _parse_parameter(param, sort_order: int) -> dict:
    """Parse one method parameter."""
    name = param.get("name", "")
    is_optional = param.get("optional") == "true"

    short_description = _extract_text(param.find("shortdesc"))
    long_description = _extract_text(param.find("description"))
    description = _merge_descriptions(short_description, long_description)

    if not is_optional and description and "(Optional)" in description:
        is_optional = True

    data_type, data_type_ref, is_array, default_value, _, _ = _parse_datatypes(param)

    return {
        "name": name,
        "description": description,
        "data_type": data_type,
        "data_type_ref": data_type_ref,
        "is_array": is_array,
        "is_optional": is_optional,
        "default_value": default_value,
        "sort_order": sort_order,
    }


def _parse_datatypes(parent):
    """Parse one or multiple datatype child nodes from a parent node."""
    datatypes = parent.findall("./datatype")
    if not datatypes:
        return None, None, False, None, None, None

    type_parts = []
    ref_parts = []
    is_array_any = False
    default_value = None
    min_value = None
    max_value = None

    for datatype in datatypes:
        type_el = datatype.find("type")
        type_str = None
        type_href = None

        if type_el is not None:
            varies = type_el.get("varies")
            type_href = type_el.get("href")
            if varies:
                type_str = f"varies={varies}"
            else:
                type_str = _extract_text(type_el)
        if type_str:
            if datatype.find("array") is not None:
                type_str = f"{type_str}[]"
            type_parts.append(type_str)
        if type_href:
            ref_parts.append(_normalize_type_href(type_href))

        is_array_any = is_array_any or (datatype.find("array") is not None)

        if default_value is None:
            default_value = _extract_text(datatype.find("value")) or None
        if min_value is None:
            min_value = _extract_text(datatype.find("min")) or None
        if max_value is None:
            max_value = _extract_text(datatype.find("max")) or None

    unique_types = list(dict.fromkeys(type_parts))
    unique_refs = list(dict.fromkeys(ref_parts))
    data_type = "|".join(unique_types) if unique_types else None
    data_type_ref = "|".join(unique_refs) if unique_refs else None
    return data_type, data_type_ref, is_array_any, default_value, min_value, max_value


def _normalize_type_href(href: str) -> str:
    """Normalize href values for storage and lookup."""
    href = href.strip()
    if href.startswith("$COMMON/javascript.xml#/"):
        return f"javascript:{href.split('#/', 1)[1]}"
    if href.startswith("$COMMON/scriptui.xml#/"):
        return f"scriptui:{href.split('#/', 1)[1]}"
    if href.startswith("#/"):
        return f"local:{href[2:]}"
    return href


def _extract_text(elem) -> str:
    """Extract normalized plain text from an XML element."""
    if elem is None:
        return ""
    text = "".join(elem.itertext())
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _merge_descriptions(short_description: str, long_description: str) -> str:
    """Merge short and long description text into one field."""
    if short_description and long_description:
        if long_description.startswith(short_description):
            return long_description
        return f"{short_description}\n{long_description}"
    return short_description or long_description or ""


# ---------------------------------------------------------------------------
# Analysis Report
# ---------------------------------------------------------------------------

def analyze(data: dict, xml_path: str) -> dict:
    """Generate an analysis report from parsed data."""
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
        "source_key": data.get("source_key", "unknown"),
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
    print("  InDesign DOM Analysis Report")
    print(f"  Source key: {stats.get('source_key', 'unknown')}")
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

-- Source dimensions
CREATE TABLE IF NOT EXISTS sources (
    id          INTEGER PRIMARY KEY,
    key         TEXT NOT NULL UNIQUE,
    label       TEXT NOT NULL,
    file        TEXT
);

-- Suites
CREATE TABLE IF NOT EXISTS suites (
    id          INTEGER PRIMARY KEY,
    source_id   INTEGER NOT NULL REFERENCES sources(id),
    name        TEXT NOT NULL,
    UNIQUE(name, source_id)
);

-- Classes
CREATE TABLE IF NOT EXISTS classes (
    id              INTEGER PRIMARY KEY,
    name            TEXT NOT NULL,
    source_id       INTEGER NOT NULL REFERENCES sources(id),
    suite_id        INTEGER REFERENCES suites(id),
    is_enum         BOOLEAN NOT NULL DEFAULT 0,
    is_dynamic      BOOLEAN NOT NULL DEFAULT 0,
    description     TEXT,
    description_long TEXT,
    superclass_name TEXT,
    UNIQUE(name, source_id)
);

-- Properties
CREATE TABLE IF NOT EXISTS properties (
    id              INTEGER PRIMARY KEY,
    class_id        INTEGER NOT NULL REFERENCES classes(id),
    name            TEXT NOT NULL,
    description     TEXT,
    data_type       TEXT,
    data_type_ref   TEXT,
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
    return_type_ref TEXT,
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
    data_type_ref   TEXT,
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
    description,
    source
);

-- Indices
CREATE INDEX IF NOT EXISTS idx_properties_class ON properties(class_id);
CREATE INDEX IF NOT EXISTS idx_methods_class ON methods(class_id);
CREATE INDEX IF NOT EXISTS idx_parameters_method ON parameters(method_id);
CREATE INDEX IF NOT EXISTS idx_classes_suite ON classes(suite_id);
CREATE INDEX IF NOT EXISTS idx_classes_superclass ON classes(superclass_name);
CREATE INDEX IF NOT EXISTS idx_classes_source_name ON classes(source_id, name);
"""


def build_database(data: dict | list[dict], db_path: str, xml_path: str | None = None) -> dict:
    """Build SQLite database from one or many parsed source payloads."""
    sources_data = data if isinstance(data, list) else [data]

    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    # Create schema
    conn.executescript(DB_SCHEMA)

    build_ts = datetime.now().isoformat(timespec="seconds")
    source_files = [d.get("source_file", "") for d in sources_data]
    source_keys = [d.get("source_key", "") for d in sources_data]
    dom_payload = next((d for d in sources_data if d.get("source_key") == "dom"), sources_data[0])
    meta = {
        "source_file": os.path.basename(xml_path) if xml_path else dom_payload.get("source_file", ""),
        "source_files": json.dumps(source_files),
        "source_keys": ",".join(source_keys),
        "dom_version": dom_payload.get("version", ""),
        "dom_title": dom_payload.get("title", ""),
        "build_timestamp": build_ts,
        "parser_version": PARSER_VERSION,
    }

    for key, value in meta.items():
        conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", (key, value))

    source_ids = {}
    source_labels = {
        "dom": "InDesign DOM",
        "javascript": "Core JavaScript",
        "scriptui": "ScriptUI",
    }
    for payload in sources_data:
        source_key = payload["source_key"]
        cur = conn.execute(
            "INSERT INTO sources (key, label, file) VALUES (?, ?, ?)",
            (source_key, source_labels.get(source_key, source_key), payload.get("source_file")),
        )
        source_ids[source_key] = cur.lastrowid

    suite_ids = {}
    for payload in sources_data:
        source_id = source_ids[payload["source_key"]]
        for suite_name in sorted(payload["suites"].keys()):
            cur = conn.execute(
                "INSERT INTO suites (source_id, name) VALUES (?, ?)",
                (source_id, suite_name),
            )
            suite_ids[(payload["source_key"], suite_name)] = cur.lastrowid

    total_classes = 0
    total_props = 0
    total_meths = 0
    total_params = 0
    fts_rows = []

    for payload in sources_data:
        source_key = payload["source_key"]
        source_id = source_ids[source_key]
        for cls in payload["classes"]:
            suite_id = suite_ids.get((source_key, cls["suite"]))
            cur = conn.execute(
                """INSERT INTO classes
                   (name, source_id, suite_id, is_enum, is_dynamic, description, description_long, superclass_name)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    cls["name"],
                    source_id,
                    suite_id,
                    int(cls["is_enum"]),
                    int(cls["is_dynamic"]),
                    cls["description"],
                    cls["description_long"],
                    cls["superclass_name"],
                ),
            )
            class_id = cur.lastrowid
            total_classes += 1

            entity_type = "enum" if cls["is_enum"] else "class"
            fts_rows.append((entity_type, cls["name"], "", cls["description"], source_key))

            for prop in cls["properties"]:
                conn.execute(
                    """INSERT INTO properties
                       (class_id, name, description, data_type, data_type_ref, is_array,
                        is_readonly, element_type, default_value, min_value, max_value)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        class_id,
                        prop["name"],
                        prop["description"],
                        prop["data_type"],
                        prop["data_type_ref"],
                        int(prop["is_array"]),
                        int(prop["is_readonly"]),
                        prop["element_type"],
                        prop["default_value"],
                        prop["min_value"],
                        prop["max_value"],
                    ),
                )
                total_props += 1
                fts_rows.append(("property", prop["name"], cls["name"], prop["description"], source_key))

            for meth in cls["methods"]:
                cur = conn.execute(
                    """INSERT INTO methods
                       (class_id, name, description, return_type, return_type_ref, return_is_array, element_type)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (
                        class_id,
                        meth["name"],
                        meth["description"],
                        meth["return_type"],
                        meth["return_type_ref"],
                        int(meth["return_is_array"]),
                        meth["element_type"],
                    ),
                )
                method_id = cur.lastrowid
                total_meths += 1
                fts_rows.append(("method", meth["name"], cls["name"], meth["description"], source_key))

                for param in meth["parameters"]:
                    conn.execute(
                        """INSERT INTO parameters
                           (method_id, name, description, data_type, data_type_ref, is_array,
                            is_optional, default_value, sort_order)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            method_id,
                            param["name"],
                            param["description"],
                            param["data_type"],
                            param["data_type_ref"],
                            int(param["is_array"]),
                            int(param["is_optional"]),
                            param["default_value"],
                            param["sort_order"],
                        ),
                    )
                    total_params += 1
                    fts_rows.append(("parameter", param["name"], cls["name"], param["description"], source_key))

    conn.executemany(
        "INSERT INTO dom_search (entity_type, entity_name, parent_name, description, source) VALUES (?, ?, ?, ?, ?)",
        fts_rows,
    )

    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("source_count", str(len(sources_data))))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("class_count", str(total_classes)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("property_count", str(total_props)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("method_count", str(total_meths)))
    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("parameter_count", str(total_params)))
    conn.execute(
        "INSERT INTO db_meta (key, value) VALUES (?, ?)",
        ("suite_count", str(sum(len(payload["suites"]) for payload in sources_data))),
    )

    conn.commit()
    conn.close()

    return {
        "source_count": len(sources_data),
        "class_count": total_classes,
        "property_count": total_props,
        "method_count": total_meths,
        "parameter_count": total_params,
        "suite_count": sum(len(payload["suites"]) for payload in sources_data),
        "fts_rows": len(fts_rows),
        "db_path": db_path,
        "build_timestamp": build_ts,
    }


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate(data: dict | list[dict] | None, db_path: str, expect_sources: list[str] | None = None) -> tuple[bool, list[str]]:
    """Validate database against parsed source data."""
    messages = []
    passed = True

    if not os.path.exists(db_path):
        return False, ["Database file does not exist"]

    conn = sqlite3.connect(db_path)

    source_payloads = data if isinstance(data, list) else ([data] if isinstance(data, dict) else [])
    expected_classes = sum(len(payload["classes"]) for payload in source_payloads)
    expected_suites = sum(len(payload["suites"]) for payload in source_payloads)
    expected_props = sum(len(c["properties"]) for payload in source_payloads for c in payload["classes"])
    expected_methods = sum(len(c["methods"]) for payload in source_payloads for c in payload["classes"])
    expected_params = sum(len(m["parameters"]) for payload in source_payloads for c in payload["classes"] for m in c["methods"])

    checks = [
        ("classes", expected_classes),
        ("suites", expected_suites),
        ("properties", expected_props),
        ("methods", expected_methods),
        ("parameters", expected_params),
    ]

    if source_payloads:
        for table, expected in checks:
            (actual,) = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
            if actual != expected:
                messages.append(f"FAIL: {table} count mismatch: DB={actual}, XML={expected}")
                passed = False
            else:
                messages.append(f"  OK: {table} count = {actual}")

    if expect_sources:
        found_sources = {row[0] for row in conn.execute("SELECT key FROM sources").fetchall()}
        missing_sources = set(expect_sources) - found_sources
        if missing_sources:
            messages.append(f"FAIL: missing sources: {sorted(missing_sources)}")
            passed = False
        else:
            messages.append(f"  OK: sources present: {sorted(found_sources)}")

    (fts_count,) = conn.execute("SELECT COUNT(*) FROM dom_search").fetchone()
    if source_payloads:
        expected_fts = expected_classes + expected_props + expected_methods + expected_params
        if fts_count == expected_fts:
            messages.append(f"  OK: FTS index rows = {fts_count}")
        else:
            messages.append(f"FAIL: FTS index count mismatch: DB={fts_count}, expected={expected_fts}")
            passed = False

    meta_keys = ["source_file", "source_files", "source_keys", "dom_version", "dom_title", "build_timestamp", "parser_version"]
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
