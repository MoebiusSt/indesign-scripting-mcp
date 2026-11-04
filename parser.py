"""
InDesign DOM Parser - Fetches and parses indesignjs.de api.json sources into SQLite.

Source: Gregor Fellenz' Adobe ExtendScript API reference (https://www.indesignjs.de/indesignapi/),
generated from Adobe's OMV exports. Each target ships an api.json with the whole object model.
"""

import json
import os
import re
import sqlite3
import urllib.request
from datetime import datetime
from pathlib import Path

PARSER_VERSION = "3.0.0"

API_BASE_URL = "https://www.indesignjs.de/indesignapi"
USER_AGENT = "indesign-scripting-mcp"

# Source key -> default indesignjs.de target slug.
DEFAULT_TARGETS = {
    "dom": "indesign",
    "javascript": "javascript",
    "scriptui": "scriptui",
}

SOURCE_LABELS = {
    "dom": "InDesign DOM",
    "javascript": "Core JavaScript (ExtendScript only)",
    "scriptui": "ScriptUI (ExtendScript only)",
}

KIND_MAP = {
    "Object": "object",
    "Collection": "collection",
    "Enumeration": "enum",
}

# api.json carries no static flag on methods. Core JavaScript and ScriptUI are frozen, so the
# class-level (static) methods are listed here.
STATIC_METHODS = {
    ("javascript", "$"): {"about", "bp", "colorPicker", "evalFile", "gc", "getenv", "setenv", "sleep",
                          "toString", "write", "writeln"},
    ("javascript", "Date"): {"UTC", "parse"},
    ("javascript", "File"): {"decode", "encode", "isEncodingAvailable", "openDialog", "saveDialog"},
    ("javascript", "Folder"): {"decode", "encode", "isEncodingAvailable", "selectDialog"},
    ("javascript", "Math"): {"*"},
    ("javascript", "Object"): {"isValid"},
    ("javascript", "String"): {"fromCharCode"},
    ("javascript", "XML"): {"defaultSettings", "setSettings", "settings"},
    ("scriptui", "ScriptUI"): {"getResourceText", "newFont", "newImage"},
    ("scriptui", "Window"): {"alert", "confirm", "find", "prompt"},
}

PRIMITIVE_REFS = {
    "number": "javascript:Number",
    "string": "javascript:String",
    "boolean": "javascript:Boolean",
}
IGNORED_TYPE_TOKENS = {"any", "void", "null", "undefined", "VariesType"}

UXP_NOTES_PATH = Path(__file__).parent / "uxp_notes.json"

UXP_NOTE_ENUM = (
    "In UXP, compare with equals() - the == and === operators do not work on DOM objects: "
    "myPath.pathType.equals(PathType.CLOSED_PATH), not myPath.pathType == PathType.CLOSED_PATH."
)
UXP_NOTE_INDEX_ACCESSOR = "Not available in UXP: collections have no [] index accessor. Use item(index) instead."
UXP_NOTE_FILE = (
    "In UXP, File/Folder are entries from require('uxp').storage.localFileSystem - "
    "obtaining them and their read and write methods return promises (async)."
)
UXP_NOTE_FILE_HANDLER = "In UXP the handler must be a Function; a File (script file) handler is ExtendScript only."
UXP_NOTE_SOURCE = (
    "ExtendScript only - UXP uses a different engine. In UXP use the standard JavaScript classes (MDN) "
    "and require('uxp').storage for files."
)


# ---------------------------------------------------------------------------
# Fetching
# ---------------------------------------------------------------------------

def target_url(target: str, filename: str = "api.json") -> str:
    """Return the indesignjs.de URL for one target file."""
    return f"{API_BASE_URL}/{target}/{filename}"


def doc_url(target: str, object_slug: str) -> str:
    """Return the Markdown twin URL of one object page."""
    return f"{API_BASE_URL}/{target}/{object_slug}.md"


def _http_get(url: str, timeout: int = 120) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def fetch_sources(dest_dir: str, targets: dict[str, str] | None = None) -> list[tuple[str, str]]:
    """Download api.json for every source into dest_dir/<target>/api.json.

    Returns [(source_key, local_path), ...].
    """
    targets = targets or DEFAULT_TARGETS
    result = []
    for source_key, target in targets.items():
        url = target_url(target)
        payload = _http_get(url)
        json.loads(payload)
        path = Path(dest_dir) / target / "api.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".json.tmp")
        tmp.write_bytes(payload)
        os.replace(tmp, path)
        result.append((source_key, str(path)))
    return result


def fetch_remote_versions() -> dict:
    """Read the top-level llms.txt (about 2 KB).

    Returns {"generated": "YYYY-MM-DD" | "", "targets": {target_slug: version line}}.
    """
    text = _http_get(f"{API_BASE_URL}/llms.txt", timeout=30).decode("utf-8")
    versions = {}
    for match in re.finditer(r"\]\(https?://[^)]+/indesignapi/([\w-]+)/llms\.txt\):\s*(.+)", text):
        versions[match.group(1)] = match.group(2).strip()
    generated = re.search(r"\bon (\d{4}-\d{2}-\d{2})\b", text)
    return {"generated": generated.group(1) if generated else "", "targets": versions}


def local_source_paths(sources_dir: str, targets: dict[str, str] | None = None) -> list[tuple[str, str]]:
    """Return [(source_key, api.json path), ...] for an already fetched sources directory."""
    targets = targets or DEFAULT_TARGETS
    return [(key, str(Path(sources_dir) / target / "api.json")) for key, target in targets.items()]


# ---------------------------------------------------------------------------
# api.json Parsing
# ---------------------------------------------------------------------------

def parse_api_json(json_path: str, source_key: str = "dom") -> dict:
    """Parse one api.json file and return structured data."""
    with open(json_path, encoding="utf-8") as f:
        raw = json.load(f)

    target = raw.get("slug") or DEFAULT_TARGETS.get(source_key, source_key)
    objects = raw.get("objects", [])

    renames = {o["name"]: _canonical_name(o["name"]) for o in objects if _canonical_name(o["name"]) != o["name"]}
    uxp_overrides = _load_uxp_overrides().get(source_key, {})

    classes = [
        _parse_object(obj, source_key, target, renames, uxp_overrides)
        for obj in objects
    ]

    version_label = raw.get("version", "")
    return {
        "source_key": source_key,
        "source_file": os.path.basename(json_path),
        "target": target,
        "url": target_url(target),
        "product": raw.get("product", ""),
        "version": _extract_version_number(version_label) or version_label,
        "title": version_label,
        "generated": raw.get("generated", ""),
        "classes": classes,
    }


def parse_sources(sources: list[tuple[str, str]]) -> list[dict]:
    """Parse multiple api.json sources.

    sources: [(source_key, json_path), ...]
    """
    parsed = [parse_api_json(path, source_key=key) for key, path in sources]
    _resolve_type_refs(parsed)
    return parsed


def _canonical_name(name: str) -> str:
    """Undo the trailing-underscore escaping applied to reserved names (e.g. Index_ -> Index)."""
    if len(name) > 1 and name.endswith("_") and not name.endswith("__"):
        return name[:-1]
    return name


def _extract_version_number(label: str) -> str:
    match = re.search(r"\((\d+(?:\.\d+)+)\)", label)
    return match.group(1) if match else ""


def _load_uxp_overrides() -> dict:
    if not UXP_NOTES_PATH.exists():
        return {}
    with UXP_NOTES_PATH.open(encoding="utf-8") as f:
        data = json.load(f)
    return data.get("sources", {})


def _rename_type(type_str: str | None, renames: dict[str, str]) -> str | None:
    if not type_str or not renames:
        return type_str
    return re.sub(r"[A-Za-z_$][\w$]*", lambda m: renames.get(m.group(0), m.group(0)), type_str)


def _is_array_type(type_str: str | None) -> bool:
    return bool(type_str) and type_str.rstrip().endswith("]")


def _parse_object(obj: dict, source_key: str, target: str, renames: dict, uxp_overrides: dict) -> dict:
    slug = obj["name"]
    name = renames.get(slug, slug)
    kind = KIND_MAP.get(obj.get("kind", ""), "object")
    is_enum = kind == "enum"
    class_overrides = uxp_overrides.get(name, {})

    class_uxp_note = class_overrides.get("class")
    if not class_uxp_note and is_enum and source_key == "dom":
        class_uxp_note = UXP_NOTE_ENUM

    properties = []
    for prop in obj.get("properties", []) or []:
        properties.append(_parse_property(prop, source_key, is_enum, renames, class_overrides))

    methods = []
    for meth in obj.get("methods", []) or []:
        methods.append(_parse_method(meth, name, source_key, renames, class_overrides))

    events = [
        {"name": ev.get("name", ""), "description": ev.get("description", "") or ""}
        for ev in obj.get("events", []) or []
    ]

    return {
        "name": name,
        "slug": slug,
        "source_key": source_key,
        "kind": kind,
        "is_enum": is_enum,
        "description": obj.get("description", "") or "",
        "superclass_name": renames.get(obj.get("extends"), obj.get("extends")) or None,
        "collection_of": renames.get(obj.get("elementType"), obj.get("elementType")) or None,
        "doc_url": doc_url(target, slug),
        "uxp_note": class_uxp_note,
        "properties": properties,
        "methods": methods,
        "events": events,
    }


def _parse_property(prop: dict, source_key: str, is_enum: bool, renames: dict, class_overrides: dict) -> dict:
    name = prop.get("name", "")
    data_type = _rename_type(prop.get("type"), renames)
    rng = prop.get("range") or {}
    element_type = "class" if (is_enum or prop.get("static")) else "instance"

    uxp_note = class_overrides.get("properties", {}).get(name)
    if not uxp_note and source_key == "dom" and data_type:
        if re.search(r"\bFile\b", data_type) and re.search(r"\bFunction\b", data_type):
            uxp_note = UXP_NOTE_FILE_HANDLER
        elif re.search(r"\b(File|Folder)\b", data_type) and not _is_array_type(data_type):
            uxp_note = UXP_NOTE_FILE

    return {
        "name": name,
        "description": prop.get("description", "") or "",
        "data_type": data_type,
        "data_type_ref": None,
        "is_array": _is_array_type(data_type),
        "is_readonly": prop.get("access") == "readonly",
        "element_type": element_type,
        "default_value": prop.get("value"),
        "min_value": rng.get("min"),
        "max_value": rng.get("max"),
        "measurement_unit": bool(prop.get("measurementUnit")),
        "uxp_note": uxp_note,
    }


def _method_element_type(class_name: str, method_name: str, source_key: str) -> str:
    if method_name == class_name and source_key != "dom":
        return "constructor"
    statics = STATIC_METHODS.get((source_key, class_name), set())
    if "*" in statics or method_name in statics:
        return "class"
    if source_key == "scriptui" and re.match(r"on[A-Z]", method_name):
        return "event"
    return "instance"


def _parse_method(meth: dict, class_name: str, source_key: str, renames: dict, class_overrides: dict) -> dict:
    name = meth.get("name", "")
    return_type = _rename_type(meth.get("returns"), renames)
    es_only = name == "[]" or name in class_overrides.get("es_only_methods", [])

    uxp_note = class_overrides.get("methods", {}).get(name)
    if not uxp_note and es_only:
        uxp_note = UXP_NOTE_INDEX_ACCESSOR

    parameters = []
    for idx, param in enumerate(meth.get("parameters", []) or []):
        p_type = _rename_type(param.get("type"), renames)
        rng = param.get("range") or {}
        p_desc = param.get("description", "") or ""
        if not uxp_note and source_key == "dom" and p_type and re.search(r"\bFile\b", p_type) and re.search(r"\bFunction\b", p_type):
            uxp_note = UXP_NOTE_FILE_HANDLER
        parameters.append(
            {
                "name": param.get("name", ""),
                "description": p_desc,
                "data_type": p_type,
                "data_type_ref": None,
                "is_array": _is_array_type(p_type),
                "is_optional": bool(param.get("optional")) or "(Optional)" in p_desc,
                "default_value": None,
                "min_value": rng.get("min"),
                "max_value": rng.get("max"),
                "measurement_unit": bool(param.get("measurementUnit")),
                "sort_order": idx,
            }
        )

    return {
        "name": name,
        "description": meth.get("description", "") or "",
        "return_type": return_type,
        "return_type_ref": None,
        "return_is_array": _is_array_type(return_type),
        "element_type": _method_element_type(class_name, name, source_key),
        "es_only": es_only,
        "uxp_note": uxp_note,
        "parameters": parameters,
    }


def _resolve_type_refs(parsed: list[dict]) -> None:
    """Fill *_type_ref fields: same source first, then javascript, then scriptui."""
    names_by_source = {p["source_key"]: {c["name"] for c in p["classes"]} for p in parsed}

    def refs_for(type_str: str | None, source_key: str) -> str | None:
        if not type_str:
            return None
        refs = []
        for token in re.findall(r"[A-Za-z_$][\w$]*", type_str):
            if token in IGNORED_TYPE_TOKENS:
                continue
            if token in PRIMITIVE_REFS:
                refs.append(PRIMITIVE_REFS[token])
            elif token in names_by_source.get(source_key, ()):
                refs.append(f"local:{token}")
            elif token in names_by_source.get("javascript", ()):
                refs.append(f"javascript:{token}")
            elif token in names_by_source.get("scriptui", ()):
                refs.append(f"scriptui:{token}")
        unique = list(dict.fromkeys(refs))
        return "|".join(unique) if unique else None

    for payload in parsed:
        key = payload["source_key"]
        for cls in payload["classes"]:
            for prop in cls["properties"]:
                prop["data_type_ref"] = refs_for(prop["data_type"], key)
            for meth in cls["methods"]:
                meth["return_type_ref"] = refs_for(meth["return_type"], key)
                for param in meth["parameters"]:
                    param["data_type_ref"] = refs_for(param["data_type"], key)


# ---------------------------------------------------------------------------
# Analysis Report
# ---------------------------------------------------------------------------

def analyze(data: dict) -> dict:
    """Generate an analysis report from parsed data."""
    classes = data["classes"]
    kinds = {k: sum(1 for c in classes if c["kind"] == k) for k in ("object", "collection", "enum")}

    class_sizes = sorted(
        ((c["name"], len(c["properties"]), len(c["methods"])) for c in classes if not c["is_enum"]),
        key=lambda x: x[1] + x[2],
        reverse=True,
    )

    return {
        "source_key": data["source_key"],
        "source_file": data["source_file"],
        "url": data["url"],
        "version": data["version"],
        "title": data["title"],
        "generated": data["generated"],
        "class_count": len(classes),
        "object_count": kinds["object"],
        "collection_count": kinds["collection"],
        "enum_count": kinds["enum"],
        "property_count": sum(len(c["properties"]) for c in classes),
        "method_count": sum(len(c["methods"]) for c in classes),
        "parameter_count": sum(len(m["parameters"]) for c in classes for m in c["methods"]),
        "event_count": sum(len(c["events"]) for c in classes),
        "superclass_count": sum(1 for c in classes if c["superclass_name"]),
        "uxp_note_count": sum(
            (1 if c["uxp_note"] else 0)
            + sum(1 for p in c["properties"] if p["uxp_note"])
            + sum(1 for m in c["methods"] if m["uxp_note"])
            for c in classes
        ),
        "top_classes": class_sizes[:5],
    }


def print_report(stats: dict):
    """Print a formatted analysis report."""
    sep = "=" * 55
    print(sep)
    print("  InDesign DOM Analysis Report")
    print(f"  Source key: {stats['source_key']}")
    print(f"  Source:     {stats['url']}")
    print(f"  Version:    {stats['version']} | {stats['title']}")
    print(f"  Generated:  {stats['generated']}")
    print(sep)
    print(f"  Classes (total):     {stats['class_count']:>5}")
    print(f"    Objects:           {stats['object_count']:>5}")
    print(f"    Collections:       {stats['collection_count']:>5}")
    print(f"    Enumerations:      {stats['enum_count']:>5}")
    print(f"  Properties:          {stats['property_count']:>5}")
    print(f"  Methods:             {stats['method_count']:>5}")
    print(f"  Parameters:          {stats['parameter_count']:>5}")
    print(f"  Events:              {stats['event_count']:>5}")
    print(f"  Superclass refs:     {stats['superclass_count']:>5}")
    print(f"  UXP notes:           {stats['uxp_note_count']:>5}")
    print(sep)
    print("  Top classes by size:")
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
    target      TEXT,
    url         TEXT,
    version     TEXT,
    title       TEXT,
    generated   TEXT,
    uxp_note    TEXT
);

-- Classes (objects, collections, enumerations)
CREATE TABLE IF NOT EXISTS classes (
    id              INTEGER PRIMARY KEY,
    name            TEXT NOT NULL,
    source_id       INTEGER NOT NULL REFERENCES sources(id),
    kind            TEXT NOT NULL,
    is_enum         BOOLEAN NOT NULL DEFAULT 0,
    description     TEXT,
    superclass_name TEXT,
    collection_of   TEXT,
    doc_url         TEXT,
    uxp_note        TEXT,
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
    max_value       TEXT,
    measurement_unit BOOLEAN NOT NULL DEFAULT 0,
    uxp_note        TEXT
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
    element_type    TEXT NOT NULL DEFAULT 'instance',
    es_only         BOOLEAN NOT NULL DEFAULT 0,
    uxp_note        TEXT
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
    min_value       TEXT,
    max_value       TEXT,
    measurement_unit BOOLEAN NOT NULL DEFAULT 0,
    sort_order      INTEGER NOT NULL
);

-- Events dispatched by a class
CREATE TABLE IF NOT EXISTS events (
    id              INTEGER PRIMARY KEY,
    class_id        INTEGER NOT NULL REFERENCES classes(id),
    name            TEXT NOT NULL,
    description     TEXT
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
CREATE INDEX IF NOT EXISTS idx_events_class ON events(class_id);
CREATE INDEX IF NOT EXISTS idx_classes_superclass ON classes(superclass_name);
CREATE INDEX IF NOT EXISTS idx_classes_source_name ON classes(source_id, name);
"""


def build_database(data: dict | list[dict], db_path: str) -> dict:
    """Build SQLite database from one or many parsed source payloads."""
    sources_data = data if isinstance(data, list) else [data]

    tmp_path = f"{db_path}.tmp"
    for path in (tmp_path, f"{tmp_path}-wal", f"{tmp_path}-shm"):
        if os.path.exists(path):
            os.remove(path)

    conn = sqlite3.connect(tmp_path)
    conn.executescript(DB_SCHEMA)

    build_ts = datetime.now().isoformat(timespec="seconds")
    dom_payload = next((d for d in sources_data if d["source_key"] == "dom"), sources_data[0])
    meta = {
        "source_origin": API_BASE_URL,
        "source_files": json.dumps([d["url"] for d in sources_data]),
        "source_keys": ",".join(d["source_key"] for d in sources_data),
        "dom_target": dom_payload["target"],
        "dom_version": dom_payload["version"],
        "dom_title": dom_payload["title"],
        "dom_generated": dom_payload["generated"],
        "build_timestamp": build_ts,
        "parser_version": PARSER_VERSION,
    }
    for key, value in meta.items():
        conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", (key, value))

    totals = {"classes": 0, "properties": 0, "methods": 0, "parameters": 0, "events": 0}
    fts_rows = []

    for payload in sources_data:
        source_key = payload["source_key"]
        cur = conn.execute(
            """INSERT INTO sources (key, label, target, url, version, title, generated, uxp_note)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                source_key,
                SOURCE_LABELS.get(source_key, source_key),
                payload["target"],
                payload["url"],
                payload["version"],
                payload["title"],
                payload["generated"],
                UXP_NOTE_SOURCE if source_key in ("javascript", "scriptui") else None,
            ),
        )
        source_id = cur.lastrowid

        for cls in payload["classes"]:
            cur = conn.execute(
                """INSERT INTO classes
                   (name, source_id, kind, is_enum, description, superclass_name, collection_of, doc_url, uxp_note)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    cls["name"], source_id, cls["kind"], int(cls["is_enum"]), cls["description"],
                    cls["superclass_name"], cls["collection_of"], cls["doc_url"], cls["uxp_note"],
                ),
            )
            class_id = cur.lastrowid
            totals["classes"] += 1
            entity_type = "enum" if cls["is_enum"] else "class"
            fts_rows.append((entity_type, cls["name"], "", cls["description"], source_key))

            for prop in cls["properties"]:
                conn.execute(
                    """INSERT INTO properties
                       (class_id, name, description, data_type, data_type_ref, is_array, is_readonly,
                        element_type, default_value, min_value, max_value, measurement_unit, uxp_note)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        class_id, prop["name"], prop["description"], prop["data_type"], prop["data_type_ref"],
                        int(prop["is_array"]), int(prop["is_readonly"]), prop["element_type"],
                        prop["default_value"], prop["min_value"], prop["max_value"],
                        int(prop["measurement_unit"]), prop["uxp_note"],
                    ),
                )
                totals["properties"] += 1
                fts_rows.append(("property", prop["name"], cls["name"], prop["description"], source_key))

            for meth in cls["methods"]:
                cur = conn.execute(
                    """INSERT INTO methods
                       (class_id, name, description, return_type, return_type_ref, return_is_array,
                        element_type, es_only, uxp_note)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        class_id, meth["name"], meth["description"], meth["return_type"], meth["return_type_ref"],
                        int(meth["return_is_array"]), meth["element_type"], int(meth["es_only"]), meth["uxp_note"],
                    ),
                )
                method_id = cur.lastrowid
                totals["methods"] += 1
                fts_rows.append(("method", meth["name"], cls["name"], meth["description"], source_key))

                for param in meth["parameters"]:
                    conn.execute(
                        """INSERT INTO parameters
                           (method_id, name, description, data_type, data_type_ref, is_array, is_optional,
                            default_value, min_value, max_value, measurement_unit, sort_order)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            method_id, param["name"], param["description"], param["data_type"],
                            param["data_type_ref"], int(param["is_array"]), int(param["is_optional"]),
                            param["default_value"], param["min_value"], param["max_value"],
                            int(param["measurement_unit"]), param["sort_order"],
                        ),
                    )
                    totals["parameters"] += 1
                    fts_rows.append(("parameter", param["name"], cls["name"], param["description"], source_key))

            for ev in cls["events"]:
                conn.execute(
                    "INSERT INTO events (class_id, name, description) VALUES (?, ?, ?)",
                    (class_id, ev["name"], ev["description"]),
                )
                totals["events"] += 1
                fts_rows.append(("event", ev["name"], cls["name"], ev["description"], source_key))

    conn.executemany(
        "INSERT INTO dom_search (entity_type, entity_name, parent_name, description, source) VALUES (?, ?, ?, ?, ?)",
        fts_rows,
    )

    conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", ("source_count", str(len(sources_data))))
    for key, value in totals.items():
        conn.execute("INSERT INTO db_meta (key, value) VALUES (?, ?)", (f"{key[:-1]}_count", str(value)))

    conn.commit()
    conn.close()

    for path in (f"{db_path}-wal", f"{db_path}-shm"):
        if os.path.exists(path):
            os.remove(path)
    os.replace(tmp_path, db_path)

    return {
        "source_count": len(sources_data),
        "class_count": totals["classes"],
        "property_count": totals["properties"],
        "method_count": totals["methods"],
        "parameter_count": totals["parameters"],
        "event_count": totals["events"],
        "fts_rows": len(fts_rows),
        "db_path": db_path,
        "build_timestamp": build_ts,
    }


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

REQUIRED_TABLES = {"db_meta", "sources", "classes", "properties", "methods", "parameters", "events", "dom_search"}
REQUIRED_META = ["source_origin", "source_files", "source_keys", "dom_version", "dom_title", "build_timestamp",
                 "parser_version"]


def validate(data: dict | list[dict] | None, db_path: str, expect_sources: list[str] | None = None) -> tuple[bool, list[str]]:
    """Validate database against parsed source data."""
    messages = []
    passed = True

    if not os.path.exists(db_path):
        return False, ["Database file does not exist"]

    conn = sqlite3.connect(db_path)
    tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    missing_tables = REQUIRED_TABLES - tables
    if missing_tables:
        conn.close()
        return False, [f"FAIL: missing tables: {sorted(missing_tables)} (rebuild with 'manage.py update')"]

    payloads = data if isinstance(data, list) else ([data] if isinstance(data, dict) else [])
    if payloads:
        expected = {
            "classes": sum(len(p["classes"]) for p in payloads),
            "properties": sum(len(c["properties"]) for p in payloads for c in p["classes"]),
            "methods": sum(len(c["methods"]) for p in payloads for c in p["classes"]),
            "parameters": sum(len(m["parameters"]) for p in payloads for c in p["classes"] for m in c["methods"]),
            "events": sum(len(c["events"]) for p in payloads for c in p["classes"]),
        }
        for table, count in expected.items():
            (actual,) = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
            if actual != count:
                messages.append(f"FAIL: {table} count mismatch: DB={actual}, source={count}")
                passed = False
            else:
                messages.append(f"  OK: {table} count = {actual}")

        (fts_count,) = conn.execute("SELECT COUNT(*) FROM dom_search").fetchone()
        expected_fts = sum(expected.values())
        if fts_count == expected_fts:
            messages.append(f"  OK: FTS index rows = {fts_count}")
        else:
            messages.append(f"FAIL: FTS index count mismatch: DB={fts_count}, expected={expected_fts}")
            passed = False

    if expect_sources:
        found = {row[0] for row in conn.execute("SELECT key FROM sources").fetchall()}
        missing = set(expect_sources) - found
        if missing:
            messages.append(f"FAIL: missing sources: {sorted(missing)}")
            passed = False
        else:
            messages.append(f"  OK: sources present: {sorted(found)}")

    for key in REQUIRED_META:
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
    print(f"  Structure validation: {'[OK] PASSED' if passed else '[FAIL] FAILED'}")
    print(sep)


# ---------------------------------------------------------------------------
# UXP coverage check (rendered HTML pages of indesignapi.zip)
# ---------------------------------------------------------------------------
# The build never depends on HTML. This check reads Gregor's UXP markers from the
# rendered pages and reports the ones that the rules + uxp_notes.json do not cover.

ARCHIVE_URL = f"{API_BASE_URL}/indesignapi.zip"

_MEMBER_ANCHOR_RE = re.compile(r'<(?:tr|article)[^>]*\bid="([pme])-([^"]+)"([^>]*)>')
_UXP_TEXT_RE = re.compile(r'<(p|span)[^>]*class="m?warn"[^>]*data-only="uxp"[^>]*>(.*?)</\1>', re.S)
_UXP_ASYNC_RE = re.compile(r'<span class="uxpv"[^>]*data-only="uxp"')


def download_archive(dest: str, url: str = ARCHIVE_URL) -> str:
    """Download indesignapi.zip (about 19 MB) to dest and return the path."""
    payload = _http_get(url, timeout=300)
    path = Path(dest)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".zip.tmp")
    tmp.write_bytes(payload)
    os.replace(tmp, path)
    return str(path)


def scan_uxp_markers(archive_path: str, target: str = "indesign") -> dict:
    """Collect UXP markers per object page.

    Returns {object_name: {"class_note": str | None,
                           "members": {(kind, name): {"hidden": bool, "note": str | None, "async": bool}}}}
    with kind "p" (property), "m" (method) or "e" (event). Object names are canonical (Index_ -> Index).
    """
    import html as html_lib
    import zipfile

    def clean(fragment: str) -> str:
        return re.sub(r"\s+", " ", html_lib.unescape(re.sub(r"<[^>]+>", " ", fragment))).strip()

    def notes_in(fragment: str) -> str | None:
        texts = list(dict.fromkeys(clean(m.group(2)) for m in _UXP_TEXT_RE.finditer(fragment)))
        return " ".join(t for t in texts if t) or None

    if Path(archive_path).is_dir():
        base = Path(archive_path) / target
        pages = [(p.stem, p.read_text(encoding="utf-8")) for p in base.glob("*.html")]
    else:
        with zipfile.ZipFile(archive_path) as zf:
            prefix = f"{target}/"
            pages = [
                (n[len(prefix):-5], zf.read(n).decode("utf-8"))
                for n in zf.namelist()
                if n.startswith(prefix) and n.endswith(".html") and "/" not in n[len(prefix):]
            ]

    result = {}
    for slug, page in pages:
        if slug == "index":
            continue
        anchors = list(_MEMBER_ANCHOR_RE.finditer(page))
        head = page[: anchors[0].start()] if anchors else page
        members = {}
        for idx, anchor in enumerate(anchors):
            end = anchors[idx + 1].start() if idx + 1 < len(anchors) else len(page)
            segment = page[anchor.start():end]
            entry = {
                "hidden": 'data-uxp="hide"' in anchor.group(3),
                "note": notes_in(segment),
                "async": bool(_UXP_ASYNC_RE.search(segment)),
            }
            if entry["hidden"] or entry["note"] or entry["async"]:
                members[(anchor.group(1), html_lib.unescape(anchor.group(2)))] = entry
        class_note = notes_in(head)
        if class_note or members:
            result[_canonical_name(slug)] = {"class_note": class_note, "members": members}
    return result
