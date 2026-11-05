"""
InDesign DOM MCP – CLI Management Tool.

Data source: https://www.indesignjs.de/indesignapi/ (api.json per target, by Gregor Fellenz).

Commands:
  check                     Compare upstream manifest.json (contentHash per target) with the database
  fetch                     Download manifest.json and api.json for dom/javascript/scriptui into sources/indesignjs/
  analyze                   Structure report from the downloaded api.json files
  build                     Build database from the downloaded api.json files
  update                    check + fetch + diff against current database + build
  validate                  Validate database (optionally against the api.json files)
  serve                     Start MCP server
  info                      Show database statistics
  review-submissions        Promote pending learnings to gotchas.json
"""

import argparse
import json
import os
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

import parser as dom_parser
import db as dom_db

BASE_DIR = Path(__file__).parent
DEFAULT_DB = BASE_DIR / "extendscript.db"
LEGACY_DB = BASE_DIR / "indesign_dom.db"
SOURCES_DIR = BASE_DIR / "sources" / "indesignjs"
GOTCHAS_PATH = BASE_DIR / "gotchas.json"
SUBMISSIONS_PATH = BASE_DIR / "submissions" / "pending.jsonl"
EXPECTED_SOURCES = list(dom_parser.DEFAULT_TARGETS)


def _default_db_path() -> Path:
    """Return preferred db path."""
    return DEFAULT_DB


def _target_dbs(args) -> list[str]:
    """Databases to write: --db, or the default plus the legacy file if it exists."""
    if getattr(args, "db", None):
        return [args.db]
    targets = [str(DEFAULT_DB)]
    if LEGACY_DB.exists():
        targets.append(str(LEGACY_DB))
    return targets


def _sources_dir(args) -> Path:
    return Path(args.dir) if getattr(args, "dir", None) else SOURCES_DIR


def _parse_local(sources_dir: Path) -> list[dict] | None:
    paths = dom_parser.local_source_paths(str(sources_dir))
    missing = [p for _, p in paths if not os.path.exists(p)]
    if missing:
        print(f"Error: api.json not found: {', '.join(missing)}")
        print("Run 'manage.py fetch' first.")
        return None
    print(f"Parsing {sources_dir} ...")
    return dom_parser.parse_sources(paths, manifest=dom_parser.load_manifest(str(sources_dir)))


def _remote_status(db_path: str) -> tuple[bool, list[str]]:
    """Return (up_to_date, report lines) by comparing the upstream manifest contentHash per source."""
    remote = dom_parser.fetch_manifest()["targets"]
    local = {}
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
            columns = {r[1] for r in conn.execute("PRAGMA table_info(sources)")}
            if "content_hash" in columns:
                local = {
                    key: (title, content_hash)
                    for key, title, content_hash in conn.execute("SELECT key, title, content_hash FROM sources")
                }
            conn.close()
        except sqlite3.Error:
            local = {}

    up_to_date = True
    lines = []
    for source_key, target in dom_parser.DEFAULT_TARGETS.items():
        entry = remote.get(target, {})
        title, content_hash = local.get(source_key, ("(none)", ""))
        same = bool(content_hash) and content_hash == entry.get("contentHash")
        up_to_date = up_to_date and same
        lines.append(f"  {source_key:<11} upstream: {entry.get('version', '(missing in manifest)')}")
        lines.append(f"  {'':<11} database: {title or '(none)'} - {'unchanged' if same else 'changed'}")
    return up_to_date, lines


def cmd_check(args):
    """Compare upstream versions with the database without downloading api.json."""
    up_to_date, lines = _remote_status(str(_target_dbs(args)[0]))
    print("\n".join(lines))
    print("  Status:    " + ("up to date" if up_to_date else "update available (run 'manage.py update')"))
    return 0


def cmd_fetch(args):
    """Download api.json for every source."""
    dest = _sources_dir(args)
    print(f"Downloading api.json files into {dest} ...")
    for source_key, path in dom_parser.fetch_sources(str(dest)):
        size = os.path.getsize(path) / 1e6
        print(f"  {source_key:<11} {path} ({size:.1f} MB)")
    return 0


def cmd_analyze(args):
    """Analyze downloaded api.json files and print report."""
    parsed = _parse_local(_sources_dir(args))
    if parsed is None:
        return 1
    for data in parsed:
        if args.source and data["source_key"] != args.source:
            continue
        dom_parser.print_report(dom_parser.analyze(data))
    return 0


def _build(parsed: list[dict], db_paths: list[str]) -> int:
    for data in parsed:
        dom_parser.print_report(dom_parser.analyze(data))

    passed_all = True
    for db_path in db_paths:
        print(f"\nBuilding database at {db_path} ...")
        try:
            stats = dom_parser.build_database(parsed, db_path)
        except PermissionError as exc:
            print(f"Error: cannot replace {db_path} ({exc}). Stop the MCP server using it and retry.")
            passed_all = False
            continue
        print(f"  Sources:    {stats['source_count']}")
        print(f"  Classes:    {stats['class_count']}")
        print(f"  Properties: {stats['property_count']}")
        print(f"  Methods:    {stats['method_count']}")
        print(f"  Parameters: {stats['parameter_count']}")
        print(f"  Events:     {stats['event_count']}")
        print(f"  FTS rows:   {stats['fts_rows']}")
        print(f"  Built at:   {stats['build_timestamp']}")

        passed, messages = dom_parser.validate(parsed, db_path, expect_sources=EXPECTED_SOURCES)
        dom_parser.print_validation(passed, messages)
        passed = _run_regression_checks(db_path) and passed
        passed_all = passed_all and passed
    return 0 if passed_all else 1


def cmd_build(args):
    """Build database from downloaded api.json files."""
    parsed = _parse_local(_sources_dir(args))
    if parsed is None:
        return 1
    return _build(parsed, _target_dbs(args))


def cmd_update(args):
    """Check upstream, fetch, show diff against the current database and rebuild."""
    db_paths = _target_dbs(args)
    if not args.force:
        up_to_date, lines = _remote_status(db_paths[0])
        print("\n".join(lines))
        if up_to_date:
            print("Database is up to date. Use --force to rebuild anyway.")
            return 0

    rc = cmd_fetch(args)
    if rc:
        return rc
    parsed = _parse_local(_sources_dir(args))
    if parsed is None:
        return 1
    if os.path.exists(db_paths[0]):
        _print_diff(db_paths[0], next(p for p in parsed if p["source_key"] == "dom"))
    return _build(parsed, db_paths)


def _db_members(conn: sqlite3.Connection, source_key: str) -> dict[str, set[str]]:
    members: dict[str, set[str]] = {}
    for (name,) in conn.execute(
        "SELECT c.name FROM classes c JOIN sources s ON s.id = c.source_id WHERE s.key = ?", (source_key,)
    ):
        members[name] = set()
    tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
    queries = [("properties", "p.name"), ("methods", "p.name || '()'")]
    if "events" in tables:
        queries.append(("events", "'event:' || p.name"))
    for table, expr in queries:
        for cls, member in conn.execute(
            f"""SELECT c.name, {expr} FROM {table} p
                JOIN classes c ON c.id = p.class_id JOIN sources s ON s.id = c.source_id
                WHERE s.key = ?""",
            (source_key,),
        ):
            members.setdefault(cls, set()).add(member)
    return members


def _print_diff(db_path: str, new_data: dict, limit: int = 25):
    """Print class- and member-level diff between existing DB and the new DOM payload."""
    sep = "=" * 60
    try:
        conn = sqlite3.connect(db_path)
        old_info = dict(conn.execute("SELECT key, value FROM db_meta").fetchall())
        old = _db_members(conn, new_data["source_key"])
        conn.close()
    except Exception as exc:
        print(f"Could not read existing database for diff: {exc}")
        return

    new = {
        c["name"]: {p["name"] for p in c["properties"]}
        | {f"{m['name']}()" for m in c["methods"]}
        | {f"event:{e['name']}" for e in c["events"]}
        for c in new_data["classes"]
    }
    common = set(new) & set(old)
    sections = (
        ("New classes", sorted(set(new) - set(old))),
        ("Removed classes", sorted(set(old) - set(new))),
        ("New members", sorted(f"{c}.{m}" for c in common for m in new[c] - old[c])),
        ("Removed members", sorted(f"{c}.{m}" for c in common for m in old[c] - new[c])),
    )

    print(sep)
    print(f"  DOM Update: {old_info.get('dom_title') or old_info.get('dom_version', '?')}")
    print(f"           -> {new_data['title']}")
    print(sep)
    for label, items in sections:
        print(f"  {label + ':':<17} {len(items):>5}")
        for item in items[:limit]:
            print(f"      {item}")
        if len(items) > limit:
            print(f"      ... (+{len(items) - limit})")
    print(sep)


def cmd_validate(args):
    """Validate database, optionally against the downloaded api.json files."""
    db_path = args.db or str(_default_db_path())
    expect_sources = args.expect_sources.split(",") if args.expect_sources else EXPECTED_SOURCES

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        return 1

    parsed = None
    if args.dir:
        parsed = _parse_local(Path(args.dir))
        if parsed is None:
            return 1

    passed, messages = dom_parser.validate(parsed, db_path, expect_sources=expect_sources)
    dom_parser.print_validation(passed, messages)
    passed = _run_regression_checks(db_path) and passed
    return 0 if passed else 1


def cmd_serve(args):
    """Start the MCP server."""
    db_path = args.db or str(_default_db_path())

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        print("Run 'manage.py update' first.")
        return 1

    os.environ["EXTENDSCRIPT_DB"] = db_path
    os.environ["INDESIGN_DOM_DB"] = db_path

    print("Starting InDesign DOM MCP Server ...")
    print(f"  Database: {db_path}")

    import server
    server.main()
    return 0


def cmd_info(args):
    """Show database statistics."""
    db_path = args.db or str(_default_db_path())

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        return 1

    info = dom_db.dom_info(db_path=db_path)

    sep = "=" * 60
    print(sep)
    print("  InDesign DOM Database Info")
    print(sep)
    print(f"  Origin:         {info['source_origin']}")
    print(f"  DOM:            {info['dom_title']} (generated {info['dom_generated']})")
    for src in info["sources"]:
        print(f"  {src['source']:<14}  {src['version']} ({src['generated']})")
    print(f"  Built:          {info['build_timestamp']}")
    print(f"  Parser version: {info['parser_version']}")
    print(sep)
    counts = info["counts"]
    print(f"  Classes (total):  {counts['classes']:>6}")
    print(f"    Objects:        {counts['objects']:>6}")
    print(f"    Collections:    {counts['collections']:>6}")
    print(f"    Enumerations:   {counts['enums']:>6}")
    print(f"  Properties:       {counts['properties']:>6}")
    print(f"  Methods:          {counts['methods']:>6}")
    print(f"  Parameters:       {counts['parameters']:>6}")
    print(f"  Events:           {counts['events']:>6}")
    print(sep)
    return 0


def _run_regression_checks(db_path: str) -> bool:
    """Run sample queries for quick regression coverage."""
    results = []

    def check(label: str, ok: bool):
        results.append(ok)
        print(f"    {'OK' if ok else 'FAIL'}: {label}")

    print("  Regression checks:")
    for class_name, source in (
        ("TextFrame", "dom"),
        ("Index", "dom"),
        ("UnitValue", "javascript"),
        ("$", "javascript"),
        ("File", "javascript"),
        ("RegExp", "javascript"),
        ("ScriptUI", "scriptui"),
    ):
        check(f"lookup_class('{class_name}', source='{source}')",
              bool(dom_db.lookup_class(class_name, source=source, db_path=db_path)))

    collision = dom_db.lookup_class("Window", db_path=db_path)
    check("lookup_class('Window') resolves dom + scriptui", isinstance(collision, list) and len(collision) >= 2)

    open_dialog = dom_db.get_methods("File", source="javascript", filter_text="openDialog", db_path=db_path)
    check("File.openDialog is static", any(m["element_type"] == "class" for m in open_dialog))

    def has_index_accessor(runtime):
        methods = dom_db.get_methods("TextFrames", source="dom", runtime=runtime, db_path=db_path)
        return any(m["name"] == "[]" for m in methods)

    check("TextFrames has '[]' (extendscript), not under runtime='uxp'",
          has_index_accessor(None) and not has_index_accessor("uxp"))

    justification = dom_db.get_enum_values("Justification", source="dom", db_path=db_path) or {}
    check("Justification has values and uxp_note", bool(justification.get("values")) and "uxp_note" in justification)

    full_name = dom_db.get_properties("Document", source="dom", filter_text="fullName", db_path=db_path)
    check("Document.fullName has uxp_type Promise<File>", any(p.get("uxp_type") == "Promise<File>" for p in full_name))

    item_by_name = dom_db.get_methods("XMLElements", source="dom", filter_text="itemByName", db_path=db_path)
    check("XMLElements.itemByName is extendscript_only", any(m.get("extendscript_only") for m in item_by_name))

    do_script = dom_db.get_method_detail("Application", "doScript", source="dom", db_path=db_path) or {}
    check("Application.doScript has uxp_note and extendscript_note",
          "uxp_note" in do_script and "extendscript_note" in do_script)

    index_class = dom_db.lookup_class("Index", source="dom", db_path=db_path) or {}
    check("Index doc_url points to Index_.md", str(index_class.get("doc_url", "")).endswith("/Index_.md"))

    try:
        dollar = dom_db.search_dom("$", db_path=db_path)
        check("search_dom('$') finds the $ class", any(r["entity_name"] == "$" for r in dollar))
    except sqlite3.Error:
        check("search_dom('$') finds the $ class", False)
    return all(results)

def _slugify(text: str) -> str:
    """Build a stable slug for gotcha IDs."""
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "learning"


def _next_unique_id(base: str, existing: set[str]) -> str:
    """Return unique ID by suffixing with numbers when needed."""
    if base not in existing:
        return base
    idx = 2
    while f"{base}-{idx}" in existing:
        idx += 1
    return f"{base}-{idx}"


def _load_gotchas_file() -> dict:
    """Load gotchas file or return default structure."""
    if not GOTCHAS_PATH.exists():
        return {"version": 1, "entries": []}
    with GOTCHAS_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        return {"version": 1, "entries": []}
    if "entries" not in data or not isinstance(data["entries"], list):
        data["entries"] = []
    if "version" not in data:
        data["version"] = 1
    return data


def _safe_write_text(path: Path, content: str):
    """Write text robustly on Windows sync folders."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        path.write_text(content, encoding="utf-8")
        return
    except OSError:
        # Fallback path for environments where Path.write_text intermittently fails.
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(content)


def _print_submission(idx: int, item: dict):
    """Print pending learning submission."""
    print("-" * 72)
    print(f"Submission #{idx}")
    print(f"  Category: {item.get('category', '(missing)')}")
    print(f"  Severity: {item.get('severity', '(missing)')}")
    print(f"  Triggers: {item.get('triggers', [])}")
    print(f"  Problem:  {item.get('problem', '(missing)')}")
    print(f"  Solution: {item.get('solution', '(missing)')}")
    if item.get("error_message"):
        print(f"  Error:    {item.get('error_message')}")
    if item.get("jsx_context"):
        preview = str(item.get("jsx_context")).replace("\n", "\\n")
        print(f"  JSX:      {preview[:180]}")


def cmd_review_submissions(args):
    """Review pending learning submissions and promote approved ones."""
    if not SUBMISSIONS_PATH.exists():
        print(f"No submission file found: {SUBMISSIONS_PATH}")
        return 0

    raw_lines = SUBMISSIONS_PATH.read_text(encoding="utf-8").splitlines()
    lines = [ln for ln in raw_lines if ln.strip()]
    if not lines:
        print("No pending submissions.")
        return 0

    gotchas = _load_gotchas_file()
    entries = gotchas.get("entries", [])
    existing_ids = {str(e.get("id", "")).strip() for e in entries if isinstance(e, dict)}

    approved = 0
    rejected = 0
    kept_lines: list[str] = []
    pending_tail: list[str] = []
    quit_early = False

    parsed_items: list[tuple[str, dict | None]] = []
    for ln in lines:
        try:
            parsed_items.append((ln, json.loads(ln)))
        except Exception:
            parsed_items.append((ln, None))

    for idx, (raw, item) in enumerate(parsed_items, start=1):
        if quit_early:
            pending_tail.append(raw)
            continue
        if item is None or not isinstance(item, dict):
            print(f"Skipping invalid JSON line #{idx}; keeping it in pending queue.")
            kept_lines.append(raw)
            continue

        _print_submission(idx, item)
        choice = input("Action [a=approve, s=skip, r=reject, q=quit] (default: s): ").strip().lower() or "s"
        if choice == "q":
            quit_early = True
            pending_tail = [raw]
            continue
        if choice == "r":
            rejected += 1
            continue
        if choice != "a":
            kept_lines.append(raw)
            continue

        problem = str(item.get("problem", "")).strip()
        solution = str(item.get("solution", "")).strip()
        triggers = item.get("triggers", [])
        if not problem or not solution or not isinstance(triggers, list) or not triggers:
            print("  Cannot approve: missing required fields (problem/solution/triggers). Keeping pending.")
            kept_lines.append(raw)
            continue

        base_id = _slugify(problem)[:64]
        entry_id = _next_unique_id(base_id, existing_ids)
        existing_ids.add(entry_id)
        approved_entry = {
            "id": entry_id,
            "category": str(item.get("category", "extendscript")),
            "severity": str(item.get("severity", "warning")),
            "triggers": [str(t).strip() for t in triggers if str(t).strip()],
            "problem": problem,
            "solution": solution,
            "added": date.today().isoformat(),
            "source": "auto-submission",
        }
        if item.get("jsx_context"):
            approved_entry["example_bad"] = str(item["jsx_context"])

        entries.append(approved_entry)
        approved += 1

    if pending_tail:
        kept_lines.extend(pending_tail)

    gotchas["entries"] = entries
    _safe_write_text(GOTCHAS_PATH, json.dumps(gotchas, indent=2, ensure_ascii=False) + "\n")
    _safe_write_text(SUBMISSIONS_PATH, ("\n".join(kept_lines) + "\n") if kept_lines else "")

    print("-" * 72)
    print(f"Review complete. Approved: {approved}, Rejected: {rejected}, Still pending: {len(kept_lines)}")
    print(f"Updated gotchas file: {GOTCHAS_PATH}")
    print(f"Pending queue file:   {SUBMISSIONS_PATH}")
    return 0



def main():
    parser = argparse.ArgumentParser(
        description="InDesign DOM MCP – Management CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  check                     Compare upstream versions with the database
  fetch                     Download api.json files into sources/indesignjs/
  analyze                   Analyze the downloaded api.json files
  build                     Build database from the downloaded api.json files
  update                    check + fetch + diff + build
  validate                  Validate database
  serve                     Start MCP server
  info                      Show database statistics
  review-submissions        Promote pending learnings to gotchas.json
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to run")
    dir_help = f"Folder with <target>/api.json (default: {SOURCES_DIR})"
    db_help = f"Database path (default: {DEFAULT_DB.name}, plus {LEGACY_DB.name} if present)"

    p_check = subparsers.add_parser("check", help="Compare upstream versions with the database")
    p_check.add_argument("--db", help=db_help)

    p_fetch = subparsers.add_parser("fetch", help="Download api.json files")
    p_fetch.add_argument("--dir", help=dir_help)

    p_analyze = subparsers.add_parser("analyze", help="Analyze downloaded api.json files")
    p_analyze.add_argument("--dir", help=dir_help)
    p_analyze.add_argument("--source", choices=EXPECTED_SOURCES, help="Only this source")

    p_build = subparsers.add_parser("build", help="Build database from downloaded api.json files")
    p_build.add_argument("--dir", help=dir_help)
    p_build.add_argument("--db", help=db_help)

    p_update = subparsers.add_parser("update", help="Check, fetch, diff and rebuild")
    p_update.add_argument("--dir", help=dir_help)
    p_update.add_argument("--db", help=db_help)
    p_update.add_argument("--force", action="store_true", help="Rebuild even if the database is up to date")

    p_validate = subparsers.add_parser("validate", help="Validate database")
    p_validate.add_argument("--dir", help="Compare counts against the api.json files in this folder")
    p_validate.add_argument("--expect-sources", help="Comma-separated expected sources (default: dom,javascript,scriptui)")
    p_validate.add_argument("--db", help=f"Database path (default: {DEFAULT_DB.name})")

    p_serve = subparsers.add_parser("serve", help="Start MCP server")
    p_serve.add_argument("--db", help=f"Database path (default: {DEFAULT_DB.name})")

    p_info = subparsers.add_parser("info", help="Show database statistics")
    p_info.add_argument("--db", help=f"Database path (default: {DEFAULT_DB.name})")

    subparsers.add_parser(
        "review-submissions",
        help="Review local learning submissions and promote approved ones",
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    commands = {
        "check": cmd_check,
        "fetch": cmd_fetch,
        "analyze": cmd_analyze,
        "build": cmd_build,
        "update": cmd_update,
        "validate": cmd_validate,
        "serve": cmd_serve,
        "info": cmd_info,
        "review-submissions": cmd_review_submissions,
    }

    return commands[args.command](args)


if __name__ == "__main__":
    sys.exit(main() or 0)
