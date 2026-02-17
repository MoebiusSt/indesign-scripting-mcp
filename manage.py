"""
InDesign DOM MCP – CLI Management Tool.

Commands:
  analyze   --xml <file>     Structure report from XML
  build     --xml <file>     Build new database from XML
  build-all --dom --js --sui Build multi-source database
  update    --xml <file>     Update database (diff + rebuild)
  validate  [--xml <file>]   Validate database against XML
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
GOTCHAS_PATH = BASE_DIR / "gotchas.json"
SUBMISSIONS_PATH = BASE_DIR / "submissions" / "pending.jsonl"


def _default_db_path() -> Path:
    """Return preferred db path."""
    return DEFAULT_DB


def cmd_analyze(args):
    """Analyze XML structure and print report."""
    xml_path = args.xml
    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    source_key = args.source
    print(f"Parsing {xml_path} (source={source_key}) ...")
    data = dom_parser.parse_xml(xml_path, source_key=source_key)
    stats = dom_parser.analyze(data, xml_path)
    dom_parser.print_report(stats)
    return 0


def cmd_build(args):
    """Build database from XML."""
    xml_path = args.xml
    db_path = args.db or str(_default_db_path())
    source_key = args.source

    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    print(f"Parsing {xml_path} (source={source_key}) ...")
    data = dom_parser.parse_xml(xml_path, source_key=source_key)

    # Analysis report
    stats = dom_parser.analyze(data, xml_path)
    dom_parser.print_report(stats)

    # Build DB
    print(f"\nBuilding database at {db_path} ...")
    build_stats = dom_parser.build_database(data, db_path, xml_path)
    print(f"  Classes:    {build_stats['class_count']}")
    print(f"  Properties: {build_stats['property_count']}")
    print(f"  Methods:    {build_stats['method_count']}")
    print(f"  Parameters: {build_stats['parameter_count']}")
    print(f"  FTS rows:   {build_stats['fts_rows']}")
    print(f"  Built at:   {build_stats['build_timestamp']}")

    # Validate
    print("\nValidating ...")
    passed, messages = dom_parser.validate(data, db_path)
    dom_parser.print_validation(passed, messages)

    if passed:
        print("\nDatabase built successfully.")
    else:
        print("\nDatabase built with validation errors.")

    return 0 if passed else 1


def cmd_build_all(args):
    """Build database from DOM + JavaScript + ScriptUI XML sources."""
    db_path = args.db or str(_default_db_path())
    xml_sources = [
        ("dom", args.dom),
        ("javascript", args.js),
        ("scriptui", args.sui),
    ]

    for source_key, xml_path in xml_sources:
        if not os.path.exists(xml_path):
            print(f"Error: XML file not found for source '{source_key}': {xml_path}")
            return 1

    parsed_sources = []
    for source_key, xml_path in xml_sources:
        print(f"Parsing {xml_path} (source={source_key}) ...")
        data = dom_parser.parse_xml(xml_path, source_key=source_key)
        parsed_sources.append(data)
        stats = dom_parser.analyze(data, xml_path)
        dom_parser.print_report(stats)

    print(f"\nBuilding multi-source database at {db_path} ...")
    build_stats = dom_parser.build_database(parsed_sources, db_path)
    print(f"  Sources:    {build_stats['source_count']}")
    print(f"  Classes:    {build_stats['class_count']}")
    print(f"  Properties: {build_stats['property_count']}")
    print(f"  Methods:    {build_stats['method_count']}")
    print(f"  Parameters: {build_stats['parameter_count']}")
    print(f"  FTS rows:   {build_stats['fts_rows']}")
    print(f"  Built at:   {build_stats['build_timestamp']}")

    print("\nValidating ...")
    passed, messages = dom_parser.validate(parsed_sources, db_path, expect_sources=["dom", "javascript", "scriptui"])
    dom_parser.print_validation(passed, messages)
    return 0 if passed else 1


def cmd_update(args):
    """Update database from new XML (diff + rebuild)."""
    xml_path = args.xml
    db_path = args.db or str(_default_db_path())
    source_key = args.source

    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    # Parse new XML
    print(f"Parsing {xml_path} (source={source_key}) ...")
    new_data = dom_parser.parse_xml(xml_path, source_key=source_key)

    # If DB exists, show diff
    if os.path.exists(db_path):
        _print_diff(db_path, new_data)
    else:
        print("No existing database found. Building fresh.")

    # Rebuild
    print(f"\nRebuilding database at {db_path} ...")
    build_stats = dom_parser.build_database(new_data, db_path, xml_path)

    print(f"  Classes:    {build_stats['class_count']}")
    print(f"  Properties: {build_stats['property_count']}")
    print(f"  Methods:    {build_stats['method_count']}")
    print(f"  Parameters: {build_stats['parameter_count']}")

    # Validate
    print("\nValidating ...")
    passed, messages = dom_parser.validate(new_data, db_path)
    dom_parser.print_validation(passed, messages)

    if passed:
        print("\nDatabase updated successfully.")
    else:
        print("\nDatabase updated with validation errors.")

    return 0 if passed else 1


def _print_diff(db_path: str, new_data: dict):
    """Print diff between existing DB and new XML data."""
    sep = "=" * 55

    try:
        old_info = dom_db.dom_info(db_path=db_path)
    except Exception:
        print("Could not read existing database for diff.")
        return

    old_version = old_info.get("dom_version", "?")
    new_version = new_data["version"]

    # Get old class names
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        old_classes = {
            r["name"]
            for r in conn.execute("SELECT name FROM classes").fetchall()
        }
        old_enums = {
            r["name"]
            for r in conn.execute("SELECT name FROM classes WHERE is_enum = 1").fetchall()
        }
        old_counts = old_info.get("counts", {})
        conn.close()
    except Exception:
        print("Could not read existing database for diff.")
        return

    new_classes = {c["name"] for c in new_data["classes"]}
    new_enums = {c["name"] for c in new_data["classes"] if c["is_enum"]}

    added_classes = new_classes - old_classes
    removed_classes = old_classes - new_classes
    added_enums = new_enums - old_enums

    new_prop_count = sum(len(c["properties"]) for c in new_data["classes"])
    new_meth_count = sum(len(c["methods"]) for c in new_data["classes"])

    old_prop_count = old_counts.get("properties", 0)
    old_meth_count = old_counts.get("methods", 0)

    prop_diff = new_prop_count - old_prop_count
    meth_diff = new_meth_count - old_meth_count

    print(sep)
    print(f"  DOM Update: {old_version} -> {new_version}")
    print(sep)

    if added_classes:
        names = ", ".join(sorted(added_classes)[:10])
        if len(added_classes) > 10:
            names += ", ..."
        print(f"  New classes:      +{len(added_classes):>3}   ({names})")
    else:
        print(f"  New classes:      +  0")

    if removed_classes:
        names = ", ".join(sorted(removed_classes)[:10])
        if len(removed_classes) > 10:
            names += ", ..."
        print(f"  Removed classes:  -{len(removed_classes):>3}   ({names})")
    else:
        print(f"  Removed classes:  -  0")

    # Modified classes: classes present in both but potentially different
    common = old_classes & new_classes
    print(f"  Common classes:   {len(common):>5}")

    if added_enums:
        print(f"  New enums:        +{len(added_enums):>3}")

    sign_p = "+" if prop_diff >= 0 else ""
    sign_m = "+" if meth_diff >= 0 else ""
    print(f"  Properties delta: {sign_p}{prop_diff}")
    print(f"  Methods delta:    {sign_m}{meth_diff}")
    print(sep)


def cmd_validate(args):
    """Validate database against XML."""
    db_path = args.db or str(_default_db_path())
    xml_path = args.xml
    expect_sources = args.expect_sources.split(",") if args.expect_sources else None

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        return 1

    if xml_path:
        if not os.path.exists(xml_path):
            print(f"Error: XML file not found: {xml_path}")
            return 1
        source_key = args.source
        print(f"Parsing {xml_path} for validation (source={source_key}) ...")
        data = dom_parser.parse_xml(xml_path, source_key=source_key)
    else:
        print("Validating database structure (no XML comparison) ...")
        conn = sqlite3.connect(db_path)
        tables = [
            r[0]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        ]
        conn.close()

        expected = {"db_meta", "sources", "suites", "classes", "properties", "methods", "parameters", "dom_search"}
        missing = expected - set(tables)
        if missing:
            print(f"FAIL: Missing tables: {missing}")
            return 1

        if expect_sources:
            conn = sqlite3.connect(db_path)
            found_sources = {r[0] for r in conn.execute("SELECT key FROM sources").fetchall()}
            conn.close()
            missing_sources = set(expect_sources) - found_sources
            if missing_sources:
                print(f"FAIL: Missing sources: {sorted(missing_sources)}")
                return 1

        info = dom_db.dom_info(db_path=db_path)
        print(f"  Version:    {info['dom_version']}")
        print(f"  Title:      {info['dom_title']}")
        print(f"  Source:     {info['source_file']}")
        print(f"  Built:      {info['build_timestamp']}")
        print(f"  Classes:    {info['counts']['classes']}")
        print(f"  Properties: {info['counts']['properties']}")
        print(f"  Methods:    {info['counts']['methods']}")

        _run_regression_checks(db_path)
        print("  Structure validation: [OK] PASSED")
        return 0

    passed, messages = dom_parser.validate(data, db_path, expect_sources=expect_sources)
    dom_parser.print_validation(passed, messages)
    _run_regression_checks(db_path)
    return 0 if passed else 1


def cmd_serve(args):
    """Start the MCP server."""
    db_path = args.db or str(_default_db_path())

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        print("Run 'manage.py build --xml <file>' first.")
        return 1

    # Set DB path as environment variable for server
    os.environ["EXTENDSCRIPT_DB"] = db_path
    os.environ["INDESIGN_DOM_DB"] = db_path

    print(f"Starting InDesign DOM MCP Server ...")
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

    sep = "=" * 55
    print(sep)
    print(f"  InDesign DOM Database Info")
    print(sep)
    print(f"  Version:        {info['dom_version']}")
    print(f"  Title:          {info['dom_title']}")
    print(f"  Source file:    {info['source_file']}")
    print(f"  Source files:   {info['source_files']}")
    print(f"  Built:          {info['build_timestamp']}")
    print(f"  Parser version: {info['parser_version']}")
    print(sep)
    counts = info["counts"]
    print(f"  Suites:           {counts['suites']:>6}")
    print(f"  Classes (total):  {counts['classes']:>6}")
    print(f"    Regular:        {counts['regular_classes']:>6}")
    print(f"    Enumerations:   {counts['enums']:>6}")
    print(f"  Properties:       {counts['properties']:>6}")
    print(f"  Methods:          {counts['methods']:>6}")
    print(f"  Parameters:       {counts['parameters']:>6}")
    print(sep)
    return 0


def _run_regression_checks(db_path: str):
    """Run sample queries for quick regression coverage."""
    checks = [
        ("UnitValue", "javascript"),
        ("$", "javascript"),
        ("File", "javascript"),
        ("RegExp", "javascript"),
        ("ScriptUI", "scriptui"),
    ]
    print("  Regression checks:")
    for class_name, source in checks:
        row = dom_db.lookup_class(class_name, source=source, db_path=db_path)
        status = "OK" if row else "FAIL"
        print(f"    {status}: lookup_class('{class_name}', source='{source}')")
    collision = dom_db.lookup_class("Window", db_path=db_path)
    if isinstance(collision, list) and len(collision) >= 2:
        print("    OK: lookup_class('Window') resolves multiple sources")
    else:
        print("    FAIL: lookup_class('Window') should return multiple sources")


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
  analyze   --xml <file>     Analyze XML structure
  build     --xml <file>     Build new database
  build-all --dom --js --sui Build multi-source database
  update    --xml <file>     Update database (diff + rebuild)
  validate  [--xml <file>]   Validate database
  serve                     Start MCP server
  info                      Show database statistics
  review-submissions        Promote pending learnings to gotchas.json
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # analyze
    p_analyze = subparsers.add_parser("analyze", help="Analyze XML structure")
    p_analyze.add_argument("--xml", required=True, help="Path to OMV-XML file")
    p_analyze.add_argument(
        "--source",
        default="dom",
        choices=["dom", "javascript", "scriptui"],
        help="Source key for this XML file",
    )

    # build
    p_build = subparsers.add_parser("build", help="Build new database from XML")
    p_build.add_argument("--xml", required=True, help="Path to OMV-XML file")
    p_build.add_argument(
        "--source",
        default="dom",
        choices=["dom", "javascript", "scriptui"],
        help="Source key for this XML file",
    )
    p_build.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # build-all
    p_build_all = subparsers.add_parser("build-all", help="Build database from DOM+JavaScript+ScriptUI XML")
    p_build_all.add_argument("--dom", required=True, help="Path to InDesign DOM OMV XML")
    p_build_all.add_argument("--js", required=True, help="Path to javascript.xml")
    p_build_all.add_argument("--sui", required=True, help="Path to scriptui.xml")
    p_build_all.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # update
    p_update = subparsers.add_parser("update", help="Update database from new XML")
    p_update.add_argument("--xml", required=True, help="Path to new OMV-XML file")
    p_update.add_argument(
        "--source",
        default="dom",
        choices=["dom", "javascript", "scriptui"],
        help="Source key for this XML file",
    )
    p_update.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # validate
    p_validate = subparsers.add_parser("validate", help="Validate database")
    p_validate.add_argument("--xml", help="Path to OMV-XML file (optional)")
    p_validate.add_argument(
        "--source",
        default="dom",
        choices=["dom", "javascript", "scriptui"],
        help="Source key for XML validation mode",
    )
    p_validate.add_argument("--expect-sources", help="Comma-separated expected sources (e.g. dom,javascript,scriptui)")
    p_validate.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # serve
    p_serve = subparsers.add_parser("serve", help="Start MCP server")
    p_serve.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # info
    p_info = subparsers.add_parser("info", help="Show database statistics")
    p_info.add_argument("--db", help=f"Database path (default: {_default_db_path()})")

    # review-submissions
    subparsers.add_parser(
        "review-submissions",
        help="Review local learning submissions and promote approved ones",
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    commands = {
        "analyze": cmd_analyze,
        "build": cmd_build,
        "build-all": cmd_build_all,
        "update": cmd_update,
        "validate": cmd_validate,
        "serve": cmd_serve,
        "info": cmd_info,
        "review-submissions": cmd_review_submissions,
    }

    return commands[args.command](args)


if __name__ == "__main__":
    sys.exit(main() or 0)
