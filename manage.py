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
"""

import argparse
import os
import sqlite3
import sys
from pathlib import Path

import parser as dom_parser
import db as dom_db

BASE_DIR = Path(__file__).parent
DEFAULT_DB = BASE_DIR / "extendscript.db"
LEGACY_DB = BASE_DIR / "indesign_dom.db"


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
    }

    return commands[args.command](args)


if __name__ == "__main__":
    sys.exit(main() or 0)
