"""
InDesign DOM MCP – CLI Management Tool.

Commands:
  analyze  --xml <file>     Structure report from XML
  build    --xml <file>     Build new database from XML
  update   --xml <file>     Update database (diff + rebuild)
  validate [--xml <file>]   Validate database against XML
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
DEFAULT_DB = BASE_DIR / "indesign_dom.db"


def cmd_analyze(args):
    """Analyze XML structure and print report."""
    xml_path = args.xml
    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    print(f"Parsing {xml_path} ...")
    data = dom_parser.parse_xml(xml_path)
    stats = dom_parser.analyze(data, xml_path)
    dom_parser.print_report(stats)
    return 0


def cmd_build(args):
    """Build database from XML."""
    xml_path = args.xml
    db_path = args.db or str(DEFAULT_DB)

    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    print(f"Parsing {xml_path} ...")
    data = dom_parser.parse_xml(xml_path)

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


def cmd_update(args):
    """Update database from new XML (diff + rebuild)."""
    xml_path = args.xml
    db_path = args.db or str(DEFAULT_DB)

    if not os.path.exists(xml_path):
        print(f"Error: XML file not found: {xml_path}")
        return 1

    # Parse new XML
    print(f"Parsing {xml_path} ...")
    new_data = dom_parser.parse_xml(xml_path)

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
    db_path = args.db or str(DEFAULT_DB)
    xml_path = args.xml

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        return 1

    if xml_path:
        if not os.path.exists(xml_path):
            print(f"Error: XML file not found: {xml_path}")
            return 1
        print(f"Parsing {xml_path} for validation ...")
        data = dom_parser.parse_xml(xml_path)
    else:
        # Validate DB structure only (no XML comparison)
        print("Validating database structure (no XML comparison) ...")
        conn = sqlite3.connect(db_path)
        tables = [
            r[0]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        ]
        conn.close()

        expected = {"db_meta", "suites", "classes", "properties", "methods", "parameters", "dom_search"}
        missing = expected - set(tables)
        if missing:
            print(f"FAIL: Missing tables: {missing}")
            return 1

        info = dom_db.dom_info(db_path=db_path)
        print(f"  Version:    {info['dom_version']}")
        print(f"  Title:      {info['dom_title']}")
        print(f"  Source:     {info['source_file']}")
        print(f"  Built:      {info['build_timestamp']}")
        print(f"  Classes:    {info['counts']['classes']}")
        print(f"  Properties: {info['counts']['properties']}")
        print(f"  Methods:    {info['counts']['methods']}")
        print("  Structure validation: [OK] PASSED")
        return 0

    passed, messages = dom_parser.validate(data, db_path)
    dom_parser.print_validation(passed, messages)
    return 0 if passed else 1


def cmd_serve(args):
    """Start the MCP server."""
    db_path = args.db or str(DEFAULT_DB)

    if not os.path.exists(db_path):
        print(f"Error: Database not found: {db_path}")
        print("Run 'manage.py build --xml <file>' first.")
        return 1

    # Set DB path as environment variable for server
    os.environ["INDESIGN_DOM_DB"] = db_path

    print(f"Starting InDesign DOM MCP Server ...")
    print(f"  Database: {db_path}")

    import server
    server.main()
    return 0


def cmd_info(args):
    """Show database statistics."""
    db_path = args.db or str(DEFAULT_DB)

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


def main():
    parser = argparse.ArgumentParser(
        description="InDesign DOM MCP – Management CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  analyze  --xml <file>     Analyze XML structure
  build    --xml <file>     Build new database
  update   --xml <file>     Update database (diff + rebuild)
  validate [--xml <file>]   Validate database
  serve                     Start MCP server
  info                      Show database statistics
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # analyze
    p_analyze = subparsers.add_parser("analyze", help="Analyze XML structure")
    p_analyze.add_argument("--xml", required=True, help="Path to OMV-XML file")

    # build
    p_build = subparsers.add_parser("build", help="Build new database from XML")
    p_build.add_argument("--xml", required=True, help="Path to OMV-XML file")
    p_build.add_argument("--db", help=f"Database path (default: {DEFAULT_DB})")

    # update
    p_update = subparsers.add_parser("update", help="Update database from new XML")
    p_update.add_argument("--xml", required=True, help="Path to new OMV-XML file")
    p_update.add_argument("--db", help=f"Database path (default: {DEFAULT_DB})")

    # validate
    p_validate = subparsers.add_parser("validate", help="Validate database")
    p_validate.add_argument("--xml", help="Path to OMV-XML file (optional)")
    p_validate.add_argument("--db", help=f"Database path (default: {DEFAULT_DB})")

    # serve
    p_serve = subparsers.add_parser("serve", help="Start MCP server")
    p_serve.add_argument("--db", help=f"Database path (default: {DEFAULT_DB})")

    # info
    p_info = subparsers.add_parser("info", help="Show database statistics")
    p_info.add_argument("--db", help=f"Database path (default: {DEFAULT_DB})")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    commands = {
        "analyze": cmd_analyze,
        "build": cmd_build,
        "update": cmd_update,
        "validate": cmd_validate,
        "serve": cmd_serve,
        "info": cmd_info,
    }

    return commands[args.command](args)


if __name__ == "__main__":
    sys.exit(main() or 0)
