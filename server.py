# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "mcp>=1.0.0",
# ]
# ///
"""
InDesign DOM MCP Server.

Provides tools for querying multi-source ExtendScript knowledge:
  1. lookup_class       – Full class info
  2. get_properties     – Properties with optional filter
  3. get_methods        – Methods with short signatures
  4. get_method_detail  – Single method with all parameters
  5. get_enum_values    – Enum constant values
  6. get_hierarchy      – Inheritance chain + subclasses
  7. search_dom         – Full-text search
  8. list_classes       – Class overview by suite/type
  9. dom_info           – DB metadata and statistics
"""

import json
import os
import sys
from pathlib import Path

# Add script directory to sys.path so db module can be imported
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

import db

# Resolve DB path
DB_PATH = os.environ.get(
    "EXTENDSCRIPT_DB",
    os.environ.get(
        "INDESIGN_DOM_DB",
        str(Path(__file__).parent / "extendscript.db"),
    ),
)

mcp = FastMCP(
    "InDesign DOM",
    instructions=(
        "This server provides access to three knowledge sources for InDesign scripting:\n"
        "1. InDesign DOM: classes, properties, methods, enums and inheritance.\n"
        "2. ExtendScript JavaScript Core: Adobe-specific JS extensions ($, UnitValue, "
        "File, Folder, Socket, XML/XMLList, RegExp, Reflection).\n"
        "3. ScriptUI: legacy dialog/UI framework (Window, Button, EditText, ListBox).\n\n"
        "Call knowledge_overview() first to see available capabilities.\n"
        "Use source='dom'|'javascript'|'scriptui' to disambiguate duplicate class names "
        "such as Window, Group, Panel, and Event.\n\n"
        "IMPORTANT: For ExtendScript code generation, check JavaScript source for "
        "UnitValue, $, File/Folder, Socket, and XML/XMLList before assuming "
        "standard JavaScript behavior.\n\n"
        "TIP: For common DOM patterns (navigation, page items, images, text, styles, "
        "geometry, find/change, export), the InDesign Exec MCP provides a "
        "get_quick_reference tool that returns a comprehensive cheatsheet. "
        "Check that first -- it may save you multiple lookups here."
    ),
)


def _fmt(obj) -> str:
    """Format result as indented JSON string."""
    return json.dumps(obj, indent=2, ensure_ascii=False)


def _append_scriptui_note(payload, source: str | None):
    """Append ScriptUI modernization note when ScriptUI data is requested."""
    note = (
        "ScriptUI is legacy technology. For new UI development, prefer a modern "
        "UXP plugin. ScriptUI remains useful for small dialogs and for maintaining "
        "existing scripts."
    )
    if source == "scriptui":
        if isinstance(payload, dict):
            payload = dict(payload)
            payload["note"] = note
            return payload
        if isinstance(payload, list):
            return [{"note": note, **item} if isinstance(item, dict) else item for item in payload]
    return payload


# ---------------------------------------------------------------------------
# Tool 1: lookup_class
# ---------------------------------------------------------------------------

@mcp.tool()
def lookup_class(name: str, source: str | None = None) -> str:
    """Look up full information for an InDesign DOM class.

    Returns suite, superclass, description, property/method counts,
    and direct subclasses.

    Args:
        name: The exact class name (e.g. "TextFrame", "Document", "Application")
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    result = db.lookup_class(name, source=source, db_path=DB_PATH)
    if not result:
        return f"Class '{name}' not found."
    return _fmt(_append_scriptui_note(result, source))


# ---------------------------------------------------------------------------
# Tool 2: get_properties
# ---------------------------------------------------------------------------

@mcp.tool()
def get_properties(
    class_name: str,
    source: str | None = None,
    filter: str | None = None,
    include_inherited: bool = False,
) -> str:
    """Get properties of an InDesign DOM class.

    Args:
        class_name: The class name (e.g. "TextFrame")
        source: Optional source filter ("dom", "javascript", "scriptui")
        filter: Optional substring filter on property name or description
        include_inherited: If true, includes properties from superclasses
    """
    results = db.get_properties(
        class_name,
        source=source,
        filter_text=filter,
        include_inherited=include_inherited,
        db_path=DB_PATH,
    )
    if not results:
        msg = f"No properties found for '{class_name}'"
        if filter:
            msg += f" matching '{filter}'"
        return msg + "."
    return _fmt(_append_scriptui_note(results, source))


# ---------------------------------------------------------------------------
# Tool 3: get_methods
# ---------------------------------------------------------------------------

@mcp.tool()
def get_methods(
    class_name: str,
    source: str | None = None,
    filter: str | None = None,
    include_inherited: bool = False,
) -> str:
    """Get methods of an InDesign DOM class with short signatures.

    Args:
        class_name: The class name (e.g. "Document")
        source: Optional source filter ("dom", "javascript", "scriptui")
        filter: Optional substring filter on method name or description
        include_inherited: If true, includes methods from superclasses
    """
    results = db.get_methods(
        class_name,
        source=source,
        filter_text=filter,
        include_inherited=include_inherited,
        db_path=DB_PATH,
    )
    if not results:
        msg = f"No methods found for '{class_name}'"
        if filter:
            msg += f" matching '{filter}'"
        return msg + "."
    return _fmt(_append_scriptui_note(results, source))


# ---------------------------------------------------------------------------
# Tool 4: get_method_detail
# ---------------------------------------------------------------------------

@mcp.tool()
def get_method_detail(class_name: str, method_name: str, source: str | None = None) -> str:
    """Get full detail for a single method including all parameters.

    Args:
        class_name: The class that owns the method (e.g. "Application")
        method_name: The method name (e.g. "findGrep")
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    result = db.get_method_detail(class_name, method_name, source=source, db_path=DB_PATH)
    if not result:
        return f"Method '{method_name}' not found on class '{class_name}'."
    return _fmt(_append_scriptui_note(result, source))


# ---------------------------------------------------------------------------
# Tool 5: get_enum_values
# ---------------------------------------------------------------------------

@mcp.tool()
def get_enum_values(enum_name: str, source: str | None = None) -> str:
    """Get all values of an InDesign DOM enumeration.

    Args:
        enum_name: The enum class name (e.g. "Justification")
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    result = db.get_enum_values(enum_name, source=source, db_path=DB_PATH)
    if not result:
        return f"Enum '{enum_name}' not found."
    return _fmt(_append_scriptui_note(result, source))


# ---------------------------------------------------------------------------
# Tool 6: get_hierarchy
# ---------------------------------------------------------------------------

@mcp.tool()
def get_hierarchy(class_name: str, source: str | None = None) -> str:
    """Get the full inheritance chain and direct subclasses of a class.

    Args:
        class_name: The class name (e.g. "TextFrame")
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    result = db.get_hierarchy(class_name, source=source, db_path=DB_PATH)
    if not result:
        return f"Class '{class_name}' not found."
    return _fmt(_append_scriptui_note(result, source))


# ---------------------------------------------------------------------------
# Tool 7: search_dom
# ---------------------------------------------------------------------------

@mcp.tool()
def search_dom(query: str, source: str | None = None) -> str:
    """Full-text search across all InDesign DOM entities.

    Searches class names, property names, method names, parameter names,
    and their descriptions. Returns up to 20 results.

    Args:
        query: Search terms (e.g. "find grep change", "hyperlink", "export pdf")
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    results = db.search_dom(query, source=source, max_results=20, db_path=DB_PATH)
    if not results:
        return f"No results found for '{query}'."
    return _fmt(_append_scriptui_note(results, source))


# ---------------------------------------------------------------------------
# Tool 8: list_classes
# ---------------------------------------------------------------------------

@mcp.tool()
def list_classes(
    suite: str | None = None,
    type: str = "all",
    source: str | None = None,
) -> str:
    """List InDesign DOM classes, optionally filtered by suite or type.

    Args:
        suite: Filter by suite name (e.g. "Text Suite", "Color Suite")
        type: Filter by type: "class", "enum", or "all" (default)
        source: Optional source filter ("dom", "javascript", "scriptui")
    """
    results = db.list_classes(suite=suite, type_filter=type, source=source, db_path=DB_PATH)
    if not results:
        msg = "No classes found"
        if suite:
            msg += f" in suite '{suite}'"
        return msg + "."
    return _fmt(_append_scriptui_note(results, source))


# ---------------------------------------------------------------------------
# Tool 9: dom_info
# ---------------------------------------------------------------------------

@mcp.tool()
def dom_info() -> str:
    """Get InDesign DOM database metadata and statistics.

    Returns DOM version, source file, build timestamp, and entity counts.
    """
    result = db.dom_info(db_path=DB_PATH)
    return _fmt(result)


@mcp.tool()
def list_sources() -> str:
    """List loaded knowledge sources and counts."""
    return _fmt(db.list_sources(db_path=DB_PATH))


@mcp.tool()
def knowledge_overview() -> str:
    """Get a compact capability overview for this MCP server."""
    return _fmt(db.knowledge_overview(db_path=DB_PATH))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Run the MCP server via stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
