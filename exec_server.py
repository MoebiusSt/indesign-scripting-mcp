# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "mcp>=1.0.0",
#     "pywin32>=306",
# ]
# ///
"""
InDesign Exec MCP Server.

Provides 5 tools for executing JSX code in Adobe InDesign via COM/OLE:
  1. run_jsx          - Execute JSX code with undo grouping
  2. get_document_info - Query active document overview
  3. get_selection     - Query current selection
  4. eval_expression   - Evaluate a short expression
  5. undo              - Undo last agent operation(s)

Requires InDesign Desktop running on Windows.
Uses UndoModes.ENTIRE_SCRIPT to group all operations per call.
No eval() anywhere — JSX code is inlined, results serialised via __safeStringify.
"""

import json
import sys
from pathlib import Path

# Add script directory to sys.path so indesign_com can be imported
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

import indesign_com as com

mcp = FastMCP(
    "InDesign Exec",
    instructions=(
        "This server executes ExtendScript (JSX) code directly in a running Adobe InDesign instance. "
        "Use run_jsx to perform operations in InDesign documents. Every run_jsx call with "
        "undo_mode='entire' (default) groups all changes under a single undo step, so "
        "Ctrl+Z reverts the entire operation. Use the undo tool to programmatically revert changes.\n\n"
        "IMPORTANT: In your JSX code, assign the return value to the variable __result. Example:\n"
        "  var doc = app.activeDocument;\n"
        "  __result = {name: doc.name, pages: doc.pages.length};\n\n"
        "Do NOT use return statements — the wrapper handles serialisation.\n"
        "Always verify results after modifications using eval_expression or get_document_info.\n"
        "Use the InDesign DOM MCP server to look up classes, properties and methods "
        "before writing JSX code.\n\n"
        "PERFORMANCE: InDesign has a powerful Collection API. ALWAYS prefer collection methods "
        "over manual for-loops:\n"
        "  - everyItem() for bulk operations: collection.everyItem().prop = value (ONE command for ALL items)\n"
        "  - itemByName('x'), itemByID(n), itemByRange(a,b) for direct access without loops\n"
        "  - Nested: doc.stories.everyItem().paragraphs.everyItem().appliedParagraphStyle (reads ALL in ONE call)\n"
        "  - Only use loops when each element needs a DIFFERENT value (use getElements() first)\n"
        "  Example — set all docs to page 26 WITHOUT a loop:\n"
        "    app.documents.everyItem().layoutWindows.everyItem().activePage = "
        "app.documents.everyItem().pages.itemByName('26');\n"
        "  Example — read all page names in one call:\n"
        "    __result = doc.pages.everyItem().name;  // returns Array"
    ),
)


# ---------------------------------------------------------------------------
# MCP Resource: Usage instructions for the agent
# (Inspired by zachshallbetter/indesign-mcp-server's help system)
# ---------------------------------------------------------------------------

@mcp.resource("config://usage")
def usage_instructions() -> str:
    """Usage guide for the InDesign Exec MCP — read this first."""
    return """\
# InDesign Exec MCP — Usage Guide

## Core Principle
Use the **indesign-dom** MCP to look up classes, properties, methods and enums,
then write JSX code and execute it via **run_jsx**.

## The __result Convention
Assign your return value to `__result`. The wrapper serialises it as JSON.

    var doc = app.activeDocument;
    __result = {name: doc.name, pages: doc.pages.length};

If you don't assign __result, the tool returns `{success: true, result: null}`.
Do NOT use `return` — the wrapper handles that.

## Undo Modes
- `undo_mode="entire"` (default) — groups ALL changes as one Ctrl+Z step.
  Always provide a descriptive `undo_name` like "Agent: Format headings".
- `undo_mode="none"` — for read-only queries. No undo step created.
- `undo_mode="auto"` — each DOM change gets its own undo step.

## Workflow Pattern
1. **Inspect**: `get_document_info` or `eval_expression` to understand the document
2. **Look up**: Use indesign-dom MCP to find the right classes/methods
3. **Execute**: `run_jsx` with undo_mode="entire" and a descriptive undo_name
4. **Verify**: `eval_expression` or `get_selection` to check the result
5. **Rollback**: `undo` if the result is wrong, then try a different approach

## Common Patterns

### Iterate pages
    var doc = app.activeDocument;
    var data = [];
    for (var i = 0; i < doc.pages.length; i++) {
        var pg = doc.pages[i];
        data.push({name: pg.name, items: pg.allPageItems.length});
    }
    __result = data;

### Find/Change text
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;
    app.findTextPreferences.findWhat = "old text";
    app.changeTextPreferences.changeTo = "new text";
    var found = app.activeDocument.changeText();
    __result = {replaced: found.length};

### Apply paragraph style
    var doc = app.activeDocument;
    var style = doc.paragraphStyles.itemByName("Heading 1");
    var story = doc.stories[0];
    story.paragraphs[0].appliedParagraphStyle = style;
    __result = {applied: style.name};

### Create a text frame
    var doc = app.activeDocument;
    var page = doc.pages[0];
    var tf = page.textFrames.add();
    tf.geometricBounds = [20, 20, 100, 180]; // [top, left, bottom, right] in document units
    tf.contents = "Hello from the Agent";
    __result = {id: tf.id, bounds: tf.geometricBounds};

## Collection Patterns (PREFER over loops)

InDesign collections support **collective specifiers** via `everyItem()` and direct
accessors like `itemByName()`, `itemByID()`, `itemByRange()`.  These are vastly more
efficient than loops because they send ONE command across the scripting bridge
instead of N separate round-trips.

**RULE: Always use collection methods when possible. Only fall back to loops when
each element needs a different value.**

### Decision Matrix

| Situation | Use | Example |
|---|---|---|
| Same property on ALL items | `everyItem().prop = x` | `doc.rectangles.everyItem().label = "tagged"` |
| Same method on ALL items | `everyItem().method()` | `doc.pages[0].textFrames.everyItem().move(undefined, [10, 10])` |
| Set multiple properties at once | `everyItem().properties = {...}` | See example below |
| Read ALL values (returns Array) | `everyItem().prop` | `__result = doc.rectangles.everyItem().label` |
| Single element by name | `itemByName("x")` | `doc.paragraphStyles.itemByName("Heading 1")` |
| Single element by ID | `itemByID(n)` | `doc.textFrames.itemByID(12345)` |
| Range of elements | `itemByRange(a, b)` | `doc.pages.itemByRange(0, 4)` |
| First / last / middle | `firstItem()` / `lastItem()` | `doc.stories.firstItem().contents` |
| Nested bulk access | chain `everyItem()` | `doc.stories.everyItem().paragraphs.everyItem().appliedParagraphStyle` |
| **Different value per element** | **Loop + getElements()** | See loop example below |
| **Structure changes during iteration** | **Loop backwards** | Deleting items shifts indices |
| **Text ranges after text edits** | **Re-resolve specifier** | Char-index ranges go stale |

### Efficient Examples (everyItem)

    // Bulk property write — ONE bridge crossing
    doc.stories.everyItem().tables.everyItem().properties = {
        topBorderStrokeColor: "Black",
        bottomBorderStrokeColor: "Black"
    };

    // Append text to every story's last insertion point
    doc.stories.everyItem().insertionPoints.lastItem().contents = "!";

    // Read all page names in one call (returns Array)
    __result = doc.pages.everyItem().name;

    // Navigate all layout windows to page 26
    app.documents.everyItem().layoutWindows.everyItem().activePage =
        app.documents.everyItem().pages.itemByName("26");

### Efficient Examples (Direct Accessors)

    // By name — no loop needed
    var style = doc.paragraphStyles.itemByName("Heading 1");
    var swatch = doc.swatches.itemByName("Red");

    // By ID — direct access
    var frame = doc.textFrames.itemByID(12345);

    // By range — subset without loop
    var firstFive = doc.pages.itemByRange(0, 4);

### When Loops ARE Required

    // Different colour per rectangle — loop is unavoidable
    var recs = doc.rectangles.everyItem().getElements();
    var colors = doc.swatches.everyItem().getElements();
    for (var i = 0; i < recs.length; i++) {
        recs[i].fillColor = colors[i % colors.length];
    }
    __result = {colored: recs.length};

    // Deleting items — loop BACKWARDS to avoid index shift
    var items = doc.rectangles.everyItem().getElements();
    for (var i = items.length - 1; i >= 0; i--) {
        if (items[i].label === "delete_me") items[i].remove();
    }

### Critical Gotchas

1. **Any property access resolves the specifier** into a snapshot.
   New items added after resolution won't be seen. Call `getElements()`
   to re-resolve if the collection changed.

2. **everyItem() on an empty collection** returns `isValid = true`
   but produces empty results — no error is thrown.

3. **Text specifiers** (Characters, Words, Paragraphs) use character-index
   ranges internally. After text insertion/deletion the ranges go stale.
   Dynamic specifiers (`firstItem()`, `lastItem()`) re-evaluate automatically;
   resolved ones do not.

4. **Nested everyItem()** works and is powerful:
   `doc.stories.everyItem().paragraphs.everyItem().appliedParagraphStyle`
   reads ALL paragraph styles of ALL stories in ONE call.

5. **getElements()** converts a collective specifier to a plain JS Array.
   Use it when you need to loop with individual values, or to refresh
   a stale specifier.

## Safety Notes
- DOM objects in __result are serialised as specifier strings, not expanded.
- Properties that crash InDesign (scriptPreferences.properties, shadow-settings
  in find/change prefs) are automatically skipped during serialisation.
- userInteractionLevel is set to NEVER_INTERACT during execution to prevent
  modal dialogs from blocking the script.
"""


def _fmt(obj) -> str:
    """Format result as indented JSON string."""
    return json.dumps(obj, indent=2, ensure_ascii=False)


def _check_connection() -> str | None:
    """Ensure InDesign is connected. Returns error string or None."""
    try:
        com.connect()
        return None
    except ConnectionError as e:
        return str(e)


def _check_document() -> str | None:
    """Ensure a document is open. Returns error string or None."""
    err = _check_connection()
    if err:
        return err
    try:
        result = com.eval_expr("app.documents.length")
        if result == "0":
            return "No document open in InDesign."
        return None
    except Exception as e:
        return f"Error checking documents: {e}"


def _unwrap_result(result: dict) -> dict:
    """Flatten {success: true, result: {…}} into {success: true, …}.

    When __result is a dict, its keys are merged into the top-level response
    for a cleaner tool output.  Other result types are left as-is.
    """
    if result.get("success") and isinstance(result.get("result"), dict):
        data = result["result"]
        return {"success": True, **data}
    return result


# ---------------------------------------------------------------------------
# Tool 1: run_jsx
# ---------------------------------------------------------------------------

@mcp.tool()
def run_jsx(
    code: str,
    undo_name: str = "Agent Script",
    undo_mode: str = "entire",
) -> str:
    """Execute JSX (ExtendScript) code in InDesign.

    The code runs inside a safety wrapper that catches errors and returns
    structured results. When undo_mode is 'entire' (default), all changes
    are grouped under a single undo step labeled with undo_name.

    IMPORTANT: Assign the value you want to return to the variable __result.
    Example:
        var doc = app.activeDocument;
        __result = {name: doc.name, pages: doc.pages.length};

    PERFORMANCE — prefer InDesign Collection methods over loops:
        // GOOD: bulk operation via everyItem() — one command for all items
        doc.textFrames.everyItem().label = "processed";
        __result = doc.pages.everyItem().name;  // reads all names at once
        // GOOD: direct access without loop
        var style = doc.paragraphStyles.itemByName("Heading 1");
        // BAD: manual loop when everyItem() would work
        for (var i = 0; i < doc.textFrames.length; i++) { doc.textFrames[i].label = "processed"; }
    Only use loops when each element needs a DIFFERENT value.

    Args:
        code: The JSX code to execute. Assign to __result to return data.
        undo_name: Human-readable label for Edit > Undo (e.g. "Agent: Format headings")
        undo_mode: "entire" groups all changes as one undo step (default),
                   "auto" lets InDesign handle undo per-operation,
                   "none" skips undo tracking (for read-only operations)
    """
    err = _check_connection()
    if err:
        return _fmt({"success": False, "error": err})

    try:
        result = com.run_jsx(code, undo_name=undo_name, undo_mode=undo_mode)
        return _fmt(result)
    except Exception as e:
        return _fmt({"success": False, "error": str(e), "name": type(e).__name__})


# ---------------------------------------------------------------------------
# Tool 2: get_document_info
# ---------------------------------------------------------------------------

# Fixed JSX for safe document info retrieval.
# Assigns a plain JS object to __result — the wrapper serialises it.
_DOC_INFO_JSX = """\
var doc = app.activeDocument;
var sel = app.selection;
var selTypes = [];
for (var i = 0; i < sel.length && i < 20; i++) {
    selTypes.push(sel[i].constructor.name);
}

__result = {
    name: doc.name,
    fullName: doc.fullName ? doc.fullName.fsName : "(unsaved)",
    saved: doc.saved,
    modified: doc.modified,
    pages: doc.pages.length,
    spreads: doc.spreads.length,
    stories: doc.stories.length,
    allPageItems: doc.allPageItems.length,
    textFrames: doc.textFrames.length,
    rectangles: doc.rectangles.length,
    ovals: doc.ovals.length,
    graphicLines: doc.graphicLines.length,
    images: doc.allGraphics.length,
    links: doc.links.length,
    layers: doc.layers.length,
    masterSpreads: doc.masterSpreads.length,
    paragraphStyles: doc.allParagraphStyles.length,
    characterStyles: doc.allCharacterStyles.length,
    objectStyles: doc.allObjectStyles.length,
    swatches: doc.swatches.length,
    selection_count: sel.length,
    selection_types: selTypes,
    documentPreferences: {
        pageWidth: doc.documentPreferences.pageWidth,
        pageHeight: doc.documentPreferences.pageHeight,
        facingPages: doc.documentPreferences.facingPages,
        pagesPerDocument: doc.documentPreferences.pagesPerDocument
    }
};
"""


@mcp.tool()
def get_document_info() -> str:
    """Get an overview of the active InDesign document.

    Returns document name, page count, item counts, selection info,
    style counts, and document preferences. Read-only operation.
    """
    err = _check_document()
    if err:
        return _fmt({"success": False, "error": err})

    try:
        result = com.run_jsx(_DOC_INFO_JSX, undo_mode="none")
        return _fmt(_unwrap_result(result))
    except Exception as e:
        return _fmt({"success": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Tool 3: get_selection
# ---------------------------------------------------------------------------

_SELECTION_BASIC_JSX = """\
var sel = app.selection;
var items = [];
for (var i = 0; i < sel.length && i < 50; i++) {
    var obj = sel[i];
    var item = {
        index: i,
        type: obj.constructor.name,
        id: obj.id || -1
    };
    try { item.name = obj.name || ""; } catch(e) { item.name = ""; }
    try {
        if (obj.geometricBounds) {
            var b = obj.geometricBounds;
            item.bounds = {top: b[0], left: b[1], bottom: b[2], right: b[3]};
        }
    } catch(e) {}
    try {
        if (obj.contents && typeof obj.contents === 'string') {
            item.content_preview = obj.contents.substring(0, 200);
        }
    } catch(e) {}
    items.push(item);
}
__result = {count: sel.length, items: items};
"""

_SELECTION_FULL_JSX = """\
var sel = app.selection;
var items = [];
for (var i = 0; i < sel.length && i < 50; i++) {
    var obj = sel[i];
    var item = {
        index: i,
        type: obj.constructor.name,
        id: obj.id || -1
    };
    try { item.name = obj.name || ""; } catch(e) { item.name = ""; }
    try {
        if (obj.geometricBounds) {
            var b = obj.geometricBounds;
            item.bounds = {top: b[0], left: b[1], bottom: b[2], right: b[3]};
        }
    } catch(e) {}
    try {
        if (obj.contents && typeof obj.contents === 'string') {
            item.content_preview = obj.contents.substring(0, 500);
        }
    } catch(e) {}
    try {
        if (obj.appliedParagraphStyle) {
            item.paragraphStyle = obj.appliedParagraphStyle.name;
        }
    } catch(e) {}
    try {
        if (obj.appliedCharacterStyle) {
            item.characterStyle = obj.appliedCharacterStyle.name;
        }
    } catch(e) {}
    try {
        if (obj.appliedObjectStyle) {
            item.objectStyle = obj.appliedObjectStyle.name;
        }
    } catch(e) {}
    try {
        if (obj.fillColor) {
            item.fillColor = obj.fillColor.name;
        }
    } catch(e) {}
    try {
        if (obj.strokeColor) {
            item.strokeColor = obj.strokeColor.name;
        }
    } catch(e) {}
    try {
        if (obj.parentPage) {
            item.page = obj.parentPage.name;
        }
    } catch(e) {}
    items.push(item);
}
__result = {count: sel.length, items: items};
"""


@mcp.tool()
def get_selection(detail_level: str = "basic") -> str:
    """Get information about the current selection in InDesign.

    Args:
        detail_level: "basic" for type/bounds/content, "full" adds styles/colors/page
    """
    err = _check_document()
    if err:
        return _fmt({"success": False, "error": err})

    jsx = _SELECTION_FULL_JSX if detail_level == "full" else _SELECTION_BASIC_JSX

    try:
        result = com.run_jsx(jsx, undo_mode="none")
        return _fmt(_unwrap_result(result))
    except Exception as e:
        return _fmt({"success": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Tool 4: eval_expression
# ---------------------------------------------------------------------------

@mcp.tool()
def eval_expression(expression: str) -> str:
    """Evaluate a short ExtendScript expression in InDesign and return the result.

    Use this for quick read-only queries like checking a property value
    or counting items. No undo wrapping is applied.

    Args:
        expression: The expression to evaluate (e.g. "app.activeDocument.pages.length")
    """
    err = _check_connection()
    if err:
        return _fmt({"success": False, "error": err})

    try:
        result = com.eval_expr(expression)
        if isinstance(result, str) and result.startswith("ERROR: "):
            return _fmt({"success": False, "error": result[7:]})
        return _fmt({"success": True, "result": result})
    except Exception as e:
        return _fmt({"success": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Tool 5: undo
# ---------------------------------------------------------------------------

# Uses __result so the standard wrapper serialises the return value.
_UNDO_JSX_TEMPLATE = """\
var doc = app.activeDocument;
var results = [];
var steps = $STEPS$;
for (var i = 0; i < steps; i++) {
    try {
        var label = doc.undoHistory.length > 0 ? doc.undoHistory[0] : "(empty)";
        doc.undo();
        results.push(label);
    } catch(e) {
        break;
    }
}
__result = {steps_undone: results.length, labels: results};
"""


@mcp.tool()
def undo(steps: int = 1) -> str:
    """Undo the last operation(s) in the active InDesign document.

    Each run_jsx call with undo_mode='entire' creates a single undo step.
    Use this tool to revert agent operations that produced incorrect results.

    Args:
        steps: Number of undo steps to perform (default: 1)
    """
    err = _check_document()
    if err:
        return _fmt({"success": False, "error": err})

    steps = max(1, min(steps, 50))  # Clamp to 1-50

    jsx = _UNDO_JSX_TEMPLATE.replace("$STEPS$", str(steps))

    try:
        result = com.run_jsx(jsx, undo_mode="none")
        return _fmt(_unwrap_result(result))
    except Exception as e:
        return _fmt({"success": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Run the MCP server via stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
