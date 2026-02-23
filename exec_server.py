# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "mcp>=1.0.0",
#     "pywin32>=306",
# ]
# ///
"""
InDesign Exec MCP Server.

Provides 8 tools for executing JSX code in Adobe InDesign via COM/OLE:
  1. run_jsx             - Execute JSX code with undo grouping
  2. get_document_info   - Query active document overview + active view context
  3. get_selection       - Query current selection
  4. eval_expression     - Evaluate a short expression
  5. undo                - Undo last agent operation(s)
  6. report_learning     - Submit local pitfall/best-practice learnings
  7. get_gotchas         - Retrieve curated gotchas (optionally context-filtered)
  8. get_quick_reference - DOM cheatsheet for common access patterns

Requires InDesign Desktop running on Windows.
Uses UndoModes.ENTIRE_SCRIPT to group all operations per call.
No eval() anywhere -- JSX code is inlined, results serialised via __safeStringify.
"""

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

# Add script directory to sys.path so indesign_com can be imported
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

import indesign_com as com

BASE_DIR = Path(__file__).parent
GOTCHAS_PATH = BASE_DIR / "gotchas.json"
SUBMISSIONS_DIR = BASE_DIR / "submissions"
SUBMISSIONS_PATH = SUBMISSIONS_DIR / "pending.jsonl"

ALLOWED_LEARNING_CATEGORIES = {"dom", "scriptui", "extendscript", "execution", "serialization"}
ALLOWED_LEARNING_SEVERITIES = {"blocker", "warning", "tip"}
SEVERITY_RANK = {"tip": 1, "warning": 2, "blocker": 3}

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
        "    __result = doc.pages.everyItem().name;  // returns Array\n\n"
        "DOM QUICK REFERENCE -- common access paths (saves DOM lookups):\n"
        "  Document:     app.activeDocument, app.documents\n"
        "  Windows:      app.activeWindow, doc.layoutWindows, app.activeWindow.activeSpread,\n"
        "                app.activeWindow.activePage\n"
        "  Selection:    app.selection, app.selection[0].constructor.name\n"
        "  Pages:        doc.pages, doc.spreads, doc.sections, page.documentOffset\n"
        "  Page items:   spread.allPageItems, spread.pageItems -- Collections on page/spread/doc:\n"
        "                .textFrames, .rectangles, .ovals, .polygons, .graphicLines, .groups\n"
        "                Constructor types: TextFrame, Rectangle, Oval, Polygon, GraphicLine,\n"
        "                Group, Image, EPS, PDF, GraphicLine\n"
        "  Images:       frame.graphics.length > 0 (test), frame.allGraphics,\n"
        "                graphic.itemLink, link.filePath, link.status\n"
        "  Text:         doc.stories, story.texts, story.paragraphs, story.characters,\n"
        "                story.words, story.insertionPoints, para.contents,\n"
        "                textFrame.parentStory, .nextTextFrame, .previousTextFrame\n"
        "  Styles:       doc.paragraphStyles.itemByName('X'), doc.characterStyles,\n"
        "                doc.objectStyles, .allParagraphStyles, .allCharacterStyles,\n"
        "                .paragraphStyleGroups, .characterStyleGroups, .objectStyleGroups,\n"
        "                obj.appliedParagraphStyle, .appliedCharacterStyle, .appliedObjectStyle\n"
        "  Layers:       doc.layers, doc.activeLayer, item.itemLayer\n"
        "  Geometry:     item.geometricBounds [top,left,bottom,right], item.visibleBounds\n"
        "  Transform:    item.transform(coordSpace, anchor, matrix), item.resolve(),\n"
        "                item.transformValuesOf(), app.transformationMatrices.add(),\n"
        "                CoordinateSpaces (INNER_, PARENT_, PASTEBOARD_, SPREAD_, PAGE_),\n"
        "                AnchorPoint (CENTER_ANCHOR, TOP_LEFT_ANCHOR, ...)\n"
        "  Parent nav:   item.parent, item.parentPage, text.parentStory,\n"
        "                text.parentTextFrames, .range\n"
        "  Swatches:     doc.swatches, doc.colors, doc.gradients,\n"
        "                item.fillColor, item.strokeColor\n"
        "  Groups:       item.groups, group.pageItems, group.allPageItems\n"
        "  Find/Change:  app.findTextPreferences = NothingEnum.NOTHING (CLEAR FIRST),\n"
        "                findGrepPreferences, findChangeGrepOptions\n"
        "  Export:        doc.exportFile(ExportFormat.PDF_TYPE, File(path), false, preset),\n"
        "                app.pdfExportPreferences, jpegExportPreferences, pngExportPreferences\n"
        "  Preferences:  doc.viewPreferences, doc.textFramePreferences, doc.guidePreferences,\n"
        "                doc.gridPreferences, doc.adjustLayoutPreferences, app.generalPreferences\n"
        "  Guides:       page.guides.add(), guide.orientation, guide.location\n"
        "  Hyperlinks:   doc.hyperlinks, doc.hyperlinkTextSources, doc.hyperlinkURLDestinations\n"
        "For a comprehensive cheatsheet with examples, call the get_quick_reference tool.\n\n"
        "LEARNING FEEDBACK LOOP:\n"
        "  - Use report_learning(...) to submit newly discovered bugs, gotchas, and best practices.\n"
        "  - Use get_gotchas(context?) to retrieve curated gotchas before writing complex JSX.\n"
        "  - Maintainers can promote local submissions with: python manage.py review-submissions\n\n"
        "OPERATIONAL POLICY FOR AGENTS (MUST/SHOULD):\n"
        "  - MUST run get_gotchas(context) before implementing a non-trivial InDesign task.\n"
        "  - SHOULD call get_quick_reference() once at the start of the first InDesign task in a conversation.\n"
        "  - MUST use report_learning(...) after a user-reported error/gotcha was resolved and the root cause is clear.\n"
        "  - SHOULD avoid duplicate reports by checking if an equivalent gotcha already exists."
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

## Operational Policy (MUST/SHOULD)
- **MUST** run `get_gotchas(context)` before implementing non-trivial InDesign tasks.
- **SHOULD** call `get_quick_reference()` once when the first InDesign task starts in a conversation.
- **MUST** call `report_learning(...)` after a user-reported bug/gotcha was resolved and root cause + fix are known.
- **SHOULD** avoid duplicate reports when an equivalent gotcha already exists.
- **MUST** classify labels by lifecycle: `_tmp_` for temporary intra-task state, `agentContext_` for intentional cross-session facts.
- **MUST** clear temporary labels at task end with `insertLabel(key, "")`.
- **SHOULD** maintain `_agentLabelRegistry` listing active persistent keys for document auditability.

Maintainer promotion command:
`python manage.py review-submissions`
This reviews entries in `submissions/pending.jsonl` and promotes accepted ones into `gotchas.json`.

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

### Label state lifecycle (multi-call workflows)
    var doc = app.activeDocument;

    // Temporary handoff state, must be cleared at end of task
    doc.insertLabel("_tmp_processedIds", "12,42,73");
    var tmp = doc.extractLabel("_tmp_processedIds");
    var ids = tmp ? tmp.split(",") : [];
    doc.insertLabel("_tmp_processedIds", ""); // cleanup required

    // Persistent context state, keep only if future tasks need it
    doc.insertLabel("agentContext_layoutMap", JSON.stringify({sectionA: [1, 2, 3]}));
    var raw = doc.extractLabel("agentContext_layoutMap");
    var map = raw ? eval("(" + raw + ")") : null;

    // Optional registry for persistent keys
    var regKey = "_agentLabelRegistry";
    var regRaw = doc.extractLabel(regKey);
    var keys = regRaw ? regRaw.split(",") : [];
    if (keys.indexOf("agentContext_layoutMap") < 0) keys.push("agentContext_layoutMap");
    doc.insertLabel(regKey, keys.join(","));

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


def _utc_now_iso() -> str:
    """Return current UTC timestamp as ISO8601."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _load_gotcha_entries() -> list[dict]:
    """Load curated gotcha entries from gotchas.json."""
    if not GOTCHAS_PATH.exists():
        return []
    try:
        data = json.loads(GOTCHAS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []
    entries = data.get("entries", [])
    return entries if isinstance(entries, list) else []


def _context_tokens(context: str) -> list[str]:
    """Tokenize context for lightweight keyword matching."""
    lowered = context.lower()
    return [tok for tok in re.split(r"[^a-z0-9_#]+", lowered) if tok]


def _score_gotcha_for_context(entry: dict, context: str, tokens: list[str]) -> int:
    """Return match score for one gotcha against context."""
    score = 0
    triggers = entry.get("triggers", [])
    if not isinstance(triggers, list):
        return score
    for trigger in triggers:
        needle = str(trigger).strip().lower()
        if not needle:
            continue
        if needle in context:
            score += 2
            continue
        if any(needle in token or token in needle for token in tokens):
            score += 1
    return score


def _normalize_text(text: str) -> str:
    """Normalize text for duplicate detection."""
    return re.sub(r"\s+", " ", text.strip().lower())


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

try {
    var win = app.activeWindow;
    if (win instanceof LayoutWindow) {
        var sp = win.activeSpread;
        var pg = win.activePage;
        __result.activeView = {
            activeSpreadIndex: sp.index,
            activeSpreadPages: sp.pages.length,
            activePageName: pg.name,
            activePageIndex: pg.documentOffset,
            zoom: win.zoomPercentage
        };
        __result.activeSpreadItems = {
            allPageItems: sp.allPageItems.length,
            textFrames: sp.textFrames.length,
            rectangles: sp.rectangles.length,
            ovals: sp.ovals.length,
            images: sp.allGraphics.length,
            groups: sp.groups.length
        };
    }
} catch(viewErr) {}
"""


@mcp.tool()
def get_document_info() -> str:
    """Get an overview of the active InDesign document.

    Returns document name, page count, item counts, selection info,
    style counts, and document preferences. Also includes active view
    context (current spread, page, zoom, and item counts on the active
    spread) when a layout window is open. Read-only operation.
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
# Tool 6: report_learning
# ---------------------------------------------------------------------------

@mcp.tool()
def report_learning(
    problem: str,
    solution: str,
    triggers: list[str],
    category: str = "extendscript",
    severity: str = "warning",
    error_message: str | None = None,
    jsx_context: str | None = None,
) -> str:
    """Submit a local learning entry for later maintainer review.

    This tool does not update curated knowledge directly. It appends
    a pending submission to submissions/pending.jsonl. Maintainers can
    then review and promote accepted learnings into gotchas.json.

    Args:
        problem: One-line description of the pitfall.
        solution: One-line actionable solution.
        triggers: Context keywords used for later matching.
        category: One of dom|scriptui|extendscript|execution|serialization.
        severity: One of blocker|warning|tip.
        error_message: Optional original runtime error text.
        jsx_context: Optional JSX snippet related to the issue.
    """
    if not problem.strip():
        return _fmt({"success": False, "error": "problem must not be empty"})
    if not solution.strip():
        return _fmt({"success": False, "error": "solution must not be empty"})
    if category not in ALLOWED_LEARNING_CATEGORIES:
        return _fmt(
            {
                "success": False,
                "error": f"category must be one of: {sorted(ALLOWED_LEARNING_CATEGORIES)}",
            }
        )
    if severity not in ALLOWED_LEARNING_SEVERITIES:
        return _fmt(
            {
                "success": False,
                "error": f"severity must be one of: {sorted(ALLOWED_LEARNING_SEVERITIES)}",
            }
        )

    cleaned_triggers = [str(t).strip() for t in triggers if str(t).strip()]
    if not cleaned_triggers:
        return _fmt({"success": False, "error": "triggers must contain at least one non-empty value"})

    normalized_problem = _normalize_text(problem)
    normalized_solution = _normalize_text(solution)
    normalized_trigger_set = {str(t).strip().lower() for t in cleaned_triggers}

    for existing in _load_gotcha_entries():
        existing_problem = _normalize_text(str(existing.get("problem", "")))
        existing_solution = _normalize_text(str(existing.get("solution", "")))
        existing_triggers = existing.get("triggers", [])
        existing_trigger_set = {
            str(t).strip().lower() for t in existing_triggers if str(t).strip()
        } if isinstance(existing_triggers, list) else set()
        same_problem = normalized_problem and normalized_problem == existing_problem
        same_solution = normalized_solution and normalized_solution == existing_solution
        trigger_overlap = bool(normalized_trigger_set and existing_trigger_set and normalized_trigger_set & existing_trigger_set)
        if same_problem or (same_solution and trigger_overlap):
            return _fmt(
                {
                    "success": True,
                    "duplicate": True,
                    "message": "Equivalent gotcha already exists; skipping new submission.",
                    "existing_id": existing.get("id"),
                }
            )

    submission = {
        "timestamp": _utc_now_iso(),
        "status": "pending",
        "category": category,
        "severity": severity,
        "triggers": cleaned_triggers,
        "problem": problem.strip(),
        "solution": solution.strip(),
        "error_message": (error_message or "").strip() or None,
        "jsx_context": (jsx_context or "").strip() or None,
    }

    try:
        SUBMISSIONS_DIR.mkdir(parents=True, exist_ok=True)
        with SUBMISSIONS_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(submission, ensure_ascii=False) + "\n")
    except Exception as e:
        return _fmt({"success": False, "error": f"failed to persist submission: {e}"})

    return _fmt(
        {
            "success": True,
            "message": "Learning submitted to local review queue.",
            "submission_file": str(SUBMISSIONS_PATH),
            "next_step": "Run `python manage.py review-submissions` to review and promote entries.",
            "submission": submission,
        }
    )


# ---------------------------------------------------------------------------
# Tool 7: get_gotchas
# ---------------------------------------------------------------------------

@mcp.tool()
def get_gotchas(
    context: str | None = None,
    min_severity: str = "tip",
    top_n: int | None = None,
) -> str:
    """Get curated gotchas, optionally filtered by context keywords.

    Args:
        context: Optional context string (for example "modeless palette move by UnitValue")
                 used to rank matching gotchas.
        min_severity: Minimum severity to include: tip|warning|blocker (default: tip).
        top_n: Optional maximum number of entries to return.
    """
    if min_severity not in ALLOWED_LEARNING_SEVERITIES:
        return _fmt(
            {
                "success": False,
                "error": f"min_severity must be one of: {sorted(ALLOWED_LEARNING_SEVERITIES)}",
            }
        )
    if top_n is not None and top_n <= 0:
        return _fmt({"success": False, "error": "top_n must be > 0 when provided"})

    entries = _load_gotcha_entries()
    min_rank = SEVERITY_RANK[min_severity]
    filtered_entries = [
        entry
        for entry in entries
        if SEVERITY_RANK.get(str(entry.get("severity", "tip")).lower(), 0) >= min_rank
    ]

    if not context:
        if top_n is not None:
            filtered_entries = filtered_entries[:top_n]
        return _fmt(
            {
                "success": True,
                "min_severity": min_severity,
                "count": len(filtered_entries),
                "entries": filtered_entries,
            }
        )

    lowered_context = context.lower()
    tokens = _context_tokens(context)
    scored: list[tuple[int, dict]] = []
    for entry in filtered_entries:
        score = _score_gotcha_for_context(entry, lowered_context, tokens)
        if score > 0:
            scored.append((score, entry))

    scored.sort(key=lambda item: item[0], reverse=True)
    ranked_entries = [entry for _, entry in scored]
    if top_n is not None:
        ranked_entries = ranked_entries[:top_n]
    return _fmt(
        {
            "success": True,
            "context": context,
            "min_severity": min_severity,
            "count": len(ranked_entries),
            "entries": ranked_entries,
        }
    )


# ---------------------------------------------------------------------------
# Tool 8: get_quick_reference
# ---------------------------------------------------------------------------

_QUICK_REFERENCE = """\
# InDesign ExtendScript Quick Reference

## Document & Windows
  app.activeDocument                        Current document
  app.documents                             All open documents
  app.activeWindow                          Current window (LayoutWindow or StoryWindow)
  doc.layoutWindows                         All layout windows for this document
  app.activeWindow.activeSpread             Currently visible spread
  app.activeWindow.activePage               Currently active page

## Pages, Spreads & Sections
  doc.pages                                 All pages
  doc.pages[0], doc.pages.itemByName('3')   Access by index or name
  doc.spreads                               All spreads
  doc.masterSpreads                         All master spreads
  doc.sections                              Document sections (numbering)
  page.documentOffset                       Absolute page index in document
  page.appliedMaster                        Master spread applied to page

## Selection
  app.selection                             Array of selected objects
  app.selection[0].constructor.name         Type of first selected item
  // Common types: TextFrame, Rectangle, Oval, Polygon, GraphicLine,
  //   Group, Image, Text, InsertionPoint, Character, Word, Paragraph

## Page Items (on page, spread, document, or group)
  container.allPageItems                    ALL items (recursive into groups)
  container.pageItems                       Direct children only
  container.textFrames                      TextFrame collection
  container.rectangles                      Rectangle collection (often image frames)
  container.ovals                           Oval collection
  container.polygons                        Polygon collection
  container.graphicLines                    GraphicLine collection
  container.groups                          Group collection
  // 'container' = doc, page, spread, group, or masterSpread
  // Constructor types: TextFrame, Rectangle, Oval, Polygon, GraphicLine,
  //   Group, Image, EPS, PDF, GraphicLine

## Images, Graphics & Links
  frame.graphics                            Graphics placed inside a frame
  frame.graphics.length > 0                 Test: frame contains placed image?
  frame.allGraphics                         All graphics (recursive into groups)
  frame.graphics[0]                         The Image/EPS/PDF object
  graphic.itemLink                          Link object for placed file
  link.filePath                             Full path to linked file
  link.status                               LinkStatus.NORMAL, LINK_MISSING, etc.
  doc.links                                 All links in the document
  doc.allGraphics                           All graphics across entire document

## Text, Stories & Threading
  doc.stories                               All stories (text flows)
  story.texts                               All text ranges in a story
  story.paragraphs                          Paragraph collection
  story.characters                          Character collection
  story.words                               Word collection
  story.insertionPoints                     Cursor positions between characters
  story.contents                            Full text content (read/write)
  paragraph.contents                        Text of a single paragraph
  text.range(startIdx, endIdx)              Subrange of text
  textFrame.parentStory                     Story that this frame belongs to
  textFrame.nextTextFrame                   Next frame in thread chain
  textFrame.previousTextFrame               Previous frame in thread chain
  text.parentTextFrames                     Array of frames containing a text range
  textFrame.textFramePreferences            Threading, columns, inset, auto-size

## Styles & Style Groups
  doc.paragraphStyles                       Paragraph styles (top level)
  doc.characterStyles                       Character styles (top level)
  doc.objectStyles                          Object styles (top level)
  doc.allParagraphStyles                    All para styles (flat, incl. groups) [readonly]
  doc.allCharacterStyles                    All char styles (flat, incl. groups) [readonly]
  doc.paragraphStyleGroups                  Paragraph style group folders
  doc.characterStyleGroups                  Character style group folders
  doc.objectStyleGroups                     Object style group folders
  doc.paragraphStyles.itemByName('X')       Access style by name
  obj.appliedParagraphStyle                 Style applied to text/paragraph
  obj.appliedCharacterStyle                 Style applied to text/characters
  obj.appliedObjectStyle                    Style applied to a page item

## Layers
  doc.layers                                All layers
  doc.activeLayer                           Currently active layer
  item.itemLayer                            Layer an item lives on (read/write)
  doc.layers.itemByName('X')               Access layer by name

## Geometry & Bounds
  item.geometricBounds                      [top, left, bottom, right] in doc units
  item.visibleBounds                        Bounds including stroke weight

## Coordinate Spaces & Transformations
  // InDesign uses a matrix-based transformation system.
  // Every page item has its own coordinate space (INNER), and transformations
  // describe how INNER maps to PARENT, SPREAD, PAGE, or PASTEBOARD.

  // CoordinateSpaces:
  //   INNER_COORDINATES      Item's own local space
  //   PARENT_COORDINATES     Relative to parent (group, page, spread)
  //   SPREAD_COORDINATES     Relative to the spread origin
  //   PAGE_COORDINATES       Relative to the page origin
  //   PASTEBOARD_COORDINATES Absolute pasteboard space

  // AnchorPoints (reference points for transforms):
  //   TOP_LEFT_ANCHOR, TOP_CENTER_ANCHOR, TOP_RIGHT_ANCHOR,
  //   CENTER_LEFT_ANCHOR, CENTER_ANCHOR, CENTER_RIGHT_ANCHOR,
  //   BOTTOM_LEFT_ANCHOR, BOTTOM_CENTER_ANCHOR, BOTTOM_RIGHT_ANCHOR

  // Move items:
  item.move(undefined, [deltaX, deltaY])    Relative move by offset
  item.move([x, y])                         Absolute move to position

  // Resize items:
  item.resize(CoordinateSpaces.INNER_COORDINATES,
    AnchorPoint.CENTER_ANCHOR,
    ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
    [scaleX, scaleY])

  // Create and apply a transformation matrix:
  var matrix = app.transformationMatrices.add({
    counterclockwiseRotationAngle: 45,
    horizontalScaleFactor: 1.5,
    verticalScaleFactor: 1.5,
    horizontalTranslation: 10,
    verticalTranslation: 20
  });
  item.transform(CoordinateSpaces.PASTEBOARD_COORDINATES,
    AnchorPoint.CENTER_ANCHOR, matrix);

  // Resolve: get position of an anchor in a given coordinate space
  item.resolve(AnchorPoint.CENTER_ANCHOR,
    CoordinateSpaces.SPREAD_COORDINATES)[0]   // returns [[x, y]]

  // Read current transform values:
  item.transformValuesOf(CoordinateSpaces.PARENT_COORDINATES)
  // Returns TransformationMatrix array (rotation, scale, shear, translation)

## Swatches & Colors
  doc.swatches                              All swatches (colors + gradients)
  doc.colors                                Color swatches only
  doc.gradients                             Gradient swatches only
  doc.swatches.itemByName('X')              Access swatch by name
  item.fillColor                            Fill swatch (read/write)
  item.strokeColor                          Stroke swatch (read/write)
  item.strokeWeight                         Stroke weight in points
  item.opacity                              Opacity 0-100

## Parent Navigation Principle
  item.parent                               Direct parent (group, page, spread, ...)
  item.parentPage                           Page containing the item (null if pasteboard)
  text.parentStory                          Story containing the text
  text.parentTextFrames                     Array of frames containing the text range
  insertionPoint.parentTextFrame            Frame at insertion point

## Groups
  container.groups                          Group collection
  group.pageItems                           Direct children of group
  group.allPageItems                        All items recursive inside group
  group.ungroup()                           Ungroup the group

## Find/Change (Text & GREP)
  // ALWAYS clear prefs first:
  app.findTextPreferences = NothingEnum.NOTHING;
  app.changeTextPreferences = NothingEnum.NOTHING;
  app.findTextPreferences.findWhat = "old";
  app.changeTextPreferences.changeTo = "new";
  var results = doc.changeText();    // returns changed items
  __result = {replaced: results.length};

  // GREP find/change:
  app.findGrepPreferences = NothingEnum.NOTHING;
  app.changeGrepPreferences = NothingEnum.NOTHING;
  app.findGrepPreferences.findWhat = "\\\\d+";  // regex
  var found = doc.findGrep();
  __result = {found: found.length};

  // findChangeGrepOptions / findChangeTextOptions:
  app.findChangeGrepOptions.includeLockedLayersForFind = false;
  app.findChangeGrepOptions.includeLockedStoriesForFind = false;
  app.findChangeGrepOptions.includeMasterPages = false;

## Label State & Lifecycle
  doc.insertLabel("key", "value")              Persist string value in document
  doc.extractLabel("key")                      Read value (returns "" when missing)
  pageItem.insertLabel("key", "value")         Persist value on individual object
  pageItem.extractLabel("key")                 Read value from object
  doc.insertLabel("key", "")                   Delete/clear a label key
  // Temporary keys: _tmp_*  -> must be cleaned up at task end
  // Persistent keys: agentContext_* -> only if useful across future tasks
  // Optional: _agentLabelRegistry = comma-separated persistent keys
  // Avoid persisting toSource() index-path references across structural edits

## Reusable Utility Patterns (Agent-safe)
  // Pattern from legacy util scripts: always reset find/change prefs before AND after use.
  function safeFindGrep(target, findProps, changeProps, optionProps) {
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;
    app.findGrepPreferences.properties = findProps || {};
    if (changeProps) app.changeGrepPreferences.properties = changeProps;
    if (optionProps) app.findChangeGrepOptions.properties = optionProps;
    var out = changeProps ? target.changeGrep() : target.findGrep();
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;
    return out;
  }

  // Pattern from legacy util scripts: style lookup across nested groups.
  function getParagraphStylesByNameDeep(doc, styleName) {
    var found = [];
    function walk(parent) {
      var i;
      for (i = 0; i < parent.paragraphStyles.length; i++) {
        if (parent.paragraphStyles[i].name === styleName) found.push(parent.paragraphStyles[i]);
      }
      for (i = 0; i < parent.paragraphStyleGroups.length; i++) {
        walk(parent.paragraphStyleGroups[i]);
      }
    }
    walk(doc);
    return found;
  }

  // Do not port UI/session-bound helpers into MCP agents: alert(), exit(), activeScript path helpers.
  // Do not reuse known-problematic helpers without fixes: incomplete glyph wrappers, buggy multiReplace variants.

## Export
  // PDF with preset:
  var preset = app.pdfExportPresets.itemByName('[High Quality Print]');
  doc.exportFile(ExportFormat.PDF_TYPE, File('/path/output.pdf'), false, preset);

  // PDF with custom prefs:
  app.pdfExportPreferences.pageRange = "1-5";
  doc.exportFile(ExportFormat.PDF_TYPE, File('/path/output.pdf'));

  // JPEG:
  app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM;
  doc.exportFile(ExportFormat.JPG, File('/path/output.jpg'));

  // PNG:
  doc.exportFile(ExportFormat.PNG_FORMAT, File('/path/output.png'));
  // app.pngExportPreferences for PNG settings

  // HTML:
  doc.exportFile(ExportFormat.HTML_FPG, File('/path/output.html'));
  // app.htmlExportPreferences for HTML settings

## Hyperlinks
  doc.hyperlinks                            All hyperlinks
  doc.hyperlinkTextSources                  Text-based link sources
  doc.hyperlinkURLDestinations              URL destinations

## Preferences
  doc.viewPreferences                       Measurement units, rulers, guides display
  doc.textFramePreferences                  Default text frame settings
  doc.guidePreferences                      Guide display and behavior
  doc.gridPreferences                       Grid settings
  doc.adjustLayoutPreferences               Adjust layout rules
  doc.documentPreferences                   Page size, facing pages, page count
  app.generalPreferences                    Application-wide preferences
  app.pdfExportPreferences                  PDF export settings
  app.jpegExportPreferences                 JPEG export settings
  app.pngExportPreferences                  PNG export settings
  app.htmlExportPreferences                 HTML export settings

## Guides
  page.guides                               Guides on a page
  page.guides.add()                         Add a guide
  guide.orientation                         HorizontalOrVertical.HORIZONTAL / VERTICAL
  guide.location                            Position in document units

## Common ExtendScript Gotchas
  - geometricBounds order: [top, left, bottom, right] (NOT x,y,w,h)
  - NothingEnum.NOTHING to clear find/change prefs (REQUIRED before each search)
  - ExtendScript is ES3: no let/const, no arrow functions, no template literals
  - File paths: new File('/c/path/to/file') or File('~/Desktop/file.pdf')
  - Check collection.length before indexing to avoid errors
  - try/catch around .parentPage (null for items on pasteboard)
  - Assign to __result (no return statement in run_jsx)
  - Use everyItem() for bulk ops (see Collection Patterns in server instructions)
"""


@mcp.tool()
def get_quick_reference() -> str:
    """Get a comprehensive DOM cheatsheet for InDesign ExtendScript.

    Call this BEFORE writing complex JSX scripts to see common DOM
    access patterns, object hierarchy, and property names. This saves
    multiple DOM lookup calls. Covers: document navigation, page items,
    images, text, styles, geometry, transformations, layers, colors,
    find/change, export, and common gotchas.
    """
    entries = _load_gotcha_entries()
    community_lines = []
    for entry in entries:
        severity = str(entry.get("severity", "")).lower()
        if severity not in {"blocker", "warning"}:
            continue
        problem = str(entry.get("problem", "")).strip()
        solution = str(entry.get("solution", "")).strip()
        entry_id = str(entry.get("id", "")).strip() or "unknown-id"
        if not problem or not solution:
            continue
        community_lines.append(f"  - [{severity}] {entry_id}: {problem} -> {solution}")

    if not community_lines:
        return _QUICK_REFERENCE

    dynamic_section = "\n## Community Gotchas (from gotchas.json)\n" + "\n".join(community_lines) + "\n"
    return _QUICK_REFERENCE + dynamic_section


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Run the MCP server via stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
