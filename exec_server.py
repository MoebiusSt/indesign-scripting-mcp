# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "mcp>=1.0.0",
#     "pywin32>=306",
# ]
# ///
"""
InDesign Exec MCP Server.

Provides 11 tools for executing JSX code in Adobe InDesign via COM/OLE:
  1. run_jsx             - Execute JSX code with undo grouping
  2. get_document_info   - Query active document overview + active view context
  3. get_selection       - Query current selection
  4. eval_expression     - Evaluate a short expression
  5. undo                - Undo last agent operation(s)
  6. report_learning     - Submit local pitfall/best-practice learnings
  7. get_gotchas         - Retrieve curated gotchas (optionally context-filtered)
  8. get_quick_reference - DOM cheatsheet for common access patterns
  9. get_change_log      - Objects changed by mutating run_jsx calls
 10. preflight_document  - Run a preflight profile, report findings (read-only)
 11. typography_audit    - Fonts, sizes, style usage and overrides (read-only)

Requires InDesign Desktop running on Windows.
Uses UndoModes.ENTIRE_SCRIPT to group all operations per call.
User code is inlined, never run via eval(); results are serialised via __safeStringify.
JSON.parse is available in ExtendScript through the polyfill (validated JSON only).
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
CHANGE_LOG_DIR = BASE_DIR / "logs"
CHANGE_LOG_PATH = CHANGE_LOG_DIR / "change_log.jsonl"
CHANGE_LOG_LIST_CAP = 20

ALLOWED_LEARNING_CATEGORIES = {"dom", "scriptui", "extendscript", "uxp", "execution", "serialization"}
ALLOWED_LEARNING_SEVERITIES = {"blocker", "warning", "tip"}
SEVERITY_RANK = {"tip": 1, "warning": 2, "blocker": 3}

mcp = FastMCP(
    "InDesign Exec",
    instructions=(
        "This server executes ExtendScript (JSX) and UXP JavaScript code directly in a running Adobe InDesign instance. "
        "Use run_jsx to perform operations in InDesign documents. Every run_jsx call with "
        "undo_mode='entire' (default) groups all changes under a single undo step, so "
        "Ctrl+Z reverts the entire operation. Use the undo tool to programmatically revert changes.\n\n"
        "EXECUTION ENGINES & LANGUAGE SELECTION:\n"
        "  run_jsx supports language='javascript' (ExtendScript ES3, DEFAULT) and language='uxp' (ES6+ / V8).\n"
        "  - DEFAULT TO ExtendScript (language='javascript'): massive LLM training corpus, zero-overhead DOM calls, "
        "proven collection API (everyItem), unconstrained File/Folder access. Use for 95% of tasks.\n"
        "  - USE UXP (language='uxp') SELECTIVELY: when modern JS simplifies complex data processing "
        "(Array map/filter/reduce, destructuring, Set/Map, top-level await / async functions, native JSON).\n"
        "  - In UXP code: ALWAYS import InDesign objects with: const { app, ...enums } = require('indesign'); and assign return value to __result.\n"
        "  - Note on Network / Fetch: Standalone InDesign UXP scripts via COM do not pump background socket loops; "
        "perform HTTP/API fetches in the agent/MCP layer and pass data into scripts.\n\n"
        "IMPORTANT: In your code, assign the return value to the variable __result. Example:\n"
        "  var doc = app.activeDocument;\n"
        "  __result = {name: doc.name, pages: doc.pages.length};\n\n"
        "Do NOT use return statements — the wrapper handles serialisation.\n"
        "Always verify results after modifications using eval_expression or get_document_info.\n"
        "ERRORS: run_jsx reports `line` relative to YOUR code plus `source_line`; syntax errors are located "
        "without executing anything. Pass debug=true for surrounding source lines (and the mapped UXP stack).\n"
        "CHANGE LOG: mutating run_jsx calls record added/removed page items of the host document; call "
        "__mcpTrack(obj, 'note') in your code to log modified objects. Query with get_change_log.\n"
        "READ-ONLY CHECKS: preflight_document (preflight profile findings), typography_audit (fonts, sizes, "
        "style usage, overrides, overset).\n"
        "Use the InDesign DOM MCP server to look up classes, properties and methods "
        "before writing script code.\n\n"
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
        "TOXIC NULL HYPERLINKS (blocker):\n"
        "  Copied hyperlinks with destination===null can OPEN the origin .indd when .destination\n"
        "  is read (or showDestination). run_jsx / eval_expression detect documents that appear during\n"
        "  a script. Saved, unmodified ones are closed (close_unexpected_documents=true, default).\n"
        "  Modified or untitled ones stay open. Whenever any appeared, the result has a warning:\n"
        "  treat it as a stop sign and verify the host before the next edit. Never trust\n"
        "  activeDocument after hyperlink scans — re-address the host by name/id. Scan/repair via\n"
        "  $.evalFile(File(Folder.userData + \"/Adobe/InDesign/Version 21.0/de_DE/Scripts/Scripts Panel/_lib/toxicHyperlinks.jsxinc\"));\n"
        "  then THL.scan(doc) / THL.sanitize(...). Gotcha: toxic-null-hyperlink-opens-origin-document.\n"
        "  Orphan sources block \"New hyperlink\": THL.findOrphanSources(doc) / THL.removeOrphanSources(doc, ids).\n\n"
        "LEARNING FEEDBACK LOOP:\n"
        "  - Use report_learning(...) to submit newly discovered bugs, gotchas, and best practices.\n"
        "  - Use get_gotchas(context?) to retrieve curated gotchas before writing complex scripts.\n"
        "  - Maintainers can promote local submissions with: python manage.py review-submissions\n\n"
        "OPERATIONAL POLICY FOR AGENTS (MUST/SHOULD):\n"
        "  - MUST run get_gotchas(context) before implementing a non-trivial InDesign task.\n"
        "  - SHOULD call get_quick_reference() once at the start of the first InDesign task in a conversation.\n"
        "  - MUST use report_learning(...) after a user-reported error/gotcha was resolved and the root cause is clear.\n"
        "  - SHOULD avoid duplicate reports by checking if an equivalent gotcha already exists.\n"
        "  - MUST treat activeDocument as unstable after any hyperlink.destination / showDestination access.\n"
        "  - MUST leave close_unexpected_documents=true unless intentionally opening another document."
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

## Language Selection: ExtendScript vs UXP

`run_jsx` supports two script engines via the `language` parameter:

| Engine | Parameter | JavaScript Standard | Best For |
|---|---|---|---|
| **ExtendScript (Default)** | `language="javascript"` | ES3 | 95% of tasks: layout, typography, styles, find/change, tables, File/Folder IO. Highest model accuracy. |
| **UXP Scripting** | `language="uxp"` | ES6+ / V8 | Complex data transformation (`map`/`filter`/`reduce`, `Set`, `Map`, destructuring, top-level `await` / `Promise.all`), native JSON. |

### Decision Rule for Agents
**Default to ExtendScript (`language="javascript"`).** Only switch to `language="uxp"` when modern JS features (array processing, Set/Map, async data loops, native JSON) significantly simplify the code.

### Writing UXP Scripts (`language="uxp"`)
In UXP scripts, InDesign DOM objects and enums must be imported via `require('indesign')`:

    const { app, Justification, UndoModes } = require('indesign');
    const doc = app.activeDocument;
    const items = doc.pages.everyItem().name.map(name => ({ pageName: name }));
    __result = { count: items.length, items: items };

The MCP wrapper wraps UXP code in an `async` IIFE (enabling top-level `await`), safely polls for async Promise completion, catches runtime exceptions, and automatically dismisses modal syntax-error dialogs.

Note: Standalone InDesign scripts via COM do not pump background socket loops for `fetch()`. Perform network/HTTP fetches in the agent/MCP layer and pass data into scripts.

## Undo Modes
- `undo_mode="entire"` (default) — groups ALL changes as one Ctrl+Z step.
  Always provide a descriptive `undo_name` like "Agent: Format headings".
- `undo_mode="none"` — for read-only queries. No undo step created.
- `undo_mode="auto"` — each DOM change gets its own undo step.

## Errors and Debugging
- Error results carry `line` relative to your code (1-based) and `source_line`.
- Syntax errors are located by a compile-only probe (`source: "syntax check (nothing was executed)"`).
- `debug=true` adds `context` (±3 source lines) and, for UXP, the stack mapped to your lines.

## Change Log
- Mutating calls (`undo_mode` "entire"/"auto", ExtendScript) return `changes`: host document,
  added page items (id/type/page), removed ids, page/story deltas.
- Property edits are not diffed. Record them with `__mcpTrack(obj, "note")`; it returns `obj`:

      __mcpTrack(tf, "contents").contents = "New text";

- `get_change_log()` lists the entries of this session (`all_sessions=true` for older ones).

## Read-only Checks
- `preflight_document(profile?, document?)` — findings of a preflight profile, profile source, completeness.
- `typography_audit(document?)` — fonts with status, size/leading pairs, paragraph/character style usage,
  local overrides, unstyled paragraphs, overset stories (samples carry page names).

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
- **MUST** treat `app.activeDocument` as unstable after any `hyperlink.destination` / `showDestination` access.
- **MUST** leave `close_unexpected_documents=true` (default). It only closes saved unmodified foreign docs; a `warning` means stop and verify the host.
- **MUST** load `toxicHyperlinks.jsxinc` with `$.evalFile` in every `run_jsx` that uses `THL.*` (globals do not persist between calls).

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
    var map = raw ? JSON.parse(raw) : null;

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
- `JSON.stringify` / `JSON.parse` are available in ExtendScript (polyfill; `JSON.parse` accepts strict JSON only).
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
    Preserves document-pinning fields from the wrapper envelope (success and error).
    """
    out: dict
    if result.get("success") and isinstance(result.get("result"), dict):
        data = result["result"]
        out = {"success": True, **data}
    else:
        out = dict(result)
        # Line numbers of fixed tool JSX mean nothing to the caller.
        out.pop("line", None)
        out.pop("source_line", None)
    for key in ("closed_unexpected_documents", "unexpected_documents", "warning"):
        val = result.get(key)
        if val:
            out[key] = val
    return out


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
    language: str = "javascript",
    close_unexpected_documents: bool = True,
    debug: bool = False,
    track_changes: bool | None = None,
) -> str:
    """Execute ExtendScript (JSX) or UXP JavaScript code in InDesign.

    The code runs inside a safety wrapper that catches errors and returns
    structured results. When undo_mode is 'entire' (default), all changes
    are grouped under a single undo step labeled with undo_name.

    IMPORTANT: Assign the value you want to return to the variable __result.
    Example (ExtendScript - Default):
        var doc = app.activeDocument;
        __result = {name: doc.name, pages: doc.pages.length};

    Example (UXP - Modern JS):
        const { app } = require('indesign');
        const doc = app.activeDocument;
        const names = doc.pages.everyItem().name;
        __result = { pages: names.map(n => ({ name: n })) };

    LANGUAGE SELECTION:
      - language='javascript' (ExtendScript ES3, DEFAULT): Recommended for 95% of tasks.
        Unrivalled model world knowledge, maximum stability, unconstrained File/Folder access.
      - language='uxp' (Modern JS ES6+ / V8): Use selectively when complex data transformations
        (Array map/filter/reduce, destructuring, Set/Map, async/await), native fetch, or native JSON
        handling are clearly advantageous.
        NOTE: In UXP, always import DOM with `const { app } = require('indesign');`.

    PERFORMANCE — prefer InDesign Collection methods over loops:
        // GOOD: bulk operation via everyItem() — one command for all items
        doc.textFrames.everyItem().label = "processed";
        __result = doc.pages.everyItem().name;  // reads all names at once
        // GOOD: direct access without loop
        var style = doc.paragraphStyles.itemByName("Heading 1");
        // BAD: manual loop when everyItem() would work
        for (var i = 0; i < doc.textFrames.length; i++) { doc.textFrames[i].label = "processed"; }
    Only use loops when each element needs a DIFFERENT value.

    DOCUMENT PINNING:
      Detects documents that appear during a script (by document id). Saved, unmodified
      foreign docs are closed when close_unexpected_documents=true (default); modified or
      untitled ones stay open. Whenever any foreign doc appeared, the result has warning
      and unexpected_documents (plus closed_unexpected_documents for closed paths). Treat
      warning as a stop sign and verify the host before the next edit. Set false only when
      the script intentionally opens another file that must remain open (detection still runs).

    ERRORS: `line` is relative to YOUR code (1-based) and `source_line` holds its text.
      Syntax errors are located by a compile-only probe; nothing was executed then.
      debug=True adds `context` (±3 lines) and, for UXP, the stack mapped to your lines.

    CHANGE LOG: with undo_mode 'entire'/'auto' (ExtendScript), added/removed page items,
      page/story deltas of the host document are returned as `changes` and logged
      (see get_change_log). Property edits are not diffed: wrap modified objects with
      __mcpTrack(obj, "note") — it returns obj, so `__mcpTrack(tf).contents = "x"` works.

    Args:
        code: The script code to execute. Assign to __result to return data.
        undo_name: Human-readable label for Edit > Undo (e.g. "Agent: Format headings")
        undo_mode: "entire" groups all changes as one undo step (default),
                   "auto" lets InDesign handle undo per-operation,
                   "none" skips undo tracking (for read-only operations)
        language: "javascript" for ExtendScript ES3 (default) or "uxp" for modern JS (ES6+/V8)
        close_unexpected_documents: Close foreign docs opened mid-script (default True)
        debug: Return source context around an error (default False)
        track_changes: Diff page items of the host document. Default: on unless undo_mode="none".
    """
    err = _check_connection()
    if err:
        return _fmt({"success": False, "error": err})

    if track_changes is None:
        track_changes = undo_mode != "none"

    try:
        result = com.run_jsx(
            code,
            undo_name=undo_name,
            undo_mode=undo_mode,
            language=language,
            close_unexpected_documents=close_unexpected_documents,
            track_changes=track_changes,
            debug=debug,
        )
    except Exception as e:
        result = {"success": False, "error": str(e), "name": type(e).__name__}

    if undo_mode != "none" or result.get("changes"):
        _record_change(result, code, undo_name, undo_mode, language)
    return _fmt(result)


# ---------------------------------------------------------------------------
# Change log (fed by run_jsx)
# ---------------------------------------------------------------------------

SESSION_ID = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
_session_changes: list[dict] = []


def _record_change(result: dict, code: str, undo_name: str, undo_mode: str, language: str) -> None:
    """Append one run_jsx call to the in-memory and on-disk change log."""
    changes = result.get("changes") or {}
    entry = {
        "timestamp": _utc_now_iso(),
        "session": SESSION_ID,
        "undo_name": undo_name,
        "undo_mode": undo_mode,
        "language": language,
        "success": bool(result.get("success")),
        "document": changes.get("host_document"),
    }
    if not entry["success"]:
        entry["error"] = result.get("error")
        entry["line"] = result.get("line")
        entry["partial_changes_possible"] = undo_mode != "none"
    for key in ("added_count", "removed_count", "pages_delta", "stories_delta", "tracked_count",
                "id_diff_skipped"):
        if changes.get(key):
            entry[key] = changes[key]
    for key in ("added", "removed_ids", "tracked"):
        if changes.get(key):
            entry[key] = changes[key][:CHANGE_LOG_LIST_CAP]
    if result.get("warning"):
        entry["warning"] = result["warning"]
    entry["code_preview"] = code[:300]

    _session_changes.append(entry)
    try:
        CHANGE_LOG_DIR.mkdir(parents=True, exist_ok=True)
        with CHANGE_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


@mcp.tool()
def get_change_log(
    limit: int = 20,
    document: str | None = None,
    all_sessions: bool = False,
) -> str:
    """List what mutating run_jsx calls changed (newest first).

    Each entry: undo_name, host document, success, added page items (id/type/page),
    removed ids, page/story deltas and objects recorded via __mcpTrack(obj, note).
    Use it for the completion report and to locate objects to verify or revert.

    Args:
        limit: Maximum number of entries (default 20).
        document: Only entries whose host document name or path contains this text.
        all_sessions: Include earlier server sessions from logs/change_log.jsonl.
    """
    limit = max(1, min(limit, 500))
    entries: list[dict]
    if all_sessions and CHANGE_LOG_PATH.exists():
        entries = []
        for line in CHANGE_LOG_PATH.read_text(encoding="utf-8").splitlines():
            try:
                entries.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    else:
        entries = list(_session_changes)

    if document:
        needle = document.lower()
        entries = [
            e for e in entries
            if e.get("document") and (
                needle in str(e["document"].get("name", "")).lower()
                or needle in str(e["document"].get("path", "")).lower()
            )
        ]
    entries = entries[-limit:][::-1]
    return _fmt({"success": True, "session": SESSION_ID, "count": len(entries), "entries": entries})


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
def eval_expression(
    expression: str,
    close_unexpected_documents: bool = True,
) -> str:
    """Evaluate a short ExtendScript expression in InDesign and return the result.

    Use this for quick read-only queries like checking a property value
    or counting items. No undo wrapping is applied.

    Args:
        expression: The expression to evaluate (e.g. "app.activeDocument.pages.length")
        close_unexpected_documents: Close foreign docs opened as a side effect
            (default True). Keep True when probing hyperlink.destination.
    """
    err = _check_connection()
    if err:
        return _fmt({"success": False, "error": err})

    try:
        result = com.eval_expr(
            expression,
            close_unexpected_documents=close_unexpected_documents,
        )
        warning = None
        if isinstance(result, str) and "\n[MCP WARNING] " in result:
            result, warning = result.split("\n[MCP WARNING] ", 1)
        if isinstance(result, str) and result.startswith("ERROR: "):
            payload = {"success": False, "error": result[7:]}
            if warning:
                payload["warning"] = warning
            return _fmt(payload)
        payload = {"success": True, "result": result}
        if warning:
            payload["warning"] = warning
        return _fmt(payload)
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
  // BLOCKER: copied hyperlinks with destination===null can OPEN the origin .indd
  // when .destination is read (or showDestination). Not HyperlinkExternalPageDestination.
  // run_jsx / eval_expression detect docs that appear mid-script. Saved unmodified ones
  // are closed (close_unexpected_documents=true); modified/untitled stay open. A warning
  // means stop and verify the host. Scan/repair:
  //   $.evalFile(File(Folder.userData + "/Adobe/InDesign/Version 21.0/de_DE/Scripts/Scripts Panel/_lib/toxicHyperlinks.jsxinc"));
  //   THL.scan(doc);  THL.sanitize(doc, ids, "clean-null");
  // Orphan sources block "New hyperlink": THL.findOrphanSources(doc) / THL.removeOrphanSources(doc, ids)
  // Gotcha: toxic-null-hyperlink-opens-origin-document

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
  - Toxic null hyperlinks: reading .destination may open another .indd; treat warning as stop

## UXP Scripting Reference (language='uxp')
  // Use language='uxp' in run_jsx when ES6+ / V8 features are specifically beneficial.
  // 1. Import InDesign DOM & Enums:
  const { app, Justification, UndoModes, ExportFormat } = require('indesign');
  
  // 2. Modern JS Data Operations:
  const doc = app.activeDocument;
  const pageNames = doc.pages.everyItem().name;
  const processed = pageNames
    .filter(name => name !== 'Cover')
    .map((name, idx) => ({ page: name, index: idx }));
  
  // 3. Set, Map, Destructuring, Template Literals:
  const uniqueStyles = new Set(doc.allParagraphStyles.map(s => s.name));
  const { width, height } = { width: 210, height: 297 };
  const label = `Doc: ${doc.name} (${width}x${height}mm)`;
  
  // 4. Return via __result:
  __result = { summary: label, stylesCount: uniqueStyles.size, pages: processed };
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
# Shared JSX: resolve the target document by name, full path or id
# ---------------------------------------------------------------------------

_RESOLVE_DOC_JSX = """\
function __mcpResolveDoc(spec) {
    if (!app.documents.length) throw new Error("No document open in InDesign.");
    if (spec === null || spec === "") return app.activeDocument;
    var names = [];
    for (var i = 0; i < app.documents.length; i++) {
        var d = app.documents[i];
        var path = "";
        try { path = String(d.fullName.fsName); } catch (e) {}
        if (String(d.name) === spec || path === spec || String(d.id) === spec) return d;
        names.push(String(d.name));
    }
    throw new Error("Document not found: " + spec + ". Open documents: " + names.join(", "));
}
"""


def _doc_spec_literal(document: str | None) -> str:
    """Encode the document selector as a JS literal for __mcpResolveDoc."""
    return json.dumps(document.strip()) if document and document.strip() else "null"


# ---------------------------------------------------------------------------
# Tool 10: preflight_document
# ---------------------------------------------------------------------------

_PREFLIGHT_JSX = _RESOLVE_DOC_JSX + """\
var doc = __mcpResolveDoc($DOC$);
var wanted = $PROFILE$;
var maxFindings = $MAX$;

function findProfile(coll, name) {
    for (var i = 0; i < coll.length; i++) {
        if (String(coll[i].name) === name) return coll[i];
    }
    return null;
}
function profileNames(coll) {
    var out = [];
    for (var i = 0; i < coll.length; i++) out.push(String(coll[i].name));
    return out;
}

var profile = null, source = "";
if (wanted !== null) {
    profile = findProfile(doc.preflightProfiles, wanted);
    source = "document (embedded)";
    if (!profile) { profile = findProfile(app.preflightProfiles, wanted); source = "application"; }
} else {
    var wp = doc.preflightOptions.preflightWorkingProfile;
    if (wp && typeof wp !== "string" && wp.isValid) {
        profile = wp;
        source = "document working profile";
    } else if (typeof wp === "string" && wp !== "") {
        profile = findProfile(doc.preflightProfiles, wp) || findProfile(app.preflightProfiles, wp);
        source = "document working profile (by name)";
    }
}
if (!profile) {
    throw new Error("Preflight profile not found: " + (wanted === null ? "(document working profile)" : wanted) +
        ". Document profiles: " + profileNames(doc.preflightProfiles).join(" | ") +
        ". Application profiles: " + profileNames(app.preflightProfiles).join(" | "));
}

var opts = doc.preflightOptions;
var modifiedBefore = doc.modified;
var t0 = new Date().getTime();
var proc = app.preflightProcesses.add(doc, profile);
// waitForProcess() returns false even after completion; completeness is read from processResults.
try { proc.waitForProcess($TIMEOUT$); } catch (eWait) {}
var elapsed = new Date().getTime() - t0;

var rows = [];
var readError = null;
try {
    var agg = proc.aggregatedResults;
    rows = (agg && agg[2]) ? agg[2] : [];
} catch (eAgg) { readError = String(eAgg.message || eAgg); }
var report = "";
try { report = String(proc.processResults).replace(/\\s+$/, ""); } catch (eRep) {}
try { proc.remove(); } catch (eRem) {}

// Rows are a pre-order tree: row[0] is the level; rows without page/info/details are
// category nodes ("TEXT (2)" > "Overset text (2)"), the others are the affected items.
var byCategory = {}, findings = [], path = [], total = 0;
for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var level = Number(row[0]) || 1;
    var label = String(row[1]);
    var info = String(row[3] || "");
    var page = String(row[2] || "");
    var hasDetails = row[4] && row[4].length;
    if (!page && !info && !hasDetails) {
        path.length = Math.min(path.length, Math.max(0, level - 1));
        path.push(label.replace(/\\s*\\(\\d+\\)$/, ""));
        continue;
    }
    total++;
    var category = path.slice(0, Math.max(0, level - 1)).join(" > ") || path.join(" > ");
    byCategory[category] = (byCategory[category] || 0) + 1;
    if (findings.length < maxFindings) {
        var details = [];
        if (hasDetails) {
            for (var j = 0; j < row[4].length; j++) {
                details.push(String(row[4][j][0]) + ": " + String(row[4][j][1]));
            }
        }
        findings.push({ category: category, item: label, page: page, details: details.length ? details : [info] });
    }
}
var stillRunning = /still looking/i.test(report);

__result = {
    document: { name: String(doc.name), path: doc.saved ? String(doc.fullName.fsName) : "(unsaved)" },
    profile: { name: String(profile.name), source: source },
    settings: {
        scope: String(opts.preflightScope),
        layers: String(opts.preflightWhichLayers),
        include_pasteboard: opts.preflightIncludeObjectsOnPasteboard,
        include_nonprinting: opts.preflightIncludeNonprintingObjects,
        live_preflight_off: opts.preflightOff
    },
    complete: !stillRunning && readError === null,
    read_error: readError,
    elapsed_ms: elapsed,
    findings_total: total,
    by_category: byCategory,
    findings: findings,
    findings_truncated: total > findings.length,
    process_results: report,
    document_modified_before: modifiedBefore,
    document_modified_after: doc.modified
};
"""


@mcp.tool()
def preflight_document(
    profile: str | None = None,
    document: str | None = None,
    max_findings: int = 100,
    timeout_s: int = 120,
) -> str:
    """Run a preflight profile on a document and report the findings. Read-only.

    Reports what actually happened, not what was requested: the profile used and
    where it came from (document working profile, embedded, application), whether
    the process finished within timeout_s (`complete`), the scope/layer settings
    the document applies, counts per category and the findings (category, item,
    page, details). `complete: false` means the findings may be incomplete.

    Args:
        profile: Profile name. Default: the document's working preflight profile.
        document: Document name, full path or id. Default: active document.
        max_findings: Maximum findings listed (counts per rule are always complete).
        timeout_s: Seconds to wait for the preflight process.
    """
    err = _check_document()
    if err:
        return _fmt({"success": False, "error": err})

    jsx = (
        _PREFLIGHT_JSX
        .replace("$DOC$", _doc_spec_literal(document))
        .replace("$PROFILE$", json.dumps(profile) if profile else "null")
        .replace("$MAX$", str(max(0, min(max_findings, 2000))))
        .replace("$TIMEOUT$", str(max(1, min(timeout_s, 900))))
    )
    try:
        result = com.run_jsx(jsx, undo_mode="none")
        return _fmt(_unwrap_result(result))
    except Exception as e:
        return _fmt({"success": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Tool 11: typography_audit
# ---------------------------------------------------------------------------

_TYPO_AUDIT_JSX = _RESOLVE_DOC_JSX + """\
var doc = __mcpResolveDoc($DOC$);
var maxParagraphs = $MAXPARA$;
var maxSamples = $SAMPLES$;
var includeTables = $TABLES$;

var fontStatus = {};
var docFonts = doc.fonts.everyItem().getElements();
for (var df = 0; df < docFonts.length; df++) {
    try { fontStatus[String(docFonts[df].name)] = String(docFonts[df].status); } catch (eF) {}
}

var noStyleId = doc.paragraphStyles[0].id;
var fonts = {}, sizes = {}, pstyles = {}, cstyles = {};
var unstyled = { paragraphs: 0, samples: [] };
var counted = { containers: 0, paragraphs: 0, characters: 0 };
var truncated = false;

function pageOf(textObj) {
    try {
        var f = textObj.parentTextFrames;
        if (f && f.length && f[0].parentPage) return String(f[0].parentPage.name);
    } catch (e) {}
    return null;
}
function sample(textObj) {
    var s = "";
    try { s = String(textObj.contents); } catch (e) {}
    s = s.replace(/[\\r\\n]+/g, " ");
    return { page: pageOf(textObj), preview: s.length > 60 ? s.substring(0, 60) : s };
}
function styleLabel(style) {
    var name = String(style.name);
    try {
        if (style.parent.constructor.name.indexOf("StyleGroup") > -1) name = String(style.parent.name) + ":" + name;
    } catch (e) {}
    return name;
}
function asArray(v, n) {
    if (v instanceof Array) return v;
    var out = [];
    for (var i = 0; i < n; i++) out.push(v);
    return out;
}

function auditText(txt) {
    var nRanges = txt.textStyleRanges.length;
    if (!nRanges) return;
    var nParas = txt.paragraphs.length;
    if (counted.paragraphs + nParas > maxParagraphs) { truncated = true; return; }
    counted.containers++;
    counted.paragraphs += nParas;

    var all = txt.textStyleRanges.everyItem();
    var contents = asArray(all.contents, nRanges);
    var rFonts = asArray(all.appliedFont, nRanges);
    var rSizes = asArray(all.pointSize, nRanges);
    var rLead = asArray(all.leading, nRanges);
    var rAuto = asArray(all.autoLeading, nRanges);
    var rCs = asArray(all.appliedCharacterStyle, nRanges);

    for (var i = 0; i < nRanges; i++) {
        var len = (typeof contents[i] === "string") ? contents[i].length : 1;
        counted.characters += len;

        var fname = "unresolved";
        try { fname = (typeof rFonts[i] === "string") ? rFonts[i] : String(rFonts[i].name); } catch (eN) {}
        var fr = fonts[fname];
        if (!fr) {
            var st = fontStatus[fname];
            if (st === undefined) {
                try { st = String(rFonts[i].status); } catch (eS) { st = "UNKNOWN"; }
            }
            var parts = fname.split("\\t");
            fr = fonts[fname] = { family: parts[0], style: parts[1] || "", status: st, characters: 0, ranges: 0, samples: [] };
        }
        fr.characters += len;
        fr.ranges++;
        if (fr.status !== "INSTALLED" && fr.samples.length < maxSamples) fr.samples.push(sample(txt.textStyleRanges[i]));

        var lead = (rLead[i] == Leading.AUTO) ? "auto " + rAuto[i] + "%" : rLead[i];
        var sk = rSizes[i] + "/" + lead;
        var sr = sizes[sk];
        if (!sr) sr = sizes[sk] = { size: rSizes[i], leading: lead, characters: 0 };
        sr.characters += len;

        var csName = "[None]";
        try { csName = styleLabel(rCs[i]); } catch (eC) {}
        cstyles[csName] = (cstyles[csName] || 0) + len;
    }

    if (!nParas) return;
    var paras = txt.paragraphs.everyItem();
    var pStyle = asArray(paras.appliedParagraphStyle, nParas);
    var pOver = asArray(paras.styleOverridden, nParas);
    for (var p = 0; p < nParas; p++) {
        var style = pStyle[p];
        var sid = String(style.id);
        var rec = pstyles[sid];
        if (!rec) rec = pstyles[sid] = { name: styleLabel(style), paragraphs: 0, overridden: 0, overridden_samples: [] };
        rec.paragraphs++;
        if (pOver[p]) {
            rec.overridden++;
            if (rec.overridden_samples.length < maxSamples) rec.overridden_samples.push(sample(txt.paragraphs[p]));
        }
        if (style.id === noStyleId) {
            unstyled.paragraphs++;
            if (unstyled.samples.length < maxSamples) unstyled.samples.push(sample(txt.paragraphs[p]));
        }
    }
}

var stories = doc.stories.everyItem().getElements();
var overset = { count: 0, samples: [] };
for (var s = 0; s < stories.length && !truncated; s++) {
    var story = stories[s];
    auditText(story);
    try {
        if (story.overflows) {
            overset.count++;
            if (overset.samples.length < maxSamples) {
                var tc = story.textContainers;
                var last = tc.length ? tc[tc.length - 1] : null;
                var pg = null;
                try { if (last && last.parentPage) pg = String(last.parentPage.name); } catch (eP) {}
                overset.samples.push({ page: pg, preview: String(story.contents).substring(0, 60).replace(/[\\r\\n]+/g, " ") });
            }
        }
    } catch (eO) {}
    try {
        var notes = story.footnotes.everyItem().getElements();
        for (var fn = 0; fn < notes.length && !truncated; fn++) {
            if (notes[fn].texts.length) auditText(notes[fn].texts[0]);
        }
    } catch (eFn) {}
    if (includeTables && story.tables.length) {
        var cells = story.tables.everyItem().cells.everyItem().getElements();
        for (var c = 0; c < cells.length && !truncated; c++) {
            if (cells[c].texts.length) auditText(cells[c].texts[0]);
        }
    }
}

function sortedList(map, key) {
    var out = [];
    for (var k in map) {
        if (map.hasOwnProperty(k)) {
            var v = map[k];
            if (key === "font") v.name = k;
            out.push(v);
        }
    }
    out.sort(function (a, b) { return (b.characters || b.paragraphs || 0) - (a.characters || a.paragraphs || 0); });
    return out;
}
var fontList = sortedList(fonts, "font");
var problemFonts = [];
for (var fi = 0; fi < fontList.length; fi++) {
    if (fontList[fi].status !== "INSTALLED") problemFonts.push(fontList[fi].name + " (" + fontList[fi].status + ")");
    if (fontList[fi].status === "INSTALLED") delete fontList[fi].samples;
}
var csList = [];
for (var cn in cstyles) if (cstyles.hasOwnProperty(cn)) csList.push({ name: cn, characters: cstyles[cn] });
csList.sort(function (a, b) { return b.characters - a.characters; });
var sizeList = sortedList(sizes, "size");

__result = {
    document: String(doc.name),
    scanned: counted,
    truncated: truncated,
    problem_fonts: problemFonts,
    fonts: fontList,
    sizes: sizeList.slice(0, 40),
    sizes_total: sizeList.length,
    paragraph_styles: sortedList(pstyles, "para"),
    character_styles: csList,
    unstyled_paragraphs: unstyled,
    overset_stories: overset
};
"""


@mcp.tool()
def typography_audit(
    document: str | None = None,
    include_tables: bool = True,
    max_paragraphs: int = 50000,
    samples: int = 3,
) -> str:
    """Audit the typography of a document without changing it.

    Reports fonts in use (with status and character counts; samples with page for
    fonts that are not INSTALLED), size/leading combinations, paragraph styles with
    paragraph counts and local overrides (samples with page), character style usage,
    paragraphs in [No Paragraph Style] and overset stories. Findings are not judged;
    deciding what is wrong is up to the operator. Covers all stories (pages, parent
    pages, pasteboard), their footnotes and optionally table cells.

    Args:
        document: Document name, full path or id. Default: active document.
        include_tables: Also scan table cell text (default True).
        max_paragraphs: Stop after this many paragraphs (`truncated: true`).
        samples: Samples per finding (page + text preview).
    """
    err = _check_document()
    if err:
        return _fmt({"success": False, "error": err})

    jsx = (
        _TYPO_AUDIT_JSX
        .replace("$DOC$", _doc_spec_literal(document))
        .replace("$MAXPARA$", str(max(1, max_paragraphs)))
        .replace("$SAMPLES$", str(max(0, min(samples, 20))))
        .replace("$TABLES$", "true" if include_tables else "false")
    )
    try:
        result = com.run_jsx(jsx, undo_mode="none", track_changes=False)
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
