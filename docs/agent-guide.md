# Agent Guide: How Scripts Run

[← Back to README](../README.md)

This page is for people who want to understand or tune how the agent writes and runs InDesign scripts.

## Workflow: lookup → code → execute → verify

1. Use the DOM MCP to look up the correct class/method signature.
2. For ambiguous class names (`Window`, `Group`, `Panel`, `Event`) pass `source="dom"` or `source="scriptui"`.
3. For ExtendScript-specific APIs (`UnitValue`, `$`, `File`, `Folder`, `Socket`, `XML`, `XMLList`) use `source="javascript"`.
4. Write script code that assigns a plain value to `__result` (do not `return`).
5. Execute with `run_jsx` (defaulting to `language="javascript"`, or `language="uxp"` for ES6+ data transformations).
6. Verify with `eval_expression` or `get_document_info`.
7. If needed, roll back with `undo`.

## Script Engines & Language Selection

`run_jsx` supports two script engines via the `language` parameter:

| Aspect | ExtendScript (`language="javascript"`) | UXP Scripting (`language="uxp"`) |
|---|---|---|
| **Role** | **Default engine (recommended for 95% of tasks)** | **Selective engine (for complex ES6+ data ops)** |
| **JS Standard** | ES3 + Adobe Extensions | Modern ES6+ / V8 Engine |
| **Model World Knowledge** | Massive (20+ years of community scripts/docs) | Limited (fewer public examples) |
| **Best For** | Layout, typography, styles, find/change, tables, File/Folder IO | Array `map`/`filter`/`reduce`, `Set`/`Map`, destructuring, top-level `await`, native JSON |
| **DOM Import** | Global `app` automatically available | `const { app } = require('indesign');` required |
| **Execution Bridge** | Windows COM/OLE `DoScript` (synchronous) | Windows COM/OLE `DoScript` (`ScriptLanguage.UXPSCRIPT`) |
| **Setup** | Zero-setup (start InDesign, run MCP) | Zero-setup (start InDesign, run MCP) |

Policy for agents:

- **DEFAULT to ExtendScript (`language="javascript"`)**: Decades of documentation and community solutions give LLMs unmatched accuracy for InDesign DOM operations.
- **USE UXP (`language="uxp"`) selectively**: When modern JS data manipulation (`map`, `filter`, `reduce`, `Set`, `Map`, destructuring, async microtasks) significantly simplifies the task before layout placement.

## The `__result` Convention

The execution wrapper intercepts and serializes `__result` to JSON.

**ExtendScript (`language="javascript"`, default):**
```jsx
var doc = app.activeDocument;
__result = { name: doc.name, pages: doc.pages.length };
```

**UXP Scripting (`language="uxp"`):**
```javascript
const { app, Justification } = require('indesign');
const doc = app.activeDocument;
const titles = doc.stories.everyItem().paragraphs.everyItem().contents
  .map(text => text.trim())
  .filter(text => text.length > 0);

__result = { totalParagraphs: titles.length, sample: titles.slice(0, 5) };
```

## Undo Grouping

Use `undo_mode="entire"` (default) and a descriptive `undo_name` so one Ctrl+Z reverts the whole operation cleanly. Use `undo_mode="none"` for read-only scripts.

## Errors and Debugging

- Error results carry `line` relative to the submitted code (1-based) and `source_line`.
- Syntax errors are located by a compile-only probe (`source: "syntax check (nothing was executed)"`).
- `debug=true` adds `context` (±3 source lines) and, for UXP, the stack mapped to the submitted lines.

## Change Log

- Mutating calls (`undo_mode` "entire"/"auto", ExtendScript) return `changes`: host document, added page items (id/type/page), removed ids, page/story deltas. This also works when the script fails halfway.
- Property edits are not diffed. Record them with `__mcpTrack(obj, "note")`; it returns `obj`:

  ```jsx
  __mcpTrack(tf, "contents").contents = "New text";
  ```

- `get_change_log()` lists the entries of the current session (`all_sessions=true` for older ones, stored in `logs/change_log.jsonl`).

## Foreign Documents and Toxic Hyperlinks

Copied hyperlinks with `destination === null` can open the origin `.indd` when `.destination` is read. `run_jsx` and `eval_expression` pin the open documents before each script and detect documents that appear during it. Saved, unmodified ones are closed again (`close_unexpected_documents=true`, default); modified or untitled ones stay open and the result carries a `warning`.

## Collection Patterns (prefer over loops)

InDesign collections support collective specifiers via `everyItem()` and direct accessors like `itemByName()`, `itemByID()`, `itemByRange()`. These execute as **one command** in InDesign rather than N round-trips:

```javascript
// Bulk property write — single operation
doc.textFrames.everyItem().fillColor = "None";

// Read all page names at once — returns Array
__result = doc.pages.everyItem().name;

// Direct accessor by name — no loop needed
var headingStyle = doc.paragraphStyles.itemByName("Heading 1");
```

## UXP Notes & Constraints

- **Imports**: In UXP, always import DOM classes and enums with `const { app, ...enums } = require('indesign');`.
- **Top-level `await`**: User code is automatically wrapped in an `async` IIFE. Top-level `await Promise.all(...)` and async helper functions work natively.
- **Modal dialog protection**: Unhandled runtime exceptions and syntax errors are caught by the MCP layer. InDesign modal dialogs are auto-dismissed to prevent COM lockups.
- **Timers & sockets**: Standalone InDesign scripts executed via COM do not pump main-thread macrotasks for `setTimeout()` or `fetch()`. Perform network requests in the agent/Python environment and pass the data into InDesign.
