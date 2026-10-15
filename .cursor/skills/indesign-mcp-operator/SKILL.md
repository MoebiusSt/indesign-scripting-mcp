---
name: indesign-mcp-operator
description: Orchestrates InDesign automation via MCP tools with a strict inspect-lookup-execute-verify-rollback workflow. Use when tasks involve Adobe InDesign, ExtendScript/JSX, DOM lookup, Script Labels, undo safety, or gotcha-aware execution.
---

# InDesign MCP Operator

## Goal
Execute InDesign tasks safely and repeatably with the `indesign-dom` and `indesign-exec` MCP servers.

## Prerequisites (Cursor)
- MCP servers `indesign-dom` and `indesign-exec` must be enabled in Cursor Settings → MCP.
- InDesign Desktop must be running (Exec server uses Windows COM/OLE).
- Read each tool schema under the MCP descriptors before calling tools.

## MCP Tool Access (Cursor)
| Purpose | MCP server | Example tools |
|---------|------------|---------------|
| DOM lookup | `indesign-dom` | `knowledge_overview`, `search_dom`, `lookup_class`, `get_method_detail` |
| JSX execution | `indesign-exec` | `get_quick_reference`, `get_gotchas`, `run_jsx`, `eval_expression`, `get_document_info`, `undo` |

Server names may appear with a scope prefix (e.g. `user-indesign-dom`) depending on whether they were added at user or project level. Match by the `indesign-dom`/`indesign-exec` suffix, not by exact string.

For ambiguous class names (`Window`, `Group`, `Panel`, `Event`), pass `source` (`dom`, `javascript`, or `scriptui`).

## Workflow
1. Inspect current state.
2. Look up API facts.
3. Run JSX with explicit undo policy.
4. Verify outcomes.
5. Roll back if needed.
6. Capture new learnings.

## Required Sequence
For non-trivial tasks, run these steps in order:

1. `get_quick_reference()` (exec server)
2. `get_gotchas(context)` (exec server)
3. DOM lookup (`search_dom`, `lookup_class`, `get_method_detail`, ...) on dom server
4. `run_jsx(...)` with:
   - `undo_mode="entire"` for any mutation
   - a descriptive `undo_name` (`Agent: <task>`)
5. Verification (`eval_expression`, `get_document_info`, `get_selection`)
6. If result is wrong: run some targeted corrective attempts first; use `undo(steps=1)` immediately only when document state is uncertain, partially corrupted, or broad rollback is safer than incremental fixes.

For strictly read-only checks, use `undo_mode="none"`.

## Script Labels Policy
- Temporary keys must start with `_tmp_`.
- Persistent keys must start with `agentContext_`.
- Always clear temporary keys at task end:
  - `doc.insertLabel("_tmp_key", "")`
- Keep persistent labels only when they have proven cross-session value.
- Optionally maintain `_agentLabelRegistry` with all persistent keys.

## Script Engine & Language Selection
`run_jsx` supports two engines via the `language` parameter:

| Engine | Parameter | JavaScript Standard | Best For |
|---|---|---|---|
| **ExtendScript (Default)** | `language="javascript"` | ES3 | 95% of tasks: layout, typography, styles, find/change, tables, File/Folder IO. Highest model world-knowledge & accuracy. |
| **UXP Scripting** | `language="uxp"` | ES6+ / V8 | Complex data transformation (`map`/`filter`/`reduce`, `Set`, `Map`, destructuring, top-level `await` / `Promise.all`), native JSON. |

### Decision Rule
- **DEFAULT to ExtendScript (`language="javascript"`)**: Do NOT assume modern is automatically better. ExtendScript has decades of training data and maximum reliability for InDesign DOM operations.
- **USE UXP (`language="uxp"`) selectively**: Only when modern JS features significantly simplify data preparation.
- **Do NOT use `fetch()` in InDesign scripts**: Standalone UXP scripts do not pump background socket loops. Fetch external HTTP/API data in Python/MCP and pass the data payload into InDesign.

## ExtendScript (JSX) Rules (`language="javascript"`)
- Use ES3 syntax only (`var`, classic functions, no arrow functions/const/let).
- Assign outputs to `__result`; do not use `return`.
- Prefer collection APIs over loops:
  - `everyItem()`, `itemByName()`, `itemByID()`, `itemByRange()`.
- Clear find/change prefs before usage:
  - `app.findTextPreferences = NothingEnum.NOTHING`
  - `app.changeTextPreferences = NothingEnum.NOTHING`

## UXP Script Rules (`language="uxp"`)
- Must import InDesign DOM: `const { app, ...enums } = require('indesign');` (or `require('indesign').app`).
- Enums are exported by `require('indesign')` (e.g. `const { Justification, UndoModes } = require('indesign');`).
- Top-level `await` is natively supported inside user code.
- Assign outputs to `__result` (the MCP wrapper handles serialisation via labels).
- All runtime exceptions and engine syntax errors are safely caught and auto-dismissed by the MCP wrapper.

## Safety Gates
Before mutation:
- Confirm a document is open and correct.
- Confirm key assumptions with read-only checks.
- Address the host by `fullName.fsName` / `itemByName`, not by assuming `activeDocument` stays put.

After mutation:
- Verify the specific target state changed as intended.
- If partial failure is detected, undo first, then retry with narrower scope.
- If a `run_jsx` result contains `warning`, a foreign document appeared mid-script (often a toxic null hyperlink). Stop, re-address the host by name, check `unexpected_documents`, then continue.

## Toxic null hyperlinks
Copied hyperlinks with `destination === null` can open the origin `.indd` when `.destination` is read. `run_jsx` / `eval_expression` detect documents that appear during a script. Saved, unmodified ones are closed (`close_unexpected_documents=true`, default). Modified or untitled ones stay open. Whenever any appeared, the result has a `warning`: treat it as a stop sign and verify the host before the next edit. Still never trust `activeDocument` after hyperlink scans. Load the core every call: `$.evalFile(File(Folder.userData + "/Adobe/InDesign/Version 21.0/de_DE/Scripts/Scripts Panel/_lib/toxicHyperlinks.jsxinc"));` then `THL.scan(doc)` / `THL.sanitize(...)`. Gotcha: `toxic-null-hyperlink-opens-origin-document`.

## Learning Loop
When a user-reported issue is resolved with clear root cause:
1. Check for equivalent gotcha first.
2. Submit via `report_learning(...)`.
3. Keep entries concise and actionable.

## Do Not
- Do not persist temporary `_tmp_` labels beyond the task.
- Do not assume modern JavaScript syntax.
- Do not mix inspection and destructive edits in one large opaque JSX block when the operation is risky.
- Do not set `close_unexpected_documents=false` unless the script intentionally opens another document that must stay open.
- Do not invent hyperlink destination URLs from visible link text.

## Response Contract
When reporting task completion:
- State what was verified.
- State whether undo is possible and what label was used.
- State whether temporary labels were cleaned up.

## Additional Reference
See [reference.md](reference.md) for reusable prompts and short templates.
