# Tool Reference

[← Back to README](../README.md)

## InDesign DOM MCP (`server.py`)

Read-only lookups in the InDesign object model, ExtendScript Core JavaScript and ScriptUI.

- `lookup_class(name, source?)`: Class description, kind (`object` | `collection` | `enum`), superclass, subclasses, collection element type, events, UXP note and `doc_url` (upstream Markdown page).
- `get_properties(class_name, source?, filter?, include_inherited?, runtime?)`: Get properties of a class.
- `get_methods(class_name, source?, filter?, include_inherited?, runtime?)`: Get methods of a class; `runtime="uxp"` hides ExtendScript-only methods such as the `[]` index accessor.
- `get_method_detail(class_name, method_name, source?)`: Get detailed parameter types, ranges and return type for a method.
- `get_enum_values(enum_name, source?)`: Get all member names and numeric values for an enumeration.
- `get_hierarchy(class_name, source?)`: Get inheritance chain and direct subclasses.
- `search_dom(query, source?, runtime?)`: Full-text search across class, property, event, method and parameter names and descriptions.
- `list_classes(type=all, source?, filter?)`: List classes by kind (`object` | `collection` | `enum` | `class` | `all`) and name substring.
- `dom_info()`: Database metadata and statistics.
- `list_sources()`: List indexed knowledge sources (`dom`, `javascript`, `scriptui`).
- `knowledge_overview()`: High-level overview of the scripting knowledge base.

## InDesign Exec MCP (`exec_server.py`)

Runs scripts in the running InDesign instance via Windows COM/OLE.

- `run_jsx(code, undo_name="Agent Script", undo_mode="entire", language="javascript"|"uxp", debug=false, track_changes?)`: Execute ExtendScript (ES3) or UXP (ES6+/V8) code with undo grouping and safety wrapping. Errors report `line` relative to the submitted code plus `source_line`; syntax errors are located by a compile-only probe. `debug=true` adds surrounding source lines and the mapped UXP stack. Mutating calls return `changes` (added/removed page items, page/story deltas, objects passed to `__mcpTrack(obj, note)`).
- `get_document_info()`: Overview of the active document (pages, items, styles, active view/spread context).
- `get_selection(detail_level="basic"|"full")`: Information about the current selection in InDesign.
- `eval_expression(expression)`: Quickly evaluate a short read-only expression without undo overhead.
- `undo(steps=1)`: Undo the last operation(s) in the active document.
- `report_learning(problem, solution, triggers, category?, severity?, error_message?, jsx_context?)`: Submit a local learning entry to the pending queue.
- `get_gotchas(context?)`: Retrieve curated gotchas and pitfalls, ranked by relevance to a task context.
- `get_quick_reference()`: Comprehensive DOM cheatsheet with common access paths, collection patterns, and dynamic gotchas.
- `get_change_log(limit=20, document?, all_sessions=false)`: What mutating `run_jsx` calls changed, newest first. Persisted in `logs/change_log.jsonl`.
- `preflight_document(profile?, document?, max_findings=100, timeout_s=120)`: Run a preflight profile (default: the document's working profile). Reports the profile source, completeness, findings per category with page and details. Read-only.
- `typography_audit(document?, include_tables=true, max_paragraphs=50000, samples=3)`: Fonts with status, size/leading pairs, paragraph and character style usage, local overrides, unstyled paragraphs, overset stories. Samples carry page names. Read-only.
