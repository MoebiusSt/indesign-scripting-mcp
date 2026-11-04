# Agent Handover: InDesign Scripting MCP

## Project Purpose

Provide two MCP servers for Adobe InDesign automation:

- `server.py` (`indesign-dom`): read-only lookup over a local SQLite knowledge base.
- `exec_server.py` (`indesign-exec`): run JSX in a live InDesign instance via COM/OLE.

The DOM server covers **three sources** in one DB, all taken from Gregor Fellenz'
[Adobe ExtendScript API reference](https://www.indesignjs.de/indesignapi/) (indesignjs.de):

- `dom` (InDesign object model, `indesign/api.json`)
- `javascript` (ExtendScript core classes, `javascript/api.json`)
- `scriptui` (legacy ScriptUI classes, `scriptui/api.json`)

## Current Architecture (What Exists Now)

- Build pipeline: `api.json` (per target) + `uxp_notes.json` -> `parser.py` -> `extendscript.db`
- Query layer: `db.py`
- MCP layer: `server.py`
- CLI entrypoint: `manage.py`
- Runtime execution bridge: `indesign_com.py` + `json_polyfill.jsx`

Primary DB file:

- `extendscript.db` (default). `indesign_dom.db` is the legacy name; `build`/`update` refresh it too
  when it exists, because MCP client configs may still point to it via `INDESIGN_DOM_DB`.

## Key Decisions You Must Preserve

1. **Upstream = per-target `api.json` on indesignjs.de**
   - Not the OMV XML from the ExtendScript Toolkit cache (Adobe no longer ships it reliably), and not
     `indesignapi.zip` (19 MB of HTML) for the build. api.json is the structured format Gregor publishes
     for tooling; HTML is presentation markup.
   - Version check without download: top-level `llms.txt` (~2 KB) lists every target's version line and
     the generation date. `manage.py check` / `update` compare it with `db_meta.dom_title` / `dom_generated`.

2. **Single DB + source dimension**
   - No per-source DB split. Disambiguation is done by `source` (`dom|javascript|scriptui`).
   - Name collisions between `dom` and `scriptui`: Window, Button, Event, Events, Group, ListBox, Panel,
     RadioButton, StaticText. `classes` uniqueness is `(name, source_id)`; never suffix class names.

3. **Gaps in api.json are closed in `parser.py`, not by hand-editing the DB**
   - Static methods: api.json and the `.d.ts` files carry no static flag for methods. Core JavaScript and
     ScriptUI are frozen, so `STATIC_METHODS` is a fixed table (`File.openDialog`, `Folder.selectDialog`,
     `$.*`, `Math.*`, ...). Worth reporting upstream.
   - Constructors: method named like its class (non-dom) -> `element_type='constructor'`.
   - ScriptUI callbacks `on[A-Z]*` -> `element_type='event'`.
   - Reserved names: upstream escapes `Index` as `Index_`; `_canonical_name()` restores `Index` in class
     names and type strings; `doc_url` keeps the upstream slug.
   - Suites do not exist upstream anymore; `list_classes(type=...)` filters by kind.

4. **UXP data mirrors Gregor's ExtendScript/UXP toggle**
   - Rules in `parser.py`: enum classes -> `equals()` note; `[]` methods -> `es_only`; dom properties typed
     `File`/`Folder` (not arrays) -> async note; `File | Function` handlers -> handler note;
     `javascript`/`scriptui` -> source-level "ExtendScript only" note.
   - `uxp_notes.json`: hand-written notes that no rule can derive (e.g. `Application.scriptArgs`,
     `XMLElements.itemByName`).
   - `manage.py check-uxp` scans the rendered pages of `indesignapi.zip` and lists markers not covered by
     rules + overrides. Run it after upstream updates; add gaps to `uxp_notes.json`.

5. **Events have their own table** (`events`), listed by `lookup_class` and indexed in FTS as `event`.

6. **Atomic build**: the DB is written to `<db>.tmp` and swapped in with `os.replace()`.

## Fast Entry Points for Future Changes

- Fetch/parse/schema/build: `parser.py`
  - `fetch_sources()`, `fetch_remote_versions()`
  - `parse_api_json()`, `parse_sources()`, `_resolve_type_refs()`
  - `STATIC_METHODS`, `UXP_NOTE_*`, `_canonical_name()`
  - `DB_SCHEMA`, `build_database()`, `validate()`
  - `scan_uxp_markers()` (check only, never used by the build)

- Query behavior/API shape: `db.py`
  - All lookup functions are source-aware; `get_properties`/`get_methods`/`search_dom` accept `runtime`.

- MCP tool contract: `server.py`
  - Tool signatures include optional `source`; `runtime` is validated by `_check_runtime()`.
  - ScriptUI warning note injection for `source="scriptui"`.

- CLI workflows: `manage.py`
  - `check`, `fetch`, `build`, `update [--force]`, `check-uxp`, `validate [--dir]`
  - built-in regression checks after every build/validate.

## Agent Operating Pattern (Recommended)

1. `knowledge_overview()`
2. `list_sources()`
3. `search_dom(query, source=...)`
4. `lookup_class(...)` / `get_method_detail(...)`
5. For long-form reading: fetch `doc_url` (upstream Markdown page of the class)
6. switch to `indesign-exec` tools for runtime work (`run_jsx`, `eval_expression`, `undo`)

For ExtendScript-specific features, prefer `source="javascript"`:

- `$`, `UnitValue`, `File`, `Folder`, `Socket`, `XML`, `XMLList`, `RegExp`.

For UXP scripts pass `runtime="uxp"` and respect `uxp_note` / `extendscript_only`.

## Tool Surface (Compact)

`indesign-dom` MCP:

- Discovery: `knowledge_overview()`, `list_sources()`, `search_dom(query, source?, runtime?)`
- Class/API lookup: `lookup_class(name, source?)`, `get_properties(..., runtime?)`, `get_methods(..., runtime?)`, `get_method_detail(...)`
- Structure/meta: `get_hierarchy(class_name, source?)`, `get_enum_values(enum_name, source?)`, `list_classes(type?, source?, filter?)`, `dom_info()`
- Rule: for name collisions always set `source`.

`indesign-exec` MCP:

- Execution: `run_jsx(code, undo_name?, undo_mode?, language?)` — supports `language="javascript"` (ExtendScript ES3, default) and `language="uxp"` (modern JS ES6+/V8).
- Read-only checks: `get_document_info()`, `get_selection(detail_level?)`, `eval_expression(expression)`
- Recovery: `undo(steps?)`
- Learning loop: `report_learning(...)`, `get_gotchas(context?)`, `get_quick_reference()`

## Language Selection Policy (ExtendScript vs UXP)

- **Default: ExtendScript (`language="javascript"`)**: Use for 95% of tasks (layout mutations, typography, styles, find/change, table operations, File/Folder IO). Model world-knowledge is significantly richer and more reliable.
- **Selective: UXP (`language="uxp"`)**: Use when modern JavaScript features (Array `map`/`filter`/`reduce`, `Set`, `Map`, destructuring, top-level `await`, native JSON) simplify data preparation before layout.
- In UXP: Always import DOM via `const { app, ...enums } = require('indesign');` and assign return data to `__result`.
- Do NOT use `fetch()` inside InDesign UXP scripts (standalone COM scripts do not pump background socket loops; fetch HTTP data in Python/MCP instead).

## Community Gotchas Workflow

- Curated knowledge file: `gotchas.json` (committed, PR-friendly)
- Local submission queue: `submissions/pending.jsonl` (not committed)
- Submission tool: `report_learning(problem, solution, triggers, ...)`
- Review command: `python manage.py review-submissions`
- Promotion path: pending submission -> review -> `gotchas.json` -> commit/PR

## Known Pain Points and Special Handling

- **FTS and symbols**: the FTS5 tokenizer drops `$`; `search_dom` falls back to an exact name match for such queries.
- **ScriptUI is legacy**: keep documentation available, but suggest UXP for new UI-heavy work.
- **MCP reload requirement**: after tool signature or schema changes, restart the MCP server (the client keeps the old process).
- **COM/OLE runtime constraints**: Exec server is Windows-only and requires running InDesign.

## Validation Commands (Minimal)

```bash
python -m py_compile parser.py db.py server.py manage.py
python manage.py update --force
python manage.py validate --dir sources/indesignjs
python manage.py check-uxp
```

## File Map (Targeted)

- `parser.py`: api.json fetch + parse, UXP rules, schema, DB build/validate, UXP marker scan
- `uxp_notes.json`: curated UXP overrides
- `db.py`: all SQL query semantics for MCP tools
- `server.py`: MCP interface and tool descriptions
- `manage.py`: maintenance CLI and operational workflows
- `exec_server.py`: JSX execution MCP
- `indesign_com.py`: COM wrapper + safe execution envelope
- `README.md`: user-facing setup and usage
- `THIRD_PARTY_NOTICES.md`: attribution and licensing notes

## Agent must-know: toxic null hyperlinks

If an agent enumerates `doc.hyperlinks` and reads `.destination`, a **copied**
hyperlink with a missing destination (UI red flag, API `null`) may **open the
document it was pasted from** and switch `activeDocument`. This is not
`HyperlinkExternalPageDestination`. `run_jsx` / `eval_expression` detect
documents that appear during a script. Saved, unmodified ones are closed
(`close_unexpected_documents=true`, default). Modified or untitled ones stay
open. Whenever any appeared, the result has a `warning`: treat it as a stop
sign and verify the host before the next edit. Gotcha:
`toxic-null-hyperlink-opens-origin-document`. Scan/repair via the installed core
(load every `run_jsx` call — globals do not persist):

```js
$.evalFile(File(Folder.userData +
    "/Adobe/InDesign/Version 21.0/de_DE/Scripts/Scripts Panel/_lib/toxicHyperlinks.jsxinc"));
__result = $.global.THL.scan(app.documents.itemByName("file.indd"));
```

Orphan sources block "New hyperlink": `THL.findOrphanSources(doc)` / `THL.removeOrphanSources(doc, ids)`.

## Preserved Design Rationale (No external plan files required)

- Version-agnostic ingestion: rebuild/update always starts from upstream api.json, never from manual DB edits.
- The project stays single-DB and source-aware (`dom|javascript|scriptui`) to preserve cross-source lookup and avoid duplicate tooling.
- Name collisions are resolved via `source` filter, not class renaming.
- Exec and DOM servers stay separated by responsibility: read-only knowledge lookup vs side-effecting JSX execution.
- ScriptUI is intentionally retained for legacy script support, but modern UI recommendation is UXP for new development.
- InDesign Server (`indesign-server/api.json`, same format) can be added later as another target in `DEFAULT_TARGETS`.
- `README.md` is the operational reference; implementation truth is in `parser.py`, `db.py`, `server.py`, `manage.py`.
