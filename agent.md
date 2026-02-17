# Agent Handover: InDesign Scripting MCP

## Project Purpose

Provide two MCP servers for Adobe InDesign automation:

- `server.py` (`indesign-dom`): read-only lookup over a local SQLite knowledge base.
- `exec_server.py` (`indesign-exec`): run JSX in a live InDesign instance via COM/OLE.

The DOM server now covers **three sources** in one DB:

- `dom` (InDesign OMV)
- `javascript` (ExtendScript core classes)
- `scriptui` (legacy ScriptUI classes)

## Current Architecture (What Exists Now)

- Build pipeline: XML -> `parser.py` -> `extendscript.db`
- Query layer: `db.py`
- MCP layer: `server.py`
- CLI entrypoint: `manage.py`
- Runtime execution bridge: `indesign_com.py` + `json_polyfill.jsx`

Primary DB file:

- `extendscript.db` (new default; `indesign_dom.db` is legacy)

## Key Decisions You Must Preserve

1. **Single DB + source dimension**
   - No per-source DB split.
   - Disambiguation is done by `source` (`dom|javascript|scriptui`).

2. **Name collision handling via schema, not renaming**
   - `classes` uniqueness is `(name, source_id)`.
   - Do not suffix class names (`Window (SUI)` style was intentionally avoided).

3. **Namespace normalization first**
   - Parser strips XML namespaces before normal element traversal.

4. **Description model**
   - `shortdesc` + `description` are both parsed and merged.

5. **Type details retained**
   - Multiple sibling `<datatype>` values are preserved (pipe-joined union style).
   - `type@href` references are stored (`*_type_ref` fields).

## Fast Entry Points for Future Changes

- Schema/build logic: `parser.py`
  - `DB_SCHEMA`
  - `parse_xml()`
  - `parse_sources()`
  - `_parse_datatypes()`
  - `build_database()`
  - `validate()`

- Query behavior/API shape: `db.py`
  - All lookup functions are source-aware.
  - New capability endpoints: `list_sources()`, `knowledge_overview()`.

- MCP tool contract: `server.py`
  - Tool signatures include optional `source`.
  - Added tools: `list_sources`, `knowledge_overview`.
  - ScriptUI warning note injection for `source="scriptui"`.

- CLI workflows: `manage.py`
  - `build-all --dom --js --sui`
  - `validate --expect-sources ...`
  - built-in regression smoke checks.

## Agent Operating Pattern (Recommended)

1. `knowledge_overview()`
2. `list_sources()`
3. `search_dom(query, source=...)`
4. `lookup_class(...)` / `get_method_detail(...)`
5. switch to `indesign-exec` tools for runtime work (`run_jsx`, `eval_expression`, `undo`)

When class names are ambiguous, **always pass `source`**:

- `Window`, `Group`, `Panel`, `Event` collide between `dom` and `scriptui`.

For ExtendScript-specific features, prefer `source="javascript"`:

- `$`, `UnitValue`, `File`, `Folder`, `Socket`, `XML`, `XMLList`, `RegExp`.

## Tool Surface (Compact)

`indesign-dom` MCP:

- Discovery: `knowledge_overview()`, `list_sources()`, `search_dom(query, source?)`
- Class/API lookup: `lookup_class(name, source?)`, `get_properties(...)`, `get_methods(...)`, `get_method_detail(...)`
- Structure/meta: `get_hierarchy(class_name, source?)`, `get_enum_values(enum_name, source?)`, `list_classes(suite?, type?, source?)`, `dom_info()`
- Rule: for collisions (`Window`, `Group`, `Panel`, `Event`) always set `source`.

`indesign-exec` MCP:

- Execution: `run_jsx(code, undo_name?, undo_mode?)`
- Read-only checks: `get_document_info()`, `get_selection(detail_level?)`, `eval_expression(expression)`
- Recovery: `undo(steps?)`

## Known Pain Points and Special Handling

- **PowerShell and `$` in filenames**: use quoting when passing OMV paths in CLI (`'omv$indesign-...xml'`).
- **ScriptUI is legacy**: keep documentation available, but suggest UXP for new UI-heavy work.
- **MCP reload requirement**: after tool signature changes, client-side MCP reload/restart may be needed.
- **COM/OLE runtime constraints**: Exec server is Windows-only and requires running InDesign.

## Validation Commands (Minimal)

```bash
python -m py_compile parser.py db.py server.py manage.py
python manage.py build-all --dom "sources/omv$indesign-21.064$21.0.xml" --js "sources/javascript.xml" --sui "sources/scriptui.xml"
python manage.py validate --expect-sources dom,javascript,scriptui
```

## File Map (Targeted)

- `parser.py`: XML ingestion, normalization, schema, DB build/validate
- `db.py`: all SQL query semantics for MCP tools
- `server.py`: MCP interface and tool descriptions
- `manage.py`: maintenance CLI and operational workflows
- `exec_server.py`: JSX execution MCP
- `indesign_com.py`: COM wrapper + safe execution envelope
- `README.md`: user-facing setup and usage
- `THIRD_PARTY_NOTICES.md`: attribution and licensing notes

## Preserved Design Rationale (No external plan files required)

- The original project goal was version-agnostic OMV ingestion, so rebuild/update always starts from XML input and not from manual schema edits.
- The project stays single-DB and source-aware (`dom|javascript|scriptui`) to preserve cross-source lookup and avoid duplicate tooling.
- Name collisions are resolved via `source` filter, not class renaming.
- Parser priority order is fixed: namespace normalization -> structural parse -> type/reference preservation.
- Exec and DOM servers stay separated by responsibility: read-only knowledge lookup vs side-effecting JSX execution.
- ScriptUI is intentionally retained for legacy script support, but modern UI recommendation is UXP for new development.
- `README.md` is the operational reference; implementation truth is in `parser.py`, `db.py`, `server.py`, `manage.py`.
