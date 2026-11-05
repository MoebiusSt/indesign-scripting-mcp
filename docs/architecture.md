# Architecture, Data Sources & Build

[← Back to README](../README.md)

## Overview

```mermaid
flowchart LR
  DOM_JSON["indesign/api.json"] --> Parser[parser.py]
  JS_JSON["javascript/api.json"] --> Parser
  SUI_JSON["scriptui/api.json"] --> Parser
  MANIFEST["manifest.json"] --> Parser
  Parser --> SQLite[(extendscript.db)]
  SQLite --> DomServer["server.py (InDesign DOM MCP)"]
  Agent[Agent] --> DomServer
  Agent --> ExecServer["exec_server.py (InDesign Exec MCP)"]
  ExecServer --> COM[COM/OLE DoScript]
  COM --> InDesign["Adobe InDesign (running)\n[ExtendScript ES3 / UXP V8]"]
```

## Core Files

- `server.py`: MCP tools for querying the DOM database
- `exec_server.py`: MCP tools for script execution in InDesign (ExtendScript & UXP)
- `indesign_com.py`: COM/OLE automation layer + safe JSX/UXP wrapper + modal dialog watchdog
- `json_polyfill.jsx`: ExtendScript JSON polyfill + `__safeStringify()` (DOM-safe)
- `parser.py`: manifest/api.json fetcher + parser + database builder (SQLite + FTS5)
- `db.py`: database query layer for DOM tools
- `manage.py`: CLI to check/fetch/build/update/validate the database and review gotcha submissions
- `gotchas.json`: curated gotchas and best-practice knowledge base
- `check_docs.py`: verifies that README, docs and skills name only registered tools and parameters

## Data Sources

The database is built from the [Adobe ExtendScript API reference](https://www.indesignjs.de/indesignapi/) maintained by Gregor Fellenz (continuation of the Jongware reference by Theunis de Jong). It is generated from Adobe's OMV exports, so the OMV XML files from the ExtendScript Toolkit cache are no longer needed.

| Source key | Upstream file | Content |
|---|---|---|
| `dom` | `indesign/api.json` | InDesign object model (2026: 21.5.1.73) |
| `javascript` | `javascript/api.json` | ExtendScript Core JavaScript (`$`, `File`, `Folder`, `UnitValue`, `XML`, ...) |
| `scriptui` | `scriptui/api.json` | ScriptUI |

The parser requires api.json `formatVersion` 2 and takes everything from the data:

- **Static members**: `"static": true` on properties and methods (`File.openDialog`, `Window.alert`, `$.*`, `Math.*`) -> `element_type='class'`.
- **UXP runtime data** (`runtimes: ["extendscript", "uxp"]` on the InDesign target):
  - `"uxp": {"available": false}` (the `[]` accessors, `XMLElements.itemByName`) -> `extendscript_only`, hidden with `runtime="uxp"`;
  - `"uxp": {"type": ...}` / `"uxp": {"returns": ...}` (e.g. `Promise<File>` for File/Folder properties, `string` for `Link.filePath`) -> `uxp_type`;
  - `notes` per runtime -> `uxp_note`, `extendscript_note`, `note` (e.g. `equals()` on every enumeration, `scriptArgs`, `doScript`).
  - General UXP rules: `<target>/uxp.md`, exposed as `uxp_guide_url` in `list_sources`. UXP TypeScript types: [`indesign-uxp-types`](https://www.npmjs.com/package/indesign-uxp-types).
- **Names**: `name` is the original (`Index`, `$`); `doc_url` is built from `page` (`Index_.html` -> `Index_.md`).
- **Update check**: `manifest.json` lists per target `version`, `contentHash` (api.json without `generated`) and per file `sha256`. `fetch` verifies the sha256 of every api.json; `check`/`update` compare `contentHash` with the database.
- **Suites** no longer exist upstream; `list_classes` filters by kind instead.

## Build Commands

```bash
python manage.py check       # compare upstream manifest.json (contentHash) with the database
python manage.py update      # check + fetch api.json + diff + build (use --force to rebuild)
python manage.py fetch       # only download manifest.json and api.json files into sources/indesignjs/
python manage.py build       # only rebuild from the downloaded files
python manage.py validate --dir sources/indesignjs
python check_docs.py         # docs/skills name only registered tools and parameters
```

`build`/`update` write `extendscript.db` and, if present, the legacy `indesign_dom.db`. The database is written to a temporary file and swapped in at the end, so a running DOM server keeps working during the build. The downloaded files (`sources/`) and the database are generated locally and are not committed.

## Running the Servers Manually (stdio)

```bash
uv run server.py
uv run exec_server.py
```

or with plain Python after `pip install -r requirements.txt`:

```bash
python server.py
python exec_server.py
```
