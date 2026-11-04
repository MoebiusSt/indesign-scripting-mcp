# InDesign Scripting MCP for Windows

**Let Claude or Cursor work directly in Adobe InDesign. Describe the task, the agent writes and runs the script.**

<div align="center">
  <img src="docs/screenshot_indesign_mcp_first-test2.jpg" alt="Op-Art generated in InDesign via MCP" width="600">
  <br>
  <em>Done with a single prompt from within Claude Desktop in 20 seconds (and some tinkering afterwards).</em>
</div>

<br>

![Some chats with the InDesign agent](./docs/chat-example.svg)

## What can it do?

Simple jobs, sure: *"Run preflight and fix overset frames by adjusting word and letter spacing wiggle-room. Escalate difficult cases to me."*
*"Correct all captions to sit over their images instead of under them. Write a new auto-sizing objectStyle for this."*

But these are jobs it has actually done in production, on 80–100 page magazines:

* **"Apply the comments from the editor's article proof-read PDFs to the open issue."**
  It finds each passage, skips corrections already made and adds missing
  photo credits. Then it checks its own work by exporting and looking at
  the pages.
* **"Check all hyphenations before print."** It sorts the problems
  by severity, gives page, column and line for each, and finds the cause on ~2,000 hyphenations.
* **"Set the event calendar data from the website export and build the additional tour date boxes from the editor's .docx and loosely distribute the boxes over the calendar data columns."**
  Several hundred entries, house style, typo fixes, all in one undo step.
* **"Replace the 18-page section from the old issue with the editor's new package."**
  It relinks the files and keeps every other page's left/right position
  correct.
* **"Rebuild the table of contents after the page shifts from the new page plan."** Page numbers
  and teaser texts, which it learned to take from the section markers and article headlines in this workflow. Run it again and nothing breaks.
* **"Update the issue dates in all our media kits, German and English. Here are PDFs with annotations and corrections. Export new versioned PDFs."**

Every job follows the same loop: inspect → look up the DOM → run in a named undo step
→ verify → report. It never hides a problem.

## Built for Windows

Many InDesign MCP servers are built Mac-first (UXP plugin or AppleScript) and treat Windows as an add-on. This one is built and used in daily magazine production on Windows:

- **No plugin to install.** It talks to InDesign through the native Windows COM interface. Start InDesign, start the agent, done.
- **One undo step per action.** Each script is grouped into a single, named undo step.
- **Protects your documents.** If a script unexpectedly opens another document (a known InDesign hyperlink pitfall), the agent notices, closes it if safe and warns you.
- **Knows InDesign.** A local database of the complete InDesign scripting API (incl. ExtendScript and ScriptUI) and a growing collection of known pitfalls ("gotchas") keep the agent from guessing.
- **Free-form scripts instead of hundreds of fixed tools.** A small set of tools keeps the agent's context lean; anything InDesign can script is possible.
- **Two script engines.** Classic ExtendScript (default) and modern UXP JavaScript.

## New in November 2026

- **Precise error messages:** script errors point to the exact line of the agent's code; syntax errors are found without running anything. `debug=true` shows the surrounding lines.
- **Change log:** every changing action records which objects were added, removed or edited (`get_change_log`).
- **Typography audit:** fonts with status, sizes and leading, style usage, local overrides, overset text. Read-only, useful for issue sign-off (`typography_audit`).
- **Preflight check:** runs the document's preflight profile and reports the findings per page, and whether the check actually finished (`preflight_document`).
- **Docs check:** `check_docs.py` makes sure the README, docs and agent skills only mention tools that exist.
- **API knowledge from indesignjs.de**, including notes on where UXP behaves differently from ExtendScript.

## Quick Start

### 1. Requirements

- Windows 10/11 with Adobe InDesign (2025 or 2026)
- [Python 3.11+](https://www.python.org/downloads/) and [uv](https://docs.astral.sh/uv/getting-started/installation/)
- Claude Desktop or Cursor

### 2. Download and build the API database

```bash
git clone https://github.com/MoebiusSt/indesign-scripting-mcp.git
cd indesign-scripting-mcp
python manage.py update
```

### 3. Connect your agent

Add both servers to `%APPDATA%\Claude\claude_desktop_config.json` (Claude Desktop) or to the MCP settings in Cursor. Replace the path with your folder:

```json
{
  "mcpServers": {
    "indesign-dom": {
      "command": "uv",
      "args": ["run", "--script", "C:\\path\\to\\indesign-scripting-mcp\\server.py"]
    },
    "indesign-exec": {
      "command": "uv",
      "args": ["run", "--script", "C:\\path\\to\\indesign-scripting-mcp\\exec_server.py"]
    }
  }
}
```

Restart Claude Desktop or Cursor, open a document in InDesign and start asking.

### 4. Optional: install the agent skill

The skill teaches the agent a safe routine (inspect → look up → run → verify → undo if needed):

- **Cursor:** `.cursor/skills/indesign-mcp-operator/SKILL.md` is picked up automatically.
- **Claude:** upload `claude-skills/indesign-mcp-operator/indesign-mcp-operator.skill`.

## Learn More

- [Tool reference](docs/tools.md): all tools of both servers
- [Agent guide](docs/agent-guide.md): script engines, undo, error reporting, change log, best practices
- [Architecture, data sources & build](docs/architecture.md): how the API database is built and updated
- [Contributing gotchas](docs/contributing-gotchas.md): share pitfalls and fixes with other users

## License

MIT. See `LICENSE`.

## Third-Party Notices

Safety and serialization patterns are inspired by IdExtenso (Marc Autret, MIT). See `THIRD_PARTY_NOTICES.md`.

## Acknowledgements

- Gregor Fellenz: the [Adobe ExtendScript API reference](https://www.indesignjs.de/indesignapi/) on indesignjs.de is the data source of the DOM server (descriptive texts by Adobe), and [extendscriptApiDocTransformations](https://github.com/grefel/extendscriptApiDocTransformations) provided conceptual learnings for the former OMV XML pipeline.
- Theunis de Jong (Jongware): the original JavaScript reference this work continues.
- Andrea Cacioppo ([nutriandrea/adobe-indesign-mcp](https://github.com/nutriandrea/adobe-indesign-mcp)): the ideas for line-exact script errors, the read-only typography audit, honest preflight reporting, change tracking and the docs check that only allows existing tool names.
