---
name: indesign-mcp-operator
description: Orchestrates InDesign automation via MCP tools with a strict inspect-lookup-execute-verify-rollback workflow. Use when tasks involve Adobe InDesign, ExtendScript/JSX, DOM lookup, Script Labels, undo safety, or gotcha-aware execution.
---

# InDesign MCP Operator

## Goal
Execute InDesign tasks safely and repeatably with `indesign-dom` and `indesign-exec` MCP servers.

## Workflow
1. Inspect current state.
2. Look up API facts.
3. Run JSX with explicit undo policy.
4. Verify outcomes.
5. Roll back if needed.
6. Capture new learnings.

## Required Sequence
For non-trivial tasks, run these steps in order:

1. `get_quick_reference()`
2. `get_gotchas(context)`
3. DOM lookup (`search_dom`, `lookup_class`, `get_method_detail`, ...)
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

## JSX Rules
- Use ES3 syntax only (`var`, classic functions).
- Assign outputs to `__result`; do not use `return`.
- Prefer collection APIs over loops:
  - `everyItem()`, `itemByName()`, `itemByID()`, `itemByRange()`.
- Clear find/change prefs before usage:
  - `app.findTextPreferences = NothingEnum.NOTHING`
  - `app.changeTextPreferences = NothingEnum.NOTHING`

## Safety Gates
Before mutation:
- Confirm a document is open and correct.
- Confirm key assumptions with read-only checks.

After mutation:
- Verify the specific target state changed as intended.
- If partial failure is detected, undo first, then retry with narrower scope.

## Learning Loop
When a user-reported issue is resolved with clear root cause:
1. Check for equivalent gotcha first.
2. Submit via `report_learning(...)`.
3. Keep entries concise and actionable.

## Do Not
- Do not persist temporary `_tmp_` labels beyond the task.
- Do not assume modern JavaScript syntax.
- Do not mix inspection and destructive edits in one large opaque JSX block when the operation is risky.

## Response Contract
When reporting task completion:
- State what was verified.
- State whether undo is possible and what label was used.
- State whether temporary labels were cleaned up.

## Additional Reference
See [reference.md](reference.md) for reusable prompts and short templates.
