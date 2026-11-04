# Contributing Gotchas

[← Back to README](../README.md)

The gotchas system is a shared debugging memory for this MCP:

- `gotchas.json` is the curated, versioned knowledge base committed to Git.
- Entries describe recurring pitfalls, fixes, and trigger keywords for context matching.
- Agents query relevant items with `get_gotchas(context, min_severity?, top_n?)` before writing scripts.
- `get_quick_reference()` includes static guidance and appends community gotchas at runtime.
- During troubleshooting, agents submit newly resolved issues via `report_learning(...)`.

## A) Direct curated updates via GitHub PR

1. Fork + clone this repository
2. Edit `gotchas.json` and add a new entry
3. Open a PR with the new gotcha
4. Maintainer review and merge

## B) Agent submission → maintainer review

1. Your agent calls `report_learning(...)`
2. The entry is written to `submissions/pending.jsonl` (local queue, not committed)
3. The maintainer runs:

   ```bash
   python manage.py review-submissions
   ```

4. Approved entries are promoted to `gotchas.json`
5. Commit and open a PR to share the curated learning
