# InDesign MCP Operator Reference

## Suggested Task Prompt Template
Use this prompt skeleton when applying the skill:

```text
Task: <short objective>

Constraints:
- Use indesign-dom for API lookup before JSX.
- Use get_gotchas(context) before non-trivial execution.
- Use undo_mode="entire" for mutations and set undo_name.
- Verify with read-only checks after mutation.
- Cleanup all _tmp_ Script Labels at end.
```

## Example Context Strings for `get_gotchas`
- `modeless palette event listener targetengine`
- `find change grep NothingEnum preferences`
- `script labels _tmp_ cleanup agentContext_ registry`
- `allPageItems spread corrupted fallback pageItems`
- `paragraph style itemByName grouped styles allParagraphStyles`

## Short JSX Patterns

### Read-only verification
```javascript
var doc = app.activeDocument;
__result = {
  name: doc.name,
  pages: doc.pages.length,
  selection: app.selection.length
};
```

### Mutation with explicit target check
```javascript
var doc = app.activeDocument;
var tf = doc.pages[0].textFrames.add();
tf.geometricBounds = [20, 20, 60, 120];
tf.contents = "Agent test";
__result = {id: tf.id, ok: tf.contents === "Agent test"};
```

### Temporary label lifecycle
```javascript
var doc = app.activeDocument;
doc.insertLabel("_tmp_batchIds", "1,2,3");
var raw = doc.extractLabel("_tmp_batchIds");
var ids = raw ? raw.split(",") : [];
// ... use ids ...
doc.insertLabel("_tmp_batchIds", "");
__result = {processed: ids.length, cleaned: doc.extractLabel("_tmp_batchIds") === ""};
```

## Recommended Output Checklist
At the end of each non-trivial task, report:

1. Executed undo label.
2. Verification result.
3. Whether a new gotcha was submitted.
