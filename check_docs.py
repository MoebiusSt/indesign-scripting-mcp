"""
Check that documentation only names MCP tools (and tool parameters) that exist.

Registered tools are read from the ``@mcp.tool()`` functions in exec_server.py
and server.py (via AST, no imports). Documents are then scanned for:

- snake_case calls ``name(...)``: in skill files every such call must be a
  registered tool; in other docs only near misses of a tool name are reported.
- arguments of tool calls (``run_jsx(undo_mode="none")``, ``get_gotchas(context?)``):
  every named argument must be a parameter of that tool.
- backticked snake_case tokens that are near misses of a tool name.
- "Provides N tools" in each server docstring must match the registry.

A document can exempt tokens it quotes on purpose:
    <!-- tool-names: ignore old_tool_name, other_name -->

Usage:
    python check_docs.py [extra.md ...]
Exit code 1 when findings exist.
"""

import ast
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
SERVER_FILES = [BASE_DIR / "exec_server.py", BASE_DIR / "server.py"]
DEFAULT_DOCS = [
    "README.md",
    "agent.md",
    "docs/*.md",
    "claude-skills/**/*.md",
    ".cursor/skills/**/*.md",
]
STRICT_NAMES = {"SKILL.md", "reference.md"}
NEAR_MISS_MAX = 2

CALL_RE = re.compile(r"(?<![.\w$])([a-z][a-z0-9]*(?:_[a-z0-9]+)+)\(([^()\n]*)\)")
BACKTICK_RE = re.compile(r"`([a-z][a-z0-9]*(?:_[a-z0-9]+)+)`")
IGNORE_RE = re.compile(r"<!--\s*tool-names:\s*ignore\s+([^>]*?)\s*-->")
ARG_NAME_RE = re.compile(r"^\s*([a-z_][a-z0-9_]*)\s*(?:\?|=|$)")
COUNT_RE = re.compile(r"Provides (\d+) tools")


def load_tools() -> dict[str, set[str]]:
    """Return {tool_name: {param, ...}} for all @mcp.tool() functions."""
    tools: dict[str, set[str]] = {}
    for path in SERVER_FILES:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in tree.body:
            if isinstance(node, ast.FunctionDef) and any(_is_tool_decorator(d) for d in node.decorator_list):
                params = {a.arg for a in node.args.args + node.args.kwonlyargs}
                tools[node.name] = params
    return tools


def _is_tool_decorator(dec: ast.expr) -> bool:
    target = dec.func if isinstance(dec, ast.Call) else dec
    return isinstance(target, ast.Attribute) and target.attr == "tool"


def check_tool_counts() -> list[str]:
    """Compare 'Provides N tools' in each server docstring with its registry."""
    findings = []
    for path in SERVER_FILES:
        text = path.read_text(encoding="utf-8")
        tree = ast.parse(text)
        count = sum(
            1 for node in tree.body
            if isinstance(node, ast.FunctionDef) and any(_is_tool_decorator(d) for d in node.decorator_list)
        )
        doc = ast.get_docstring(tree) or ""
        m = COUNT_RE.search(doc)
        if m and int(m.group(1)) != count:
            findings.append(f"{path.name}: docstring says {m.group(1)} tools, {count} are registered")
    return findings


def levenshtein(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def closest(token: str, names) -> tuple[str | None, int]:
    best, dist = None, 10**6
    for name in names:
        d = levenshtein(token, name)
        if d < dist:
            best, dist = name, d
    return best, dist


def review(path: Path, text: str, tools: dict[str, set[str]]) -> list[str]:
    """Return findings for one document."""
    strict = path.name in STRICT_NAMES
    ignored = set()
    for m in IGNORE_RE.finditer(text):
        ignored.update(t.strip() for t in m.group(1).split(",") if t.strip())
    all_params = set().union(*tools.values()) if tools else set()
    findings = []
    seen = set()

    for lineno, line in enumerate(text.splitlines(), 1):
        for m in CALL_RE.finditer(line):
            name, args = m.group(1), m.group(2)
            if name in ignored:
                continue
            if name in tools:
                for raw in _split_args(args):
                    am = ARG_NAME_RE.match(raw)
                    if not am or raw.strip().startswith(("'", '"')):
                        continue
                    arg = am.group(1)
                    if arg not in tools[name] and arg not in ignored:
                        findings.append(f"{path}:{lineno}: `{name}` has no parameter `{arg}`")
                continue
            best, dist = closest(name, tools)
            if dist <= NEAR_MISS_MAX:
                findings.append(f"{path}:{lineno}: unknown tool `{name}` (did you mean `{best}`?)")
            elif strict:
                findings.append(f"{path}:{lineno}: unknown tool `{name}` in skill file")
            seen.add(name)

        for m in BACKTICK_RE.finditer(line):
            token = m.group(1)
            if token in tools or token in ignored or token in all_params or token in seen:
                continue
            best, dist = closest(token, tools)
            if 0 < dist <= NEAR_MISS_MAX:
                findings.append(f"{path}:{lineno}: `{token}` looks like a misspelt tool (`{best}`?)")
    return findings


def _split_args(args: str) -> list[str]:
    """Split a call's argument text on top-level commas."""
    out, depth, cur, quote = [], 0, [], None
    for ch in args:
        if quote:
            cur.append(ch)
            if ch == quote:
                quote = None
            continue
        if ch in "'\"":
            quote = ch
        elif ch in "[{":
            depth += 1
        elif ch in "]}":
            depth -= 1
        elif ch == "," and depth == 0:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    if cur:
        out.append("".join(cur))
    return [a for a in out if a.strip() and a.strip() != "..."]


def main(argv: list[str]) -> int:
    tools = load_tools()
    paths: list[Path] = []
    for pattern in DEFAULT_DOCS:
        paths.extend(sorted(BASE_DIR.glob(pattern)))
    paths.extend(Path(p) for p in argv)

    findings = check_tool_counts()
    for path in paths:
        if path.is_file():
            findings.extend(review(path, path.read_text(encoding="utf-8"), tools))
        else:
            findings.append(f"{path}: file not found")

    print(f"{len(tools)} tools registered, {len(paths)} documents checked.")
    for f in findings:
        print(f)
    if findings:
        print(f"{len(findings)} finding(s).")
        return 1
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
