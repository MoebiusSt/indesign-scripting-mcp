"""
InDesign COM Automation Layer.

Provides safe JSX execution in InDesign via Windows COM/OLE:
- Automatic try/catch wrapping with structured error return
- userInteractionLevel protection (prevents modal dialog blocking)
- Undo grouping via UndoModes.ENTIRE_SCRIPT
- DOM-safe JSON serialisation (no eval!)
- Automatic connection management (GetActiveObject -> Dispatch fallback)

Attribution:
- Safety patterns inspired by IdExtenso (Marc Autret, MIT): https://github.com/indiscripts/IdExtenso
- See THIRD_PARTY_NOTICES.md for license text.

Safety patterns:
- Property blacklist for known crash-causing DOM properties
- Structured error objects with line/source/stack info
- DOM-aware JSON serialisation via toSpecifier()
- UnitValue === null bug workaround
"""

import json
import logging
import os
import threading
import time
from pathlib import Path
from typing import Any

import pywintypes
import win32com.client

log = logging.getLogger("indesign_com")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCRIPT_LANGUAGE_JAVASCRIPT = 1246973031

# UndoModes enum values (from InDesign DOM / OMV XML)
UNDO_ENTIRE_SCRIPT = 1699963733      # UndoModes.ENTIRE_SCRIPT
UNDO_SCRIPT_REQUEST = 1699967573     # UndoModes.SCRIPT_REQUEST
UNDO_AUTO = 1699963221               # UndoModes.AUTO_UNDO
UNDO_FAST_ENTIRE = 1699964501        # UndoModes.FAST_ENTIRE_SCRIPT

# ProgIDs to try, newest first
PROGIDS = [
    "InDesign.Application.2026",
    "InDesign.Application.2025",
    "InDesign.Application",
]

DEFAULT_TIMEOUT = int(os.environ.get("INDESIGN_EXEC_TIMEOUT", "30"))

# ---------------------------------------------------------------------------
# JSX Polyfill (loaded once from file, cached)
# ---------------------------------------------------------------------------

_POLYFILL_PATH = Path(__file__).parent / "json_polyfill.jsx"
_polyfill_code: str | None = None


def _get_polyfill() -> str:
    """Load the JSON polyfill / __safeStringify code from file (cached)."""
    global _polyfill_code
    if _polyfill_code is None:
        _polyfill_code = _POLYFILL_PATH.read_text(encoding="utf-8")
    return _polyfill_code


# ---------------------------------------------------------------------------
# JSX Wrapper
# ---------------------------------------------------------------------------

def _build_wrapper(user_code: str) -> str:
    """Build the full JSX wrapper around user code.

    Architecture — **no eval()**:
    - User code runs inline inside an IIFE
    - User assigns return data to ``__result``
    - The wrapper serialises ``__result`` via ``__safeStringify``
      (DOM-safe, handles circular refs, dangerous properties, etc.)
    - try/catch provides structured error reporting
    - userInteractionLevel is guarded to prevent modal-dialog blocking
    - The IIFE returns a JSON string; DoScript transports it back via COM

    Convention for user code::

        var doc = app.activeDocument;
        __result = {name: doc.name, pages: doc.pages.length};
    """
    polyfill = _get_polyfill()

    # The IIFE ensures a clean scope and allows explicit ``return``.
    # DoScript returns the value of the last expression — the IIFE call.
    return (
        polyfill + "\n"
        "(function() {\n"
        "var __result;\n"
        "var __uilevel = app.scriptPreferences.userInteractionLevel;\n"
        "app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;\n"
        "try {\n"
        # === USER CODE (assigns to __result) ===
        + user_code + "\n"
        # === END USER CODE ===
        "app.scriptPreferences.userInteractionLevel = __uilevel;\n"
        "if (typeof __result === 'undefined') {\n"
        "    return __safeStringify({success: true, result: null});\n"
        "}\n"
        "try {\n"
        "    return __safeStringify({success: true, result: __result});\n"
        "} catch(jsonErr) {\n"
        "    return __safeStringify({success: true, result: String(__result)});\n"
        "}\n"
        "} catch(e) {\n"
        "try { app.scriptPreferences.userInteractionLevel = __uilevel; } catch(x) {}\n"
        "return __safeStringify({\n"
        "    success: false,\n"
        "    error: e.message || String(e),\n"
        "    name: e.name || 'Error',\n"
        "    line: typeof e.line === 'number' ? e.line : -1\n"
        "});\n"
        "}\n"
        "})();\n"
    )


# Simple expression evaluator (no undo, no full wrapper)
_JSX_EVAL_TEMPLATE = (
    "(function() {\n"
    "    try {\n"
    "        var __r = $EXPRESSION$;\n"
    "        if (typeof __r === 'undefined') return 'undefined';\n"
    "        if (__r === null) return 'null';\n"
    "        return String(__r);\n"
    "    } catch(e) {\n"
    "        return 'ERROR: ' + (e.message || String(e));\n"
    "    }\n"
    "})();\n"
)


# ---------------------------------------------------------------------------
# Connection Management
# ---------------------------------------------------------------------------

_app = None
_app_lock = threading.Lock()


def connect() -> Any:
    """Connect to a running InDesign instance.

    Tries GetActiveObject first (attach to existing), falls back to Dispatch.
    Does NOT launch InDesign if not running (Dispatch may do so — we check after).

    Returns the COM Application object.
    Raises ConnectionError if InDesign is not reachable.
    """
    global _app

    with _app_lock:
        # Test existing connection
        if _app is not None:
            try:
                _ = _app.Name  # Quick connectivity check
                return _app
            except Exception:
                _app = None  # Connection lost, try reconnect

        last_error = None
        for prog_id in PROGIDS:
            # Try attaching to running instance first
            try:
                _app = win32com.client.GetActiveObject(prog_id)
                return _app
            except pywintypes.com_error:
                pass

            # Try Dispatch (may launch InDesign)
            try:
                _app = win32com.client.Dispatch(prog_id)
                # Verify it's actually running by checking a property
                _ = _app.Name
                return _app
            except pywintypes.com_error as e:
                last_error = e

        raise ConnectionError(
            f"Could not connect to InDesign. Is it running? Last error: {last_error}"
        )


def disconnect():
    """Release the COM connection."""
    global _app
    with _app_lock:
        _app = None


def is_connected() -> bool:
    """Check if we have a live connection to InDesign."""
    global _app
    if _app is None:
        return False
    try:
        _ = _app.Name
        return True
    except Exception:
        _app = None
        return False


# ---------------------------------------------------------------------------
# JSX Execution (public API)
# ---------------------------------------------------------------------------

def run_jsx(
    code: str,
    undo_name: str = "Agent Script",
    undo_mode: str = "entire",
    timeout: int | None = None,
) -> dict:
    """Execute JSX code in InDesign with safety wrapping.

    Args:
        code: The JSX code to execute.  Assign to ``__result`` to return data.
        undo_name: Label for the undo step (shown in Edit > Undo).
        undo_mode: ``"entire"`` groups all changes as one undo step (default),
                   ``"auto"`` lets InDesign handle undo per-operation,
                   ``"none"`` skips undo tracking (for read-only operations).
        timeout: Currently unused (reserved for future threading-based timeout).

    Returns:
        dict with ``success`` key.  On success: ``{success: True, result: ...}``.
        On error: ``{success: False, error: str, name: str, line: int}``.
    """
    app = connect()

    # Build the safe JSX wrapper (IIFE, no eval)
    wrapped = _build_wrapper(code)

    # Map undo_mode to DoScript parameters:
    #   "entire" → ENTIRE_SCRIPT  (one undo step for everything, labelled)
    #   "auto"   → SCRIPT_REQUEST (InDesign creates one undo step per DOM change)
    #   "none"   → plain DoScript without undo params (default SCRIPT_REQUEST
    #              behaviour, but no extra undo grouping — safe for read-only
    #              queries and for the undo tool itself)
    if undo_mode == "entire":
        return _execute_with_undo(app, wrapped, undo_name, UNDO_ENTIRE_SCRIPT)
    elif undo_mode == "auto":
        return _execute_with_undo(app, wrapped, undo_name, UNDO_SCRIPT_REQUEST)
    else:
        # "none" or unknown → plain 2-param DoScript (no undo grouping)
        return _execute(app, wrapped)


def eval_expr(expression: str, timeout: int | None = None) -> str:
    """Evaluate a simple expression in InDesign.

    Returns the result as a string.  No undo wrapping.
    For quick read-only queries (e.g. ``app.activeDocument.pages.length``).
    """
    app = connect()
    jsx = _JSX_EVAL_TEMPLATE.replace("$EXPRESSION$", expression)
    return _execute_raw(app, jsx)


# ---------------------------------------------------------------------------
# Internal Execution
# ---------------------------------------------------------------------------

def _execute(app, jsx_code: str) -> dict:
    """Execute JSX via plain DoScript (no undo params) and parse the JSON result."""
    raw = _execute_raw(app, jsx_code)
    return _parse_result(raw)


def _execute_with_undo(app, jsx_code: str, undo_name: str,
                       undo_mode: int = UNDO_ENTIRE_SCRIPT) -> dict:
    """Execute JSX via DoScript with explicit UndoMode.

    Passes script, language, withArguments, undoMode, undoName as
    positional COM arguments — no string-in-string wrapping needed.

    Note on timeout: COM DoScript blocks the calling thread until the
    ExtendScript engine finishes.  There is no safe way to abort a
    running script from Python — the only option would be to kill the
    InDesign process, which is unacceptable for production documents.
    Instead we measure wall-clock time and emit a warning when execution
    exceeds DEFAULT_TIMEOUT.
    """
    t0 = time.monotonic()
    try:
        result = app.DoScript(
            jsx_code,
            SCRIPT_LANGUAGE_JAVASCRIPT,
            [],                    # withArguments (empty)
            undo_mode,             # UndoModes enum value
            undo_name,
        )
        elapsed = time.monotonic() - t0
        if elapsed > DEFAULT_TIMEOUT:
            log.warning("DoScript took %.1fs (timeout hint: %ds)", elapsed, DEFAULT_TIMEOUT)
        parsed = _parse_result(result)
        parsed["_elapsed_s"] = round(elapsed, 2)
        return parsed
    except pywintypes.com_error as e:
        return _com_error_to_dict(e)


def _execute_raw(app, jsx_code: str) -> str | None:
    """Execute JSX via COM DoScript (no undo parameters).

    Returns the raw result (string, number, None, …).
    Raises no exceptions — COM errors are converted to JSON error strings.
    Used by eval_expr() for lightweight queries.
    """
    t0 = time.monotonic()
    try:
        result = app.DoScript(jsx_code, SCRIPT_LANGUAGE_JAVASCRIPT)
        elapsed = time.monotonic() - t0
        if elapsed > DEFAULT_TIMEOUT:
            log.warning("DoScript (raw) took %.1fs (timeout hint: %ds)", elapsed, DEFAULT_TIMEOUT)
        return result
    except pywintypes.com_error as e:
        # Return a JSON error string so callers always get parseable output
        return json.dumps(_com_error_to_dict(e))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_result(raw) -> dict:
    """Parse a JSON result string from the JSX wrapper into a dict."""
    if raw is None:
        return {"success": True, "result": None}
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except (json.JSONDecodeError, ValueError):
            return {"success": True, "result": raw}
    # COM may return tuples (from JS arrays), ints, bools, etc.
    return {"success": True, "result": raw}


def _com_error_to_dict(e: pywintypes.com_error) -> dict:
    """Extract a human-readable error from a COM exception.

    Also checks whether the error indicates a lost connection (InDesign
    crashed or was closed) and invalidates the cached COM reference so
    the next call will attempt a fresh reconnect.
    """
    global _app
    desc = ""
    hresult = e.args[0] if e.args else 0

    if hasattr(e, "args") and len(e.args) > 2:
        excep = e.args[2]
        if excep and len(excep) > 2 and excep[2]:
            desc = str(excep[2])
    if not desc:
        desc = str(e)

    # Detect connection-loss HRESULTs and invalidate the cached reference.
    # RPC_E_DISCONNECTED       = -2147417848 (0x80010108)
    # RPC_S_SERVER_UNAVAILABLE = -2147023174 (0x800706BA)
    # CO_E_OBJNOTCONNECTED     = -2147220992 (0x80040004 — varies)
    # RPC_E_SERVERFAULT        = -2147417851 (0x80010105)
    connection_loss_codes = {-2147417848, -2147023174, -2147220992, -2147417851}
    if hresult in connection_loss_codes:
        log.warning("Connection to InDesign lost (HRESULT %s). Will reconnect on next call.", hex(hresult & 0xFFFFFFFF))
        _app = None

    return {
        "success": False,
        "error": desc,
        "name": "COMError",
        "line": -1,
        "source": "COM/DoScript",
    }
