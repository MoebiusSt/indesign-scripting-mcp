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
import win32con
import win32gui

log = logging.getLogger("indesign_com")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCRIPT_LANGUAGE_JAVASCRIPT = 1246973031
SCRIPT_LANGUAGE_UXPSCRIPT = 1431522407

# UndoModes enum values (from InDesign DOM / OMV XML)
UNDO_ENTIRE_SCRIPT = 1699963733      # UndoModes.ENTIRE_SCRIPT
UNDO_SCRIPT_REQUEST = 1699967573     # UndoModes.SCRIPT_REQUEST
UNDO_AUTO = 1699963221               # UndoModes.AUTO_UNDO
UNDO_FAST_ENTIRE = 1699964501        # UndoModes.FAST_ENTIRE_SCRIPT

_UXP_LABEL_KEY = "__mcp_uxp_result__"

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

# Pin open documents by id before user code; optionally close foreign docs that
# appear mid-script (toxic null hyperlink.destination, etc.). Only closes
# foreign docs that are saved and unmodified. $CLOSE_UNEXPECTED$ -> true|false.
_DOC_PIN_HELPER_ES = (
    "var __mcpPinnedIds = {};\n"
    "var __mcpPinnedActiveId = null;\n"
    "var __mcpUnexpected = [];\n"
    "var __mcpClosedForeign = [];\n"
    "try {\n"
    "    for (var __mcpPi = 0; __mcpPi < app.documents.length; __mcpPi++) {\n"
    "        try { __mcpPinnedIds[app.documents[__mcpPi].id] = true; } catch (__mcpPe) {}\n"
    "    }\n"
    "    if (app.documents.length) {\n"
    "        try { __mcpPinnedActiveId = app.activeDocument.id; } catch (__mcpPa) {}\n"
    "    }\n"
    "} catch (__mcpPs) {}\n"
    "\n"
    "function __mcpCloseUnexpected() {\n"
    "    try {\n"
    "        var __mcpI, __mcpD, __mcpId, __mcpRec;\n"
    "        for (__mcpI = app.documents.length - 1; __mcpI >= 0; __mcpI--) {\n"
    "            __mcpD = app.documents[__mcpI];\n"
    "            try { __mcpId = __mcpD.id; } catch (__mcpE1) { continue; }\n"
    "            if (__mcpPinnedIds[__mcpId]) continue;\n"
    "            __mcpRec = { name: \"\", path: \"\", saved: false, modified: false, closed: false };\n"
    "            try { __mcpRec.name = String(__mcpD.name); } catch (__mcpE2) {}\n"
    "            try { __mcpRec.path = String(__mcpD.fullName.fsName); } catch (__mcpE3) {}\n"
    "            try { __mcpRec.saved = !!__mcpD.saved; } catch (__mcpE4) {}\n"
    "            try { __mcpRec.modified = !!__mcpD.modified; } catch (__mcpE5) {}\n"
    "            if ($CLOSE_UNEXPECTED$ && __mcpRec.saved && !__mcpRec.modified) {\n"
    "                try {\n"
    "                    __mcpD.close(SaveOptions.NO);\n"
    "                    __mcpRec.closed = true;\n"
    "                    __mcpClosedForeign.push(__mcpRec.path);\n"
    "                } catch (__mcpE6) {}\n"
    "            }\n"
    "            __mcpUnexpected.push(__mcpRec);\n"
    "        }\n"
    "        if (__mcpPinnedActiveId !== null) {\n"
    "            for (__mcpI = 0; __mcpI < app.documents.length; __mcpI++) {\n"
    "                try {\n"
    "                    if (app.documents[__mcpI].id === __mcpPinnedActiveId) {\n"
    "                        app.activeDocument = app.documents[__mcpI];\n"
    "                        break;\n"
    "                    }\n"
    "                } catch (__mcpE7) {}\n"
    "            }\n"
    "        }\n"
    "    } catch (__mcpPx) {}\n"
    "}\n"
    "\n"
    "function __mcpWarning() {\n"
    "    if (!__mcpUnexpected.length) return \"\";\n"
    "    var kept = [], k;\n"
    "    for (k = 0; k < __mcpUnexpected.length; k++) {\n"
    "        if (!__mcpUnexpected[k].closed) kept.push(__mcpUnexpected[k].name);\n"
    "    }\n"
    "    var w = \"Foreign document(s) appeared during this script (\" + __mcpUnexpected.length +\n"
    "            \"). Code after that point that used app.activeDocument may have targeted the wrong file. \" +\n"
    "            \"Re-check the host by name/id.\";\n"
    "    if (kept.length) w += \" LEFT OPEN because modified or never saved: \" + kept.join(\", \") +\n"
    "                          \". Inspect them; do not assume your edits reached the host.\";\n"
    "    return w;\n"
    "}\n"
)

_DOC_PIN_HELPER_UXP = (
    "let __mcpPinnedIds = {};\n"
    "let __mcpPinnedActiveId = null;\n"
    "let __mcpUnexpected = [];\n"
    "let __mcpClosedForeign = [];\n"
    "let __mcpIndApp = (typeof app !== 'undefined' && app)"
    " ? app : (typeof require !== 'undefined' ? require('indesign').app : null);\n"
    "try {\n"
    "    if (__mcpIndApp) {\n"
    "        for (let __mcpPi = 0; __mcpPi < __mcpIndApp.documents.length; __mcpPi++) {\n"
    "            try { __mcpPinnedIds[__mcpIndApp.documents[__mcpPi].id] = true; } catch (__mcpPe) {}\n"
    "        }\n"
    "        if (__mcpIndApp.documents.length) {\n"
    "            try { __mcpPinnedActiveId = __mcpIndApp.activeDocument.id; } catch (__mcpPa) {}\n"
    "        }\n"
    "    }\n"
    "} catch (__mcpPs) {}\n"
    "\n"
    "function __mcpCloseUnexpected() {\n"
    "    if (!__mcpIndApp) return;\n"
    "    try {\n"
    "        const { SaveOptions } = require('indesign');\n"
    "        for (let __mcpI = __mcpIndApp.documents.length - 1; __mcpI >= 0; __mcpI--) {\n"
    "            const __mcpD = __mcpIndApp.documents[__mcpI];\n"
    "            let __mcpId;\n"
    "            try { __mcpId = __mcpD.id; } catch (__mcpE1) { continue; }\n"
    "            if (__mcpPinnedIds[__mcpId]) continue;\n"
    "            const __mcpRec = { name: '', path: '', saved: false, modified: false, closed: false };\n"
    "            try { __mcpRec.name = String(__mcpD.name); } catch (__mcpE2) {}\n"
    "            try { __mcpRec.path = String(__mcpD.fullName.fsName); } catch (__mcpE3) {}\n"
    "            try { __mcpRec.saved = !!__mcpD.saved; } catch (__mcpE4) {}\n"
    "            try { __mcpRec.modified = !!__mcpD.modified; } catch (__mcpE5) {}\n"
    "            if ($CLOSE_UNEXPECTED$ && __mcpRec.saved && !__mcpRec.modified) {\n"
    "                try {\n"
    "                    __mcpD.close(SaveOptions.NO);\n"
    "                    __mcpRec.closed = true;\n"
    "                    __mcpClosedForeign.push(__mcpRec.path);\n"
    "                } catch (__mcpE6) {}\n"
    "            }\n"
    "            __mcpUnexpected.push(__mcpRec);\n"
    "        }\n"
    "        if (__mcpPinnedActiveId !== null) {\n"
    "            for (let __mcpI = 0; __mcpI < __mcpIndApp.documents.length; __mcpI++) {\n"
    "                try {\n"
    "                    if (__mcpIndApp.documents[__mcpI].id === __mcpPinnedActiveId) {\n"
    "                        __mcpIndApp.activeDocument = __mcpIndApp.documents[__mcpI];\n"
    "                        break;\n"
    "                    }\n"
    "                } catch (__mcpE7) {}\n"
    "            }\n"
    "        }\n"
    "    } catch (__mcpPx) {}\n"
    "}\n"
    "\n"
    "function __mcpWarning() {\n"
    "    if (!__mcpUnexpected.length) return '';\n"
    "    const kept = [];\n"
    "    for (let k = 0; k < __mcpUnexpected.length; k++) {\n"
    "        if (!__mcpUnexpected[k].closed) kept.push(__mcpUnexpected[k].name);\n"
    "    }\n"
    "    let w = 'Foreign document(s) appeared during this script (' + __mcpUnexpected.length +\n"
    "            '). Code after that point that used app.activeDocument may have targeted the wrong file. ' +\n"
    "            'Re-check the host by name/id.';\n"
    "    if (kept.length) w += ' LEFT OPEN because modified or never saved: ' + kept.join(', ') +\n"
    "                          '. Inspect them; do not assume your edits reached the host.';\n"
    "    return w;\n"
    "}\n"
)


def _build_wrapper(user_code: str, close_unexpected_documents: bool = True) -> str:
    """Build the full JSX wrapper around user code.

    Architecture — **no eval()**:
    - User code runs inline inside an IIFE
    - User assigns return data to ``__result``
    - The wrapper serialises ``__result`` via ``__safeStringify``
      (DOM-safe, handles circular refs, dangerous properties, etc.)
    - try/catch provides structured error reporting
    - userInteractionLevel is guarded to prevent modal-dialog blocking
    - Optionally pins the open-document set and closes foreign docs that
      appear mid-script (toxic null hyperlink.destination side effect)
    - The IIFE returns a JSON string; DoScript transports it back via COM

    Convention for user code::

        var doc = app.activeDocument;
        __result = {name: doc.name, pages: doc.pages.length};
    """
    polyfill = _get_polyfill()
    close_lit = "true" if close_unexpected_documents else "false"
    pin = _DOC_PIN_HELPER_ES.replace("$CLOSE_UNEXPECTED$", close_lit)

    # The IIFE ensures a clean scope and allows explicit ``return``.
    # DoScript returns the value of the last expression — the IIFE call.
    return (
        polyfill + "\n"
        "(function() {\n"
        "var __result;\n"
        "var __uilevel = app.scriptPreferences.userInteractionLevel;\n"
        "app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;\n"
        + pin
        + "try {\n"
        # === USER CODE (assigns to __result) ===
        + user_code + "\n"
        # === END USER CODE ===
        "__mcpCloseUnexpected();\n"
        "app.scriptPreferences.userInteractionLevel = __uilevel;\n"
        "if (typeof __result === 'undefined') {\n"
        "    return __safeStringify({success: true, result: null"
        ", closed_unexpected_documents: __mcpClosedForeign"
        ", unexpected_documents: __mcpUnexpected"
        ", warning: __mcpWarning()});\n"
        "}\n"
        "try {\n"
        "    return __safeStringify({success: true, result: __result"
        ", closed_unexpected_documents: __mcpClosedForeign"
        ", unexpected_documents: __mcpUnexpected"
        ", warning: __mcpWarning()});\n"
        "} catch(jsonErr) {\n"
        "    return __safeStringify({success: true, result: String(__result)"
        ", closed_unexpected_documents: __mcpClosedForeign"
        ", unexpected_documents: __mcpUnexpected"
        ", warning: __mcpWarning()});\n"
        "}\n"
        "} catch(e) {\n"
        "try { __mcpCloseUnexpected(); } catch(x2) {}\n"
        "try { app.scriptPreferences.userInteractionLevel = __uilevel; } catch(x) {}\n"
        "return __safeStringify({\n"
        "    success: false,\n"
        "    error: e.message || String(e),\n"
        "    name: e.name || 'Error',\n"
        "    line: typeof e.line === 'number' ? e.line : -1,\n"
        "    closed_unexpected_documents: __mcpClosedForeign,\n"
        "    unexpected_documents: __mcpUnexpected,\n"
        "    warning: __mcpWarning()\n"
        "});\n"
        "}\n"
        "})();\n"
    )


def _build_uxp_wrapper(user_code: str, close_unexpected_documents: bool = True) -> str:
    """Build the full UXP wrapper around user code.

    Architecture for UXP Scripting (ES6+ / V8 engine):
    - User code runs in an async IIFE: ``(async () => { ... })()``
    - Top-level ``await`` inside user code is natively supported
    - User can import ``const { app } = require('indesign')`` or use global ``app``
    - User assigns return data to ``__result``
    - try/catch intercepts all runtime errors to prevent InDesign modal error dialogs
    - Optionally pins the open-document set and closes foreign docs mid-script
    - Output is safely written to ``app.insertLabel('__mcp_uxp_result__', ...)``
      because COM DoScript does not return the last expression value for UXP scripts
    - Python polls the label via COM and clears it
    """
    close_lit = "true" if close_unexpected_documents else "false"
    pin = _DOC_PIN_HELPER_UXP.replace("$CLOSE_UNEXPECTED$", close_lit)
    return (
        "(async () => {\n"
        "let __result;\n"
        "let __error = null;\n"
        + pin
        + "try {\n"
        + user_code + "\n"
        "} catch(e) {\n"
        "    __error = {\n"
        "        message: e.message || String(e),\n"
        "        name: e.name || 'Error',\n"
        "        stack: e.stack || '',\n"
        "        line: typeof e.line === 'number' ? e.line : -1\n"
        "    };\n"
        "}\n"
        "try { __mcpCloseUnexpected(); } catch (__mcpCe) {}\n"
        "try {\n"
        "    let indApp = (typeof app !== 'undefined' && app) ? app : (typeof require !== 'undefined' ? require('indesign').app : null);\n"
        "    if (indApp) {\n"
        "        const payload = JSON.stringify(__error ? {\n"
        "            success: false,\n"
        "            error: __error.message,\n"
        "            name: __error.name,\n"
        "            stack: __error.stack,\n"
        "            line: __error.line,\n"
        "            closed_unexpected_documents: __mcpClosedForeign,\n"
        "            unexpected_documents: __mcpUnexpected,\n"
        "            warning: __mcpWarning()\n"
        "        } : {\n"
        "            success: true,\n"
        "            result: typeof __result !== 'undefined' ? __result : null,\n"
        "            closed_unexpected_documents: __mcpClosedForeign,\n"
        "            unexpected_documents: __mcpUnexpected,\n"
        "            warning: __mcpWarning()\n"
        "        });\n"
        f"        indApp.insertLabel('{_UXP_LABEL_KEY}', payload);\n"
        "    }\n"
        "} catch(reportErr) {}\n"
        "})();\n"
    )


# Simple expression evaluator (no undo, no full wrapper).
# If foreign docs appeared, append "\n[MCP WARNING] " + warning text.
_JSX_EVAL_TEMPLATE = (
    "(function() {\n"
    "$PIN$"
    "    function __mcpEvalOut(val) {\n"
    "        var s = String(val);\n"
    "        var w = __mcpWarning();\n"
    "        if (w) s += \"\\n[MCP WARNING] \" + w;\n"
    "        return s;\n"
    "    }\n"
    "    try {\n"
    "        var __r = $EXPRESSION$;\n"
    "        __mcpCloseUnexpected();\n"
    "        if (typeof __r === 'undefined') return __mcpEvalOut('undefined');\n"
    "        if (__r === null) return __mcpEvalOut('null');\n"
    "        return __mcpEvalOut(__r);\n"
    "    } catch(e) {\n"
    "        try { __mcpCloseUnexpected(); } catch(x) {}\n"
    "        return __mcpEvalOut('ERROR: ' + (e.message || String(e)));\n"
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
# JSX / UXP Execution (public API)
# ---------------------------------------------------------------------------

def run_jsx(
    code: str,
    undo_name: str = "Agent Script",
    undo_mode: str = "entire",
    timeout: int | None = None,
    language: str = "javascript",
    close_unexpected_documents: bool = True,
) -> dict:
    """Execute JSX (ExtendScript) or UXP code in InDesign with safety wrapping.

    Args:
        code: The JavaScript code to execute. Assign to ``__result`` to return data.
        undo_name: Label for the undo step (shown in Edit > Undo).
        undo_mode: ``"entire"`` groups all changes as one undo step (default),
                   ``"auto"`` lets InDesign handle undo per-operation,
                   ``"none"`` skips undo tracking (for read-only operations).
        timeout: Currently unused (reserved for future threading-based timeout).
        language: ``"javascript"`` (default, ExtendScript ES3) or ``"uxp"`` (modern ES6+ / V8).
        close_unexpected_documents: When True (default), close foreign documents
            that appeared mid-script only if they are saved and unmodified.
            Modified or untitled foreign docs stay open. Always reports
            ``unexpected_documents`` / ``warning`` when any foreign doc appeared.
            Set False when the script intentionally opens another file that must
            remain open (detection/warning still run).

    Returns:
        dict with ``success`` key. On success: ``{success: True, result: ...}``.
        On error: ``{success: False, error: str, name: str, line: int}``.
        May include ``closed_unexpected_documents``, ``unexpected_documents``,
        ``warning``.
    """
    app = connect()

    lang_lower = language.lower().strip()
    if lang_lower in ("uxp", "uxpscript"):
        wrapped = _build_uxp_wrapper(code, close_unexpected_documents=close_unexpected_documents)
        if undo_mode == "entire":
            return _execute_uxp_with_undo(app, wrapped, undo_name, UNDO_ENTIRE_SCRIPT)
        elif undo_mode == "auto":
            return _execute_uxp_with_undo(app, wrapped, undo_name, UNDO_SCRIPT_REQUEST)
        else:
            return _execute_uxp(app, wrapped)

    # Default: ExtendScript (ES3)
    wrapped = _build_wrapper(code, close_unexpected_documents=close_unexpected_documents)

    if undo_mode == "entire":
        return _execute_with_undo(app, wrapped, undo_name, UNDO_ENTIRE_SCRIPT)
    elif undo_mode == "auto":
        return _execute_with_undo(app, wrapped, undo_name, UNDO_SCRIPT_REQUEST)
    else:
        # "none" or unknown → plain 2-param DoScript (no undo grouping)
        return _execute(app, wrapped)


def eval_expr(
    expression: str,
    timeout: int | None = None,
    close_unexpected_documents: bool = True,
) -> str:
    """Evaluate a simple expression in InDesign.

    Returns the result as a string. No undo wrapping.
    For quick read-only queries (e.g. ``app.activeDocument.pages.length``).
    When ``close_unexpected_documents`` is True (default), foreign documents
    opened as a side effect of the expression are closed without saving.
    """
    app = connect()
    close_lit = "true" if close_unexpected_documents else "false"
    pin = _DOC_PIN_HELPER_ES.replace("$CLOSE_UNEXPECTED$", close_lit)
    jsx = (
        _JSX_EVAL_TEMPLATE
        .replace("$PIN$", pin)
        .replace("$EXPRESSION$", expression)
    )
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
    ExtendScript engine finishes. There is no safe way to abort a
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
        parsed["_engine"] = "extendscript"
        return parsed
    except pywintypes.com_error as e:
        return _com_error_to_dict(e)


def _dismiss_indesign_modal_dialogs() -> bool:
    """Check for and dismiss any modal error/alert dialogs belonging to InDesign.

    Returns True if a dialog was dismissed, False otherwise.
    """
    found = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
            title = win32gui.GetWindowText(hwnd)
            if "indesign" in title.lower() or "adobe" in title.lower():
                found.append(hwnd)
                win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return len(found) > 0


def _extract_uxp_result(app, t0: float, timeout: int = DEFAULT_TIMEOUT) -> dict:
    """Poll for UXP script execution result from label.

    Handles asynchronous Promise resolution and catches/dismisses any
    modal syntax error dialogs triggered by InDesign's UXP engine.
    """
    raw_result = ""
    poll_interval = 0.01

    while time.monotonic() - t0 < timeout:
        try:
            raw_result = app.ExtractLabel(_UXP_LABEL_KEY)
            if raw_result:
                try:
                    app.InsertLabel(_UXP_LABEL_KEY, "")
                except Exception:
                    pass
                break
        except pywintypes.com_error as e:
            # A modal dialog is blocking COM calls (e.g. parse/syntax error)
            if _dismiss_indesign_modal_dialogs():
                time.sleep(0.05)
                return {
                    "success": False,
                    "error": "Syntax or parse error in UXP script triggered a modal error dialog in InDesign (auto-dismissed). Check script syntax.",
                    "name": "SyntaxError",
                    "line": -1,
                    "_elapsed_s": round(time.monotonic() - t0, 2),
                    "_engine": "uxp",
                }
        except Exception:
            pass

        # Check if an engine modal error dialog appeared (e.g. syntax/parse error)
        if (time.monotonic() - t0) > 0.2:
            if _dismiss_indesign_modal_dialogs():
                time.sleep(0.05)
                return {
                    "success": False,
                    "error": "Syntax or parse error in UXP script triggered a modal error dialog in InDesign (auto-dismissed). Check script syntax.",
                    "name": "SyntaxError",
                    "line": -1,
                    "_elapsed_s": round(time.monotonic() - t0, 2),
                    "_engine": "uxp",
                }

        time.sleep(poll_interval)
        if poll_interval < 0.05:
            poll_interval += 0.01

    elapsed = round(time.monotonic() - t0, 2)
    if raw_result:
        parsed = _parse_result(raw_result)
    else:
        # Timeout occurred - check if a modal dialog was lingering
        if _dismiss_indesign_modal_dialogs():
            parsed = {
                "success": False,
                "error": "Syntax or parse error in UXP script triggered a modal error dialog in InDesign (auto-dismissed).",
                "name": "SyntaxError",
                "line": -1,
            }
        else:
            parsed = {
                "success": False,
                "error": f"UXP script execution timed out after {timeout}s (pending async Promise, setTimeout/setInterval timer, or unresolved fetch). Note: macrotask timers and network sockets are not pumped in standalone COM scripts.",
                "name": "TimeoutError",
                "line": -1,
            }

    parsed["_elapsed_s"] = elapsed
    parsed["_engine"] = "uxp"
    return parsed


def _execute_uxp_with_undo(app, uxp_code: str, undo_name: str,
                           undo_mode: int = UNDO_ENTIRE_SCRIPT) -> dict:
    """Execute UXP script via DoScript with explicit UndoMode and label extraction."""
    _dismiss_indesign_modal_dialogs()
    t0 = time.monotonic()
    try:
        try:
            app.InsertLabel(_UXP_LABEL_KEY, "")
        except Exception:
            pass

        app.DoScript(
            uxp_code,
            SCRIPT_LANGUAGE_UXPSCRIPT,
            [],                    # withArguments (empty)
            undo_mode,             # UndoModes enum value
            undo_name,
        )
        return _extract_uxp_result(app, t0)
    except pywintypes.com_error as e:
        _dismiss_indesign_modal_dialogs()
        return _com_error_to_dict(e)


def _execute_uxp(app, uxp_code: str) -> dict:
    """Execute UXP script via plain DoScript (no undo parameters)."""
    _dismiss_indesign_modal_dialogs()
    t0 = time.monotonic()
    try:
        try:
            app.InsertLabel(_UXP_LABEL_KEY, "")
        except Exception:
            pass

        app.DoScript(uxp_code, SCRIPT_LANGUAGE_UXPSCRIPT)
        return _extract_uxp_result(app, t0)
    except pywintypes.com_error as e:
        _dismiss_indesign_modal_dialogs()
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
