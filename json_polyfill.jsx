// DOM-safe JSON serialisation for ExtendScript (ES3).
// Always installs __safeStringify() as a global function, independent of
// native JSON availability.  Used by the MCP JSX wrapper to serialise
// __result safely.
//
// Attribution:
// - Safety patterns inspired by IdExtenso (Marc Autret, MIT)
//   https://github.com/indiscripts/IdExtenso
// - See THIRD_PARTY_NOTICES.md for license text.
//
// Safety features:
//   - InDesign DOM objects  -> toSpecifier() string
//   - Circular references   -> "[circular]"
//   - Max depth (20)        -> "[max depth]"
//   - Dangerous properties  -> silently skipped (try/catch)
//   - UnitValue===null bug  -> workaround via constructor access

// --------------------------------------------------------------------------
// String escape helper
// --------------------------------------------------------------------------
function __jsonStr(s) {
    var out = '', i, cc, ch;
    for (i = 0; i < s.length; i++) {
        cc = s.charCodeAt(i);
        ch = s.charAt(i);
        if      (ch === '\\') out += '\\\\';
        else if (ch === '"')  out += '\\"';
        else if (cc === 8)    out += '\\b';
        else if (cc === 9)    out += '\\t';
        else if (cc === 10)   out += '\\n';
        else if (cc === 12)   out += '\\f';
        else if (cc === 13)   out += '\\r';
        else if (cc < 32)     out += '\\u' + ('0000' + cc.toString(16)).slice(-4);
        else                  out += ch;
    }
    return '"' + out + '"';
}

// --------------------------------------------------------------------------
// Recursive encoder
// --------------------------------------------------------------------------
function __jsonEncode(v, depth, seen) {
    if (depth > 20) return '"[max depth]"';

    var t = typeof v;

    // --- primitives (checked via typeof, before any === comparison,
    //     to avoid the UnitValue===null bug in ExtendScript) ----------
    if (t === 'undefined') return 'null';
    if (t === 'boolean')   return v ? 'true' : 'false';
    if (t === 'number')    return isFinite(v) ? String(v) : 'null';
    if (t === 'string')    return __jsonStr(v);
    if (t === 'function')  return undefined;           // omitted in JSON

    // --- t === 'object' from here ------------------------------------

    // Safe null check: accessing .constructor on null throws TypeError.
    // This avoids the UnitValue(0.5,'pt')===null bug.
    try { var _ctor = v.constructor; } catch (_) { return 'null'; }

    // InDesign DOM object — serialise via toSpecifier().
    // Never iterate DOM properties (risk of crash / freeze).
    if (typeof v.toSpecifier === 'function') {
        try {
            return __jsonStr('[DOM:' + _ctor.name + ':' + v.toSpecifier() + ']');
        } catch (de) {
            return __jsonStr('[DOM object]');
        }
    }

    // Circular reference detection
    var si;
    for (si = 0; si < seen.length; si++) {
        if (v === seen[si]) return '"[circular]"';
    }
    seen.push(v);

    var result, i, k, encoded;

    if (v instanceof Array) {
        // --- Array ---------------------------------------------------
        var a = [];
        for (i = 0; i < v.length; i++) {
            encoded = __jsonEncode(v[i], depth + 1, seen);
            a.push(encoded === undefined ? 'null' : encoded);
        }
        result = '[' + a.join(',') + ']';
    } else {
        // --- Plain object --------------------------------------------
        var p = [];
        for (k in v) {
            if (!v.hasOwnProperty(k)) continue;
            try {
                encoded = __jsonEncode(v[k], depth + 1, seen);
                if (encoded !== undefined) {
                    p.push(__jsonStr(String(k)) + ':' + encoded);
                }
            } catch (propErr) {
                // Silently skip dangerous / unreadable properties
                // Example of dangerous/unreadable properties:
                // scriptPreferences.properties, shadow settings in find/change prefs, etc.
            }
        }
        result = '{' + p.join(',') + '}';
    }

    seen.pop();
    return result;
}

// --------------------------------------------------------------------------
// Public API — always available as a global function
// --------------------------------------------------------------------------
function __safeStringify(v) {
    return __jsonEncode(v, 0, []);
}

// Also install as JSON.stringify polyfill when native JSON is absent,
// so that user code calling JSON.stringify still works on older engines.
if (typeof JSON === 'undefined') { JSON = {}; }
if (typeof JSON.stringify !== 'function') {
    JSON.stringify = __safeStringify;
}
