# Third-Party Notices

This project includes ideas/patterns inspired by the following third-party project.

## IdExtenso (Marc Autret)

- Project: `https://github.com/indiscripts/IdExtenso`
- License: MIT
- Scope in this repository:
  - The JSON/serialization safety patterns referenced in `json_polyfill.jsx` and `indesign_com.py`
    (DOM object handling via `toSpecifier()`, circular reference handling, depth limiting,
    and avoiding known crash-prone property access).

### License text (MIT)

```text
MIT License

Copyright (c) 2017-present, Marc Autret - http://indiscripts.com

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Adobe ExtendScript API reference on indesignjs.de (Gregor Fellenz)

- Source: `https://www.indesignjs.de/indesignapi/` (continuation of the reference by Theunis de Jong / Jongware)
- Used files: `indesign/api.json`, `javascript/api.json`, `scriptui/api.json`
  (downloaded by `manage.py fetch` into `sources/indesignjs/`, not committed)
- The descriptive texts are Adobe's; the transformation is Gregor Fellenz'.
- The SQLite database is generated locally from these files and is not redistributed.
- `manage.py check-uxp` optionally reads the rendered pages of `indesignapi.zip` to compare
  the UXP runtime markers with `uxp_notes.json`.

## extendscriptApiDocTransformations (Gregor Fellenz)

- Project: `https://github.com/grefel/extendscriptApiDocTransformations`
- License: Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (repository license)
- Scope in this repository (from the former OMV XML pipeline, parser versions < 3.0):
  - Conceptual learnings only (no direct code import from XSLT files)
  - Merge strategy for multiple sources (InDesign DOM, Core JavaScript, ScriptUI)
  - Awareness of ScriptUI name collisions vs DOM classes
  - Reference-validation mindset for cross-source links
