# DVT Generator

Generate compact table-format Design Verification Test documents (`.docx`) from a JSON data file. One generator, one consistent layout — use it for any product DVT, with content per-DVT in its own JSON file. Suggested use: fork this repo as a template for each new DVT, or keep one repo and one JSON per product.

## What it produces

- **Cover sheet** with revision, owner, sign-off fields — all text from the JSON's `meta.coverRows`
- **Clickable Table of Contents** — every test row is a hyperlink; page numbers via `PAGEREF`; P/F column mirrors each detail-row checkbox via `REF`
- **Known-issues table** right after cover (optional)
- **External-reference pages** (Section A, B, …) with attestation tables — for anything qualified by a sibling DVT; zero, one, or many supported
- **Per-section test tables**, each on its own page, six columns: `Ref / Header / Brief / Criteria / Measure (blank for tester) / P/F`
- **Word content-control checkboxes** (☐ ↔ ☒) in the P/F column — clickable in Word 2010+ and LibreOffice 7+ without form protection
- **Appendices** — Equipment (A), Release-gate categories + compatibility matrix (B), Elaborations and rationale indexed by ref (C)
- US Letter, 0.7" margins, Arial body, page header + footer page-of-total
- **Section codes + cross-ref tokens** so renumbering is just an array reorder
- **Lint pass** that warns about stale cross-refs, duplicate refs, and unresolved placeholders before rendering

## Files in this repo

```
.
├── build_dvt.js                # Build script. Reads a JSON, writes a .docx.
├── dvt_data.json               # Generic example DVT — exercises every feature
├── dvt_data_template.json      # Blank schema with placeholders — copy this for a new DVT
├── package.json                # npm dependencies (docx, jszip)
├── README.md                   # This file
├── CLAUDE.md                   # Project guide for Claude Code sessions
└── .gitignore                  # Excludes node_modules and generated .docx/.pdf
```

`node_modules/` and generated `.docx`/`.pdf` files are excluded by `.gitignore`.

## Setup

Requires Node.js 18+.

```bash
git clone <repo>
cd dvt-generator
npm install
```

## Generate a DVT

```bash
node build_dvt.js [data.json]
```

Defaults:

- `data.json` → `./dvt_data.json`
- Output filename comes from the data file's meta block:
  - `meta.outputFilename` if set, else
  - sanitized from `meta.headerLine` (em-dash → hyphen, reserved filename chars → underscore), else
  - `<data-basename>.docx`.
- Output directory:
  - `meta.outputDir` if set (absolute, or resolved against the data file's directory), else
  - the data file's directory.

So `node build_dvt.js ../my-product-dvt/data.json` drops the rendered `.docx` next to `data.json`, regardless of where node was invoked from. Set `meta.outputDir: "./build"` to redirect, or set an absolute path to put output in a shared location.

A typical workflow: keep the engine in this repo, keep each product's data file in that product's own folder, and the docx lands next to the data file by default.

Console prints the number of content-control checkboxes injected (one per test row) plus any lint warnings.

## After opening the docx

Both Word and LibreOffice need a one-time field refresh after first open to populate the TOC's page numbers and P/F mirrors:

- **Microsoft Word**: Ctrl+A (select all), then F9.
- **LibreOffice Writer**: Tools → Update → Fields (or F9 with all selected).
- **Single field**: right-click → Update Field.

Toggle any detail-row checkbox by clicking it (☐ ↔ ☒). Refresh fields again to update the TOC mirror.

Optional: in Word, enable **File → Options → Display → "Update fields before printing"** so PDF exports always have current values.

## Spinning up a new DVT

1. Copy `dvt_data_template.json` (or the example `dvt_data.json`) to your product's working folder. Name it whatever you like — the script takes a path.
2. Fill in `meta` — product name, sub-title, descriptor, header line, cover rows, external-reference sections (if any), appendix titles, matrix headers.
3. Fill in `sections` with stable symbolic `code`s and human-readable `name`s. Letters (A, B, C, …) are assigned automatically by array order.
4. Fill in `steps`. Each step gets a `section` (one of your section codes) and a `subref` (the part after the letter, e.g. `"1.1"` becomes `D.1.1` if its section ends up as letter D).
5. Fill in `knownIssues`, `elaborations`, `equipment`, `releaseGate.categories`, `releaseGate.matrix` as needed (any can be empty `[]`).
6. Run `node build_dvt.js <your-data-file>.json`. Output filename and location come from your meta.

If your DVT has no external-reference sections, set `meta.externalRefs: []`. The first internal section becomes Section A.

## Section codes and cross-references

Sections have a stable symbolic `code` and the script assigns the letter (A, B, C, …) from array position. Reordering the `sections` array renumbers the doc — no other edits needed.

```json
"sections": [
  {"code": "mech",   "name": "Mechanical / Pre-energize Checks", "intro": "..."},
  {"code": "func",   "name": "Functional Bench Tests",           "intro": "..."},
  {"code": "stress", "name": "Stress and Environmental",         "intro": "..."}
]
```

Steps reference the section by code:

```json
"steps": [
  {"section": "mech", "subref": "1",   "header": "Visual inspection", ...},
  {"section": "func", "subref": "1.1", "header": "Power-on banner",   ...}
]
```

Cross-references in any text use `[code]` (just the letter) or `[code.subref]` (the full ref):

```json
{"section": "func", "subref": "2.2", "brief": "Continuous polling at production rate for 60 s. Repeat [func.2.1] under load."}
```

Insert a new section in the middle of the array? Every `[func.…]` and `[stress.…]` cross-ref auto-resolves to the new letters on the next build.

## Lint

On every build, the script prints warnings and errors before rendering:

- **Errors (build fails)**: duplicate step refs.
- **Warnings (build continues)**: unresolved `[code]` cross-refs (typo in code), unknown letter-style refs (`D.1.1` with no matching step), unresolved `{placeholder}` (typo in a meta key), missing required meta fields.

Run after any edit. Surfaces stale cross-refs and typos before they ship.

## Today's date

If `meta.date` is missing or empty, the script fills it with today's date in `M/D/YYYY` format. Combine with the `{date}` placeholder in `coverRows` and you get an auto-stamped revision row.

## String interpolation in the data file

Any string anywhere in the JSON can reference a meta field with `{key}`. The script does a single-pass replace before rendering, pulling values from the top-level `meta` block (only string and number meta values are substitutable). Unknown placeholders are left as-is (so typos show up as literal `{foo}` in the rendered doc — useful for spotting mistakes).

The example uses this for revision tracking — single edit bumps everywhere:

```json
"meta": {
  "revision": "0.1",
  "draftStatus": "Draft",
  "headerLine": "{productName} DVT — Rev {revision} ({draftStatus})",
  "coverRows": [
    ...,
    {"label": "Revision", "value": "{revision} ({draftStatus}) — {date}"},
    ...
  ]
}
```

Bumping `revision` from `"0.1"` to `"0.2"` updates the cover-sheet revision row, the page header on every page, and the output filename — all from one edit. The same mechanism works in `steps`, `sections`, `elaborations`, anywhere.

## Data schema

### `meta` (object)

| Key | Type | Notes |
| --- | --- | --- |
| `title` | string | Big top-line on cover (e.g. "DESIGN VERIFICATION TEST"). |
| `productName` | string | Sub-line on cover. |
| `subtitle` | string | Optional descriptor line. |
| `descriptor` | string | Optional italic line at bottom of cover title block. |
| `headerLine` | string | Right-aligned text in every page header. Also drives the default output filename. |
| `outputFilename` | string | Optional explicit default for the output `.docx` (overrides `headerLine`-derived default). |
| `outputDir` | string | Optional output directory. Absolute, or relative to the data file's directory. Default = data file's directory. Created if missing. |
| `creator` | string | Doc metadata Creator field. |
| `coverRows` | `[{label, value}]` | Rows of the cover-sheet table. |
| `revision`, `draftStatus`, `date` | string | Convention. Referenced from `headerLine` and `coverRows` via `{key}` interpolation. `date` defaults to today if absent. |
| `knownIssues.heading` | string | Heading above the known-issues table. |
| `knownIssues.intro` | string | Italic blurb under that heading. |
| `externalRefs` | array | One entry per external-reference section (see below). |
| `appendixA.heading` | string | Equipment appendix heading. |
| `appendixA.intro` | string | Equipment appendix intro. |
| `appendixB.heading` | string | Release-gate appendix heading. |
| `appendixB.intro` | string | Release-gate appendix intro. |
| `appendixB.categoriesTitle` | string | Sub-heading for B.1. |
| `appendixB.matrixTitle` | string | Sub-heading for B.2. |
| `appendixB.matrixHeaders` | `[string]` | Column headers for the compat matrix. Number of headers = number of cells expected per matrix row. |
| `appendixB.matrixColumnWidths` | `[number]` | Optional column widths in DXA (1440 = 1"). Must sum to 10224. If omitted, equal widths. |
| `appendixC.heading` | string | Elaborations appendix heading. |
| `appendixC.intro` | string | Elaborations appendix intro. |

### `meta.externalRefs` entry

| Key | Type | Notes |
| --- | --- | --- |
| `code` | string | Optional symbolic code, referenceable as `[code]` from anywhere in the data. |
| `sectionLetter` | string | Optional explicit single letter. If omitted, assigned by array position. |
| `title` | string | Document name (e.g. "Sub-Assembly DVT"). |
| `intro` | string | Optional italic intro paragraph. |
| `linkPlaceholder` | string | Text shown where the live link goes (default `<Link>`). |
| `passCriteria` | string | Optional. Pass criteria block on the page. |
| `attestHeaders` | `[string, string]` | Optional. If provided, renders a 2-column attestation table for the tester. |

### `steps` entry

Two styles. Preferred:

| Key | Type | Notes |
| --- | --- | --- |
| `section` | string | Section `code` this step belongs to. |
| `subref` | string | The part after the letter (e.g. `"1.1"` becomes `D.1.1` if its section ends up as letter D). |
| `header` | string | 2–5 word test name shown in the Header column and the TOC. |
| `brief` | string | Imperative-voice 1-sentence action. Use `[code]` or `[code.subref]` for cross-refs. |
| `criteria` | string | Terse pass criteria. No full sentences. |
| `measure` | string | Usually `""` — the tester writes the recorded value at run time. |

Legacy (still works): `ref: "D.1.1"` instead of `section` + `subref`. Goes stale if you renumber.

### `sections` entry

| Key | Type | Notes |
| --- | --- | --- |
| `code` | string | Stable symbolic name (e.g. `"func"`). Referenced by step `section` and by `[code]` / `[code.subref]` cross-ref tokens. |
| `name` | string | Section short name (e.g. `"Functional Bench Tests"`). Used in default title. |
| `title` | string | Optional explicit section header. If omitted, defaults to `"Section <letter> — <name>"`. |
| `intro` | string | Optional 1-2 sentence intro under the section heading. |
| `id` | string | Optional legacy single-letter field — only honored when `code` is absent (back-compat). |

### `knownIssues`, `elaborations`, `equipment`, `releaseGate.categories`, `releaseGate.matrix`

See `dvt_data_template.json` for shape. Each can be `[]` if not used.

## How it works

1. **docx-js** generates the OOXML and packs to a buffer. Each P/F cell holds a `☐` glyph wrapped in a `pf_<ref>` bookmark; each Ref cell holds a `test_<ref>` bookmark.
2. **Post-process** (`injectCheckboxControls`) unzips the buffer with JSZip, scans `word/document.xml` for each `pf_*` bookmark, and wraps the run inside it in a `<w:sdt><w14:checkbox>` content control. Re-zips and writes to disk.
3. The TOC table links each test via `InternalHyperlink` to its `test_*` bookmark, reads its page via `<w:fldSimple w:instr="PAGEREF test_… \h"/>`, and mirrors its P/F state via `REF pf_… \h`. Both field types resolve on F9 / "Update Fields".

The post-process is needed because docx-js v9 doesn't have a direct API for content-control checkboxes — the SDT XML is injected after the fact.

## Caveats

- **REF + content control** is finicky on some older Word versions. If the TOC's P/F column stays empty after Ctrl+A → F9 even though page numbers populate, your Word build doesn't expand REF over an SDT. (Tested working on Word 365 / 2021 and LibreOffice 7.5+.)
- **Bookmark naming**: a step ref `D.1.1` becomes bookmark `test_D_1_1` (dots replaced with underscores). Keep all step `ref` values unique within a single doc.
- **TOC section divider rows use `columnSpan: 4`** matching the 4-column TOC (Ref / Header / Page / P/F). If you change the TOC layout, update the columnSpan to match.
- **Plain-letter cross-references in Brief / Criteria text** are plain text. Use the `[code.subref]` form so they survive renumbering. The lint flags plain-letter refs that don't match a real step.
- **Field refresh on first open**: TOC cells appear blank until the first F9 / "Update Fields". Normal — page numbers and REF mirrors are computed by the renderer, not stored in the .docx.

## License

MIT — see [LICENSE](LICENSE). Copyright © 2026 Kane Anderson.
