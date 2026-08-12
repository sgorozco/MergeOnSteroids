# Merge on Steroids

A Scratch-like visual language for generating highly dynamic Microsoft Word documents —
mail merge on steroids.

Build a program by dragging colored blocks: connect data sources (Excel, CSV, databases),
loop over their records, open secondary sources filtered by the current record, add
conditional paragraphs and computed values, and fill tables. Running the program drives
desktop Word through OLE automation and produces one `.docx` per iteration (or whatever
structure your blocks describe).

## Projects

| Project | What it is |
|---|---|
| `src/MergeOnSteroids.Core` | Block model, data source loaders, expression engine, interpreter, document writers (Word OLE + text preview). No UI dependencies. |
| `src/MergeOnSteroids.App` | WPF visual editor: block palette, drag & drop canvas, inline editing, live paragraph previews, run panel. |
| `src/MergeOnSteroids.Cli` | `mos` command-line runner: execute programs headlessly, generate the sample. |

## Quick start

```
dotnet build
dotnet run --project src/MergeOnSteroids.App
```

Click **Load sample** — it creates `Documents\MergeOnSteroids\Sample` with two CSVs and a
program that produces one account-statement letter per customer, including a filtered
order table, computed totals, and conditional paragraphs.

- **▶ Preview run** (F5) — fast text-only rendering, no Word involved. Great for iterating.
- **▶ Generate in Word** (F6) — drives Microsoft Word via COM and saves real `.docx` files.

From the command line:

```
mos sample C:\demo
mos run C:\demo\statement-letters.mos.json          # text preview
mos run C:\demo\statement-letters.mos.json --word   # real .docx via Word
```

Programs are JSON files (`*.mos.json`) — friendly to source control.

## Blocks

| Category | Block | Purpose |
|---|---|---|
| Data sources (green) | open CSV / open Excel / open database | Load a named data source (whole table in memory). Database queries support `{expression}` interpolation, so a query can be driven by the current record. |
| | filter data source | New source containing only rows matching a condition, e.g. `CustomerId = c.Id` — the way to get "the orders of the current customer". |
| Control (gold) | for each record | Loops its contents once per record; the record is available under an alias (`c` → `c.Name`). |
| | if / else | Conditional content — paragraphs, tables, even whole documents. |
| Variables (orange) | set variable | Store a computed value under a name. |
| Document (blue) | new document | Starts a Word document, saved when the block finishes. Place inside a loop for one document per record. Supports an optional `.dotx` template. |
| | add paragraph | Styled paragraph (Normal, Heading 1–3, Title, Subtitle, Quote, List Bullet + bold/italic). Text supports `{expressions}`. The block shows a live preview styled the way Word will render it. |
| | Word paragraphs (rich) | **Real Word content embedded in the block.** Click *Edit in Word* — the fragment's own `.docx` opens in Microsoft Word, where you write and format freely (styles, colors, bullets, anything). At run time the fragment is inserted with full fidelity and `{expressions}` in its text are substituted via Word find/replace, inheriting the surrounding formatting. Try `mos richsample` for a generated demo. |
| | add table | Table filled from a data source with computed columns: `Header: expression \| Header: expression` (empty = all columns). |
| | page break | What it says. |

## Expression language

Used in conditions, computed values, and `{placeholders}` inside text.

- **Fields**: `c.Company` (loop alias), bare `Amount` (innermost record), `[Order Date]` (names with spaces)
- **Operators**: `+ - * / %`, `&` (concat), `= <> < <= > >=`, `AND OR NOT`
- **Literals**: `123`, `12.5`, `"text"`, `TRUE`, `FALSE`, `NULL`
- **Functions**: `UPPER LOWER TRIM LEN LEFT RIGHT REPLACE CONTAINS STARTSWITH ENDSWITH CONCAT FORMAT ROUND ABS VALUE IIF ISNULL/COALESCE TODAY NOW YEAR MONTH DAY`
- **Aggregates over a source**: `COUNT(orders)`, `SUM(orders, Qty * UnitPrice)`, `AVG(...)`, `MIN(...)`, `MAX(...)`, `FIRST(...)`

Examples:

```
Dear {c.Contact},
Grand total: {FORMAT(SUM(custOrders, Qty * UnitPrice), "C")}
if:  COUNT(custOrders) > 0 AND c.Balance > 1000
file name:  Statement_{c.Company}
```

`FORMAT` uses .NET format strings with the current culture (`"C"` = currency, `"D"` = long date).

## Files a program owns

```
letters.mos.json              the program (plain JSON, safe to diff and review)
customers.xlsx                your data
fragments/greeting.docx       one file per Word fragment — open it in Word any time
fragments/greeting.preview.png  Word's own rendering, shown inside the block
output/                       generated documents
```

Fragments are ordinary Word documents, so they double as a reusable paragraph library:
point two blocks (or two programs) at the same file and both use it. If you edit a
fragment directly in Word rather than through the editor, press ⟳ on the block to
re-read it.

## Notes

- Relative file paths (data sources, templates, fragments, output folder) resolve against
  the folder containing the `.mos.json` program — keep the program next to its data.
- The editor has a light and a dark theme; the button at the right end of the toolbar
  switches between them. The first run follows the Windows app mode, and the choice is
  remembered in `%APPDATA%\MergeOnSteroids\settings.json`. Fragment and paragraph previews
  stay on white "paper" in both themes, since they show what Word will print.
- Word generation requires desktop Microsoft Word (any recent version; the writer uses
  late-bound OLE automation, no interop assemblies).
- Database query interpolation inlines values into the SQL text — only use it with trusted
  data (parameterized queries are on the roadmap).
