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
| | ↳ every source block | Lists the columns it found, and clicking one drops its reference into the input you were last typing in (with `{braces}` where that field is a text template). CSV and Excel re-read the file as you type; the database block reads on ⟳ only, and asks for the query's **result schema** — no rows are fetched and `{expressions}` count as `NULL`, so it is safe to press against a live table. |
| | ↳ open CSV / open Excel | Say which row holds the column names — report titles and blank rows above it are skipped (`0` = no header row, columns are then `A`, `B`, `C`…). Excel also picks the sheet from the workbook's own list; CSV detects comma, semicolon and tab separators. |
| | filter data source | New source containing only rows matching a condition, e.g. `CustomerId = c.Id` — the way to get "the orders of the current customer". |
| Control (gold) | for each record | Loops its contents once per record; the record is available under an alias (`c` → `c.Name`). |
| | if / else | Conditional content — paragraphs, tables, even whole documents. |
| | switch + case | Picks one branch by value. Fill the switch with `case` blocks (a case answers to one value, or to several separated by commas: `"MX", "GT"`); anything none of them matched runs under *otherwise*. |
| Variables (orange) | set variable | Store a computed value under a name. |
| Folders (purple) | make folder | Creates a folder and saves everything produced inside the block there. Nests: put it in a loop for one folder per record, and a `/` in the name makes another level (`{c.City}/{c.Company}`). |
| | zip folder | The same, except the folder is zipped to `<name>.zip` beside itself when the block ends, and removed unless you tick *keep the folder too*. Everything nests inside it — loops, `make folder`, documents — so one archive can hold a whole tree. |
| Document (blue) | new document | Starts a Word document, saved when the block finishes. Place inside a loop for one document per record. Supports an optional `.dotx` template. |
| | Word paragraphs | **Real Word content embedded in the block** — the way to add text. |
| | ↳ tables from records | Draw the table in the fragment, put `{repeat orders}` in one row, and that row is filled once per record of that source — `{Product}`, `{Qty}` and the like resolve against each record, while the rest of the row's `{expressions}` and everything outside the table still resolve as usual. The row's own formatting is what every generated row gets, so shading, borders, fonts and column widths are whatever you drew. No records means no rows, leaving just the header. |
| | ↳ how it works | Click *Edit in Word* — the fragment's own `.docx` opens in Microsoft Word, where you write and format freely (styles, colors, bullets, anything). At run time the fragment is inserted with full fidelity and `{expressions}` in its text are substituted via Word find/replace, inheriting the surrounding formatting. Try `mos richsample` for a generated demo. |
| | add table | Table filled from a data source with computed columns: `Header: expression \| Header: expression` (empty = all columns). |
| | page break | What it says. |

## Expression language

Used in conditions, computed values, and `{placeholders}` inside text.

- **Fields**: `c.Company` (loop alias), bare `Amount` (innermost record), `[Order Date]` / `c.[Order Date]` (brackets for names with spaces or punctuation, names starting with a digit, and names that collide with a keyword such as `[Null]`)
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
- Folder names are built from your data, so they are cleaned before use: characters a file
  name cannot hold become `_`, and `..` segments are dropped. A program can only ever write
  inside its own output folder.
- The editor has a light and a dark theme; the button at the right end of the toolbar
  switches between them. The first run follows the Windows app mode, and the choice is
  remembered in `%APPDATA%\MergeOnSteroids\settings.json`. Fragment and paragraph previews
  stay on white "paper" in both themes, since they show what Word will print.
- Word generation requires desktop Microsoft Word (any recent version; the writer uses
  late-bound OLE automation, no interop assemblies).
- The old `add paragraph` block — text plus the *name* of a Word style — is no longer
  offered in the palette: authoring the paragraph in Word says the same thing without
  guessing. Programs that already contain one keep loading and running unchanged, and
  such a block previews itself by having Word lay the paragraph out in a hidden document
  built from the enclosing `new document` block's template. That is the same call the
  generator makes, so the block cannot show you something a run would not produce. One
  hidden Word instance serves it: ~1.5 s to start, then ~20 ms per paragraph, repeats
  cached. On by default; the toolbar's *Word previews* switches it off.
- Previews are drawn at Word's own size: the picture carries the DPI Word rendered it
  at, and the block shows it unscaled, so a paragraph on the canvas is the size it is in
  the document at 100% zoom.
- Database query interpolation inlines values into the SQL text — only use it with trusted
  data (parameterized queries are on the roadmap).
