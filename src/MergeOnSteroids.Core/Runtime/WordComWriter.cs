using System.Runtime.InteropServices;

namespace MergeOnSteroids.Core.Runtime;

/// <summary>
/// Drives desktop Microsoft Word through OLE automation (late-bound COM, so no
/// interop assemblies and no dependency on a specific Office version).
/// </summary>
public sealed class WordComWriter : IDocumentWriter
{
    // Word constants (WdBuiltinStyle etc.) — numeric so they survive late binding
    private const int WdCharacter = 1;
    private const int WdPageBreak = 7;
    private const int WdLineStyleSingle = 1;
    private const int WdFormatXMLDocument = 12; // .docx
    private const int WdCollapseEnd = 0;

    private readonly Dictionary<string, string> _fragmentXmlCache = new(StringComparer.OrdinalIgnoreCase);
    private RunOptions _options = new();          // kept for resolving template paths
    private string _outputFolder = Environment.CurrentDirectory;
    private dynamic? _app;
    private dynamic? _doc;

    public bool InDocument => _doc is not null;

    public string OutputFolder
    {
        get => _outputFolder;
        set { _outputFolder = value; Directory.CreateDirectory(value); }
    }

    public void Begin(RunOptions options)
    {
        _options = options;
        OutputFolder = options.OutputFolder;

        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed (Word.Application COM class not found).");

        _app = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Could not start Microsoft Word.");
        _app!.Visible = options.ShowWord;
        _app!.DisplayAlerts = 0; // wdAlertsNone
    }

    public void BeginDocument(string? templatePath)
    {
        var app = RequireApp();
        if (_doc is not null)
            throw new InvalidOperationException("A document is already open — 'new document' blocks cannot be nested.");

        _doc = string.IsNullOrWhiteSpace(templatePath)
            ? app.Documents.Add()
            : app.Documents.Add(_options.ResolvePath(templatePath));
    }

    public void AddParagraph(string text, string style, bool bold, bool italic)
    {
        var doc = RequireDoc();
        dynamic range = doc.Bookmarks["\\endofdoc"].Range;
        range.Text = text;
        try
        {
            range.Style = Interop.WordStyleIds.Of(style);
        }
        catch (COMException)
        {
            // Style not available in this template — leave as body text
        }
        if (bold) range.Font.Bold = 1;
        if (italic) range.Font.Italic = 1;
        range.InsertParagraphAfter();
        ReleaseCom(range);
    }

    public void AddFragment(FragmentContent fragment, string plainText,
        IReadOnlyList<KeyValuePair<string, string>> replacements,
        IReadOnlyList<FragmentRowRepeat> repeats)
    {
        var doc = RequireDoc();
        var fragmentXml = fragment.InlineXml ?? ReadFragmentXml(fragment.DocxPath);
        if (string.IsNullOrWhiteSpace(fragmentXml)) return;

        // Remember where the fragment starts so substitution stays inside it.
        dynamic insertAt = doc.Bookmarks["\\endofdoc"].Range;
        int start = (int)insertAt.Start;
        insertAt.InsertXML(fragmentXml);
        ReleaseCom(insertAt);

        // Repeated rows first: each clone gets its own record's values, so the
        // whole-fragment pass below finds nothing left to substitute inside them.
        foreach (var repeat in repeats)
            ExpandRepeat(doc, start, repeat);

        foreach (var (find, replace) in replacements)
            ReplaceInRange(doc, start, find, replace);

        // The captured fragment has no final paragraph mark (trimmed at capture),
        // so terminate its last paragraph; otherwise the next block's content
        // would merge into it.
        dynamic after = doc.Bookmarks["\\endofdoc"].Range;
        after.InsertParagraphAfter();
        ReleaseCom(after);
    }

    /// <summary>
    /// Fills a table with one row per record: the row carrying the {repeat …} marker is
    /// the template, a row is added per record — Word gives an added row the reference
    /// row's shading, borders and height — and each cell's content is copied over with
    /// its character formatting before the record's values are substituted in. The
    /// marker row is removed once its clones are in place.
    /// </summary>
    private static void ExpandRepeat(dynamic doc, int start, FragmentRowRepeat repeat)
    {
        var markerRow = FindMarkerRow(doc, start, repeat.Marker);
        if (markerRow is null) return;   // marker is not inside a table: nothing to repeat

        dynamic table = markerRow.Range.Tables[1];
        int cellCount = markerRow.Cells.Count;

        foreach (var values in repeat.Rows)
        {
            dynamic clone = table.Rows.Add(markerRow);      // inserted just above the marker row
            for (var c = 1; c <= cellCount; c++)
            {
                dynamic source = markerRow.Cells.Item(c).Range;
                source.MoveEnd(WdCharacter, -1);            // leave the end-of-cell mark behind
                dynamic target = clone.Cells.Item(c).Range;
                target.MoveEnd(WdCharacter, -1);
                target.FormattedText = source.FormattedText;
                ReleaseCom(source);
                ReleaseCom(target);

                dynamic cell = clone.Cells.Item(c);
                foreach (var (find, replace) in values)
                    ReplaceInCell(cell, find, replace);
                ReplaceInCell(cell, repeat.Marker, "");
                ReleaseCom(cell);
            }
            ReleaseCom(clone);
        }

        markerRow.Delete();
        ReleaseCom(markerRow);
        ReleaseCom(table);
    }

    /// <summary>The row of a table inside the just-inserted fragment that carries the marker.</summary>
    private static dynamic? FindMarkerRow(dynamic doc, int start, string marker)
    {
        dynamic scope = doc.Range(start, doc.Content.End);
        try
        {
            foreach (dynamic table in scope.Tables)
            {
                foreach (dynamic row in table.Rows)
                {
                    string text = row.Range.Text ?? "";
                    if (text.Contains(marker, StringComparison.Ordinal)) return row;
                    ReleaseCom(row);
                }
                ReleaseCom(table);
            }
            return null;
        }
        finally
        {
            ReleaseCom(scope);
        }
    }

    /// <summary>
    /// Literal find→replace confined to one table cell. Two Word quirks to respect:
    /// Find happily runs past the end of the range it was started on, so every hit is
    /// checked against the cell's own bounds; and it skips a match that fills the whole
    /// search range, so the range must be the entire cell — end-of-cell mark included —
    /// or a cell containing nothing but {Qty} never matches "{Qty}".
    /// </summary>
    private static void ReplaceInCell(dynamic cell, string find, string replace)
    {
        if (string.IsNullOrEmpty(find) || find.Length > 255) return;   // Word Find limit

        for (var guard = 0; guard < 50; guard++)
        {
            dynamic bounds = cell.Range;
            int cellStart = bounds.Start, cellEnd = bounds.End;
            ReleaseCom(bounds);

            dynamic search = cell.Range;
            dynamic finder = search.Find;
            finder.ClearFormatting();
            finder.Forward = true;
            finder.Wrap = 0;                        // wdFindStop
            finder.MatchCase = true;
            finder.MatchWildcards = false;
            bool found = finder.Execute(find);
            ReleaseCom(finder);

            var inside = found && (int)search.Start >= cellStart && (int)search.End <= cellEnd;
            if (!inside)
            {
                ReleaseCom(search);
                return;
            }

            search.Text = replace;                  // no length limit, unlike Find's ReplaceWith
            ReleaseCom(search);
            if (replace.Contains(find, StringComparison.Ordinal)) return;   // would find itself forever
        }
    }

    /// <summary>
    /// Flat OPC of a fragment .docx, read through Word and cached for the run so a
    /// loop over 500 records opens each fragment file once, not 500 times.
    /// </summary>
    private string ReadFragmentXml(string? docxPath)
    {
        if (string.IsNullOrWhiteSpace(docxPath)) return "";
        if (_fragmentXmlCache.TryGetValue(docxPath, out var cached)) return cached;

        var app = RequireApp();
        // Positional args: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
        // then optional up to Visible (12th) — kept invisible even when ShowWord is on.
        dynamic doc = app.Documents.Open(docxPath, false, true, false,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, false);
        try
        {
            int end = Math.Max(0, (int)doc.Content.End - 1);
            if (end == 0) return _fragmentXmlCache[docxPath] = "";

            dynamic range = doc.Range(0, end);
            string xml = range.WordOpenXML;
            ReleaseCom(range);
            return _fragmentXmlCache[docxPath] = xml;
        }
        finally
        {
            try { doc.Close(0); } catch { /* best effort */ }
            ReleaseCom(doc);
        }
    }

    /// <summary>
    /// Literal find→replace from <paramref name="start"/> to the end of the document.
    /// Word's Find spans formatting runs, so placeholders keep working even when
    /// Word split them into multiple runs; replacements inherit local formatting.
    /// </summary>
    private static void ReplaceInRange(dynamic doc, int start, string find, string replace)
    {
        if (string.IsNullOrEmpty(find) || find.Length > 255) return; // Word Find limit

        while (true)
        {
            dynamic search = doc.Range(start, doc.Content.End);
            dynamic finder = search.Find;
            finder.ClearFormatting();
            finder.Forward = true;
            finder.Wrap = 0; // wdFindStop
            finder.MatchCase = true;
            finder.MatchWildcards = false;
            bool found = finder.Execute(find);
            if (!found)
            {
                ReleaseCom(finder);
                ReleaseCom(search);
                break;
            }
            // After a successful Execute the range is the found text; assigning
            // Text replaces it (no length limits, unlike Find's ReplaceWith).
            search.Text = replace;
            start = (int)search.End;
            ReleaseCom(finder);
            ReleaseCom(search);
        }
    }

    public void AddTable(IReadOnlyList<string> headers, IReadOnlyList<string[]> rows, bool headerRow)
    {
        var doc = RequireDoc();
        var colCount = Math.Max(1, headers.Count);
        var rowCount = rows.Count + (headerRow ? 1 : 0);
        if (rowCount == 0) rowCount = 1;

        dynamic range = doc.Bookmarks["\\endofdoc"].Range;
        dynamic table = doc.Tables.Add(range, rowCount, colCount);
        table.Borders.InsideLineStyle = WdLineStyleSingle;
        table.Borders.OutsideLineStyle = WdLineStyleSingle;
        table.Range.Style = Interop.WordStyleIds.Normal;

        var r = 1;
        if (headerRow)
        {
            for (var c = 0; c < colCount; c++)
                table.Cell(1, c + 1).Range.Text = headers[c];
            table.Rows[1].Range.Font.Bold = 1;
            table.Rows[1].Shading.BackgroundPatternColor = 0xE8E8E8; // light gray (BGR)
            r = 2;
        }
        foreach (var row in rows)
        {
            for (var c = 0; c < colCount && c < row.Length; c++)
                table.Cell(r, c + 1).Range.Text = row[c];
            r++;
        }

        // Move the insertion point past the table so following content lands below it
        dynamic after = doc.Bookmarks["\\endofdoc"].Range;
        after.InsertParagraphAfter();
        ReleaseCom(after);
        ReleaseCom(table);
        ReleaseCom(range);
    }

    public void PageBreak()
    {
        var doc = RequireDoc();
        dynamic range = doc.Bookmarks["\\endofdoc"].Range;
        range.Collapse(WdCollapseEnd);
        range.InsertBreak(WdPageBreak);
        ReleaseCom(range);
    }

    public string EndDocument(string fileNameWithoutExtension)
    {
        var doc = RequireDoc();
        var name = string.IsNullOrWhiteSpace(fileNameWithoutExtension) ? "Document" : fileNameWithoutExtension;
        var path = Path.Combine(OutputFolder, name + ".docx");
        doc.SaveAs2(path, WdFormatXMLDocument);
        doc.Close(0); // wdDoNotSaveChanges — already saved
        ReleaseCom(doc);
        _doc = null;
        return path;
    }

    public void End()
    {
        if (_doc is not null)
        {
            try { _doc.Close(0); } catch { /* best effort */ }
            ReleaseCom(_doc);
            _doc = null;
        }
        if (_app is not null)
        {
            try { _app.Quit(0); } catch { /* best effort */ }
            ReleaseCom(_app);
            _app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    public void Dispose() => End();


    private dynamic RequireApp() =>
        _app ?? throw new InvalidOperationException("Word writer not started.");

    private dynamic RequireDoc() =>
        _doc ?? throw new InvalidOperationException(
            "No document is open — content blocks must be inside a 'new document' block.");

    private static void ReleaseCom(object? o)
    {
        if (o is not null && Marshal.IsComObject(o))
            Marshal.ReleaseComObject(o);
    }
}
