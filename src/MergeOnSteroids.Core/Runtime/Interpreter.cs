using System.IO.Compression;
using MergeOnSteroids.Core.Blocks;
using MergeOnSteroids.Core.Data;
using MergeOnSteroids.Core.Expressions;

namespace MergeOnSteroids.Core.Runtime;

public sealed class MosRuntimeException(string message) : Exception(message);

public sealed class RunResult
{
    public List<string> OutputFiles { get; } = [];
    public int Warnings { get; set; }
    public bool Succeeded { get; set; }
    public string? Error { get; set; }
}

/// <summary>Walks the block tree and produces documents through an <see cref="IDocumentWriter"/>.</summary>
public sealed class Interpreter
{
    private readonly ExpressionCompiler _compiler = new();
    private readonly EvalContext _ctx;
    private readonly IDocumentWriter _writer;
    private readonly RunOptions _options;
    private readonly Action<string> _log;
    private readonly RunResult _result = new();
    private readonly Dictionary<string, DataTableLite> _fileCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _fragmentTextCache = new(StringComparer.OrdinalIgnoreCase);
    private int _documentCount;

    public Interpreter(IDocumentWriter writer, RunOptions options, Action<string>? log = null)
    {
        _writer = writer;
        _options = options;
        _log = log ?? (_ => { });
        _ctx = new EvalContext();
    }

    public RunResult Run(ProgramModel program)
    {
        try
        {
            _log($"Run started — output folder: {_options.OutputFolder}");
            _writer.Begin(_options);
            ExecuteList(program.Blocks);
            _writer.End();
            _result.Succeeded = true;
            _log($"Run finished — {_result.OutputFiles.Count} output file(s), {_result.Warnings} warning(s).");
        }
        catch (OperationCanceledException)
        {
            _result.Error = "Run cancelled.";
            _log("Run cancelled by user.");
            SafeEndWriter();
        }
        catch (Exception ex)
        {
            _result.Error = ex.Message;
            _log($"ERROR: {ex.Message}");
            SafeEndWriter();
        }
        return _result;
    }

    private void SafeEndWriter()
    {
        try { _writer.End(); } catch { /* already failing */ }
    }

    // ------------------------------------------------------------- execution

    private void ExecuteList(IEnumerable<Block> blocks)
    {
        foreach (var block in blocks)
        {
            _options.Cancellation.ThrowIfCancellationRequested();
            Execute(block);
        }
    }

    private void Execute(Block block)
    {
        switch (block)
        {
            case CsvSourceBlock b: ExecuteCsvSource(b); break;
            case ExcelSourceBlock b: ExecuteExcelSource(b); break;
            case DatabaseSourceBlock b: ExecuteDatabaseSource(b); break;
            case FilterSourceBlock b: ExecuteFilterSource(b); break;
            case ForEachBlock b: ExecuteForEach(b); break;
            case IfBlock b: ExecuteIf(b); break;
            case SwitchBlock b: ExecuteSwitch(b); break;
            case CaseBlock: Warn("A 'case' block only works inside a 'switch' block — skipped."); break;
            case SetVariableBlock b: ExecuteSetVariable(b); break;
            case MakeDirectoryBlock b: ExecuteMakeDirectory(b); break;
            case ZipDirectoryBlock b: ExecuteZipDirectory(b); break;
            case NewDocumentBlock b: ExecuteNewDocument(b); break;
            case ParagraphBlock b: ExecuteParagraph(b); break;
            case WordFragmentBlock b: ExecuteWordFragment(b); break;
            case TableBlock b: ExecuteTable(b); break;
            case PageBreakBlock: _writer.PageBreak(); break;
            default:
                throw new MosRuntimeException($"Unknown block type '{block.GetType().Name}'.");
        }
    }

    private void ExecuteCsvSource(CsvSourceBlock b)
    {
        RequireSourceName(b);
        var path = _options.ResolvePath(Interpolate(b.FilePath));
        var key = $"csv|{path}|{b.HeaderRow}";
        if (!_fileCache.TryGetValue(key, out var table))
        {
            if (!File.Exists(path))
                throw new MosRuntimeException($"CSV file not found: {path} (source '{b.Name}').");
            try
            {
                table = DataSourceLoaders.LoadCsv(path, b.Name, b.HeaderRow);
            }
            catch (InvalidOperationException ex)
            {
                throw new MosRuntimeException($"Source '{b.Name}': {ex.Message}");
            }
            _fileCache[key] = table;
            _log($"Source '{b.Name}': {table.Rows.Count} row(s) from {Path.GetFileName(path)} " +
                 $"({(b.HeaderRow > 0 ? $"headers in row {b.HeaderRow}" : "no header row")}) " +
                 $"[{string.Join(", ", table.Columns)}]");
        }
        _ctx.RegisterSource(b.Name, table);
    }

    private void ExecuteExcelSource(ExcelSourceBlock b)
    {
        RequireSourceName(b);
        var path = _options.ResolvePath(Interpolate(b.FilePath));
        var key = $"xlsx|{path}|{b.SheetName}|{b.HeaderRow}";
        if (!_fileCache.TryGetValue(key, out var table))
        {
            if (!File.Exists(path))
                throw new MosRuntimeException($"Excel file not found: {path} (source '{b.Name}').");
            try
            {
                table = DataSourceLoaders.LoadExcel(path, b.SheetName, b.Name, b.HeaderRow);
            }
            catch (InvalidOperationException ex)
            {
                throw new MosRuntimeException($"Source '{b.Name}': {ex.Message}");
            }
            _fileCache[key] = table;
            _log($"Source '{b.Name}': {table.Rows.Count} row(s) from {Path.GetFileName(path)} " +
                 $"({(b.HeaderRow > 0 ? $"headers in row {b.HeaderRow}" : "no header row")}) " +
                 $"[{string.Join(", ", table.Columns)}]");
        }
        _ctx.RegisterSource(b.Name, table);
    }

    private void ExecuteDatabaseSource(DatabaseSourceBlock b)
    {
        RequireSourceName(b);
        // Query supports {expression} interpolation so secondary lookups can be
        // driven by the current record. NOTE: values are inlined into the SQL text.
        var query = Interpolate(b.Query);
        var table = DataSourceLoaders.LoadDatabase(b.Provider, b.ConnectionString, query, b.Name);
        _log($"Source '{b.Name}': {table.Rows.Count} row(s) from {b.Provider} query.");
        _ctx.RegisterSource(b.Name, table);
    }

    private void ExecuteFilterSource(FilterSourceBlock b)
    {
        RequireSourceName(b);
        var source = ResolveSourceOrThrow(b.SourceName, b.DisplayName);
        if (string.IsNullOrWhiteSpace(b.Condition))
            throw new MosRuntimeException($"Filter source '{b.Name}' has no condition.");

        var predicate = _compiler.Compile(b.Condition);
        var filtered = source.Where(i =>
        {
            using var scope = _ctx.PushRowScope(source, i);
            return Values.Truthy(predicate.Eval(_ctx));
        }, b.Name);

        _ctx.RegisterSource(b.Name, filtered);
    }

    private void ExecuteForEach(ForEachBlock b)
    {
        var source = ResolveSourceOrThrow(b.SourceName, b.DisplayName);
        var alias = string.IsNullOrWhiteSpace(b.Alias) ? "rec" : b.Alias.Trim();

        for (var i = 0; i < source.Rows.Count; i++)
        {
            _options.Cancellation.ThrowIfCancellationRequested();
            using var scope = _ctx.PushAliasScope(alias, source, i);
            ExecuteList(b.Children);
        }
    }

    private void ExecuteIf(IfBlock b)
    {
        if (string.IsNullOrWhiteSpace(b.Condition))
            throw new MosRuntimeException("An 'if' block has no condition.");
        var value = Values.Truthy(_compiler.Compile(b.Condition).Eval(_ctx));
        ExecuteList(value ? b.Children : b.Else);
    }

    private void ExecuteSwitch(SwitchBlock b)
    {
        if (string.IsNullOrWhiteSpace(b.ValueExpression))
            throw new MosRuntimeException("A 'switch' block has no value to switch on.");

        var value = _compiler.Compile(b.ValueExpression).Eval(_ctx);

        foreach (var child in b.Children)
        {
            _options.Cancellation.ThrowIfCancellationRequested();
            if (child is not CaseBlock c)
            {
                Warn($"A 'switch' block holds only 'case' blocks — '{child.DisplayName}' was skipped.");
                continue;
            }
            if (!CaseMatches(c, value)) continue;
            ExecuteList(c.Children);
            return;
        }

        ExecuteList(b.Else);
    }

    private bool CaseMatches(CaseBlock c, object? value)
    {
        foreach (var candidate in SplitTopLevel(c.MatchValues))
            if (Values.AreEqual(value, _compiler.Compile(candidate).Eval(_ctx)))
                return true;
        return false;
    }

    /// <summary>
    /// Splits a case's values on the commas that separate them — the ones inside
    /// string literals or brackets belong to an expression, not to the list.
    /// </summary>
    private static List<string> SplitTopLevel(string spec)
    {
        var parts = new List<string>();
        if (string.IsNullOrWhiteSpace(spec)) return parts;

        var depth = 0;
        char? quote = null;
        var start = 0;
        for (var i = 0; i < spec.Length; i++)
        {
            var ch = spec[i];
            if (quote is { } q)
            {
                if (ch == q) quote = null;
            }
            else switch (ch)
            {
                case '"' or '\'' or '“' or '‘': quote = ch switch { '“' => '”', '‘' => '’', _ => ch }; break;
                case '(' or '[': depth++; break;
                case ')' or ']': depth--; break;
                case ',' when depth == 0:
                    parts.Add(spec[start..i]);
                    start = i + 1;
                    break;
            }
        }
        parts.Add(spec[start..]);
        return parts.Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim()).ToList();
    }

    private void ExecuteSetVariable(SetVariableBlock b)
    {
        if (string.IsNullOrWhiteSpace(b.VariableName))
            throw new MosRuntimeException("A 'set variable' block has no variable name.");
        var value = string.IsNullOrWhiteSpace(b.ValueExpression)
            ? null
            : _compiler.Compile(b.ValueExpression).Eval(_ctx);
        _ctx.SetVariable(b.VariableName.Trim(), value);
    }

    // ------------------------------------------------------------- folders

    private void ExecuteMakeDirectory(MakeDirectoryBlock b)
    {
        var folder = EnterFolder(b);
        _log($"Folder: {folder}");
        using (UseFolder(folder)) ExecuteList(b.Children);
    }

    private void ExecuteZipDirectory(ZipDirectoryBlock b)
    {
        var folder = EnterFolder(b);
        var zipPath = folder + ".zip";

        using (UseFolder(folder)) ExecuteList(b.Children);

        if (File.Exists(zipPath)) File.Delete(zipPath);
        ZipFile.CreateFromDirectory(folder, zipPath, CompressionLevel.Optimal, includeBaseDirectory: false);

        var zipped = _result.OutputFiles
            .Where(p => p.StartsWith(folder + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (!b.KeepFolder)
        {
            Directory.Delete(folder, recursive: true);
            // those files only exist inside the archive now
            foreach (var p in zipped) _result.OutputFiles.Remove(p);
        }

        _result.OutputFiles.Add(zipPath);
        _log($"Zipped {zipped.Count} document(s) into {zipPath}" +
             (b.KeepFolder ? " (folder kept)" : ""));
    }

    /// <summary>Creates the folder a folder block asks for, inside the current one.</summary>
    private string EnterFolder(FolderBlockBase b)
    {
        if (_writer.InDocument)
            Warn($"A '{b.DisplayName}' block sits inside an open document — the document is saved " +
                 "where its 'new document' block started, not in this folder.");

        var name = SanitizeRelativePath(Interpolate(b.FolderName));
        if (name.Length == 0)
            throw new MosRuntimeException($"A '{b.DisplayName}' block has no folder name.");

        var folder = Path.GetFullPath(Path.Combine(_writer.OutputFolder, name));
        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(_options.OutputFolder));
        if (!folder.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new MosRuntimeException(
                $"A '{b.DisplayName}' block would write outside the output folder: {folder}");

        Directory.CreateDirectory(folder);
        return folder;
    }

    /// <summary>Saves documents into <paramref name="folder"/> until disposed.</summary>
    private IDisposable UseFolder(string folder)
    {
        var previous = _writer.OutputFolder;
        _writer.OutputFolder = folder;
        return new RestoreFolder(_writer, previous);
    }

    private sealed class RestoreFolder(IDocumentWriter writer, string previous) : IDisposable
    {
        public void Dispose() => writer.OutputFolder = previous;
    }

    /// <summary>
    /// Turns an interpolated folder name into a safe relative path: keeps the "/"
    /// levels, drops anything that would climb out, and cleans each segment.
    /// </summary>
    private static string SanitizeRelativePath(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var segments = name
            .Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(s => s is not ("." or ".."))
            .Select(s => new string(s.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim())
            .Where(s => s.Length > 0);
        return string.Join(Path.DirectorySeparatorChar, segments);
    }

    private void ExecuteNewDocument(NewDocumentBlock b)
    {
        var template = string.IsNullOrWhiteSpace(b.TemplatePath) ? null : Interpolate(b.TemplatePath);
        _writer.BeginDocument(template);
        _documentCount++;

        ExecuteList(b.Children);

        var rawName = Interpolate(b.FileNameTemplate);
        if (string.IsNullOrWhiteSpace(rawName)) rawName = $"Document_{_documentCount}";
        var path = _writer.EndDocument(SanitizeFileName(rawName));
        _result.OutputFiles.Add(path);
        _log($"Document saved: {path}");
    }

    private void ExecuteParagraph(ParagraphBlock b)
    {
        var text = Interpolate(b.TextTemplate);
        _writer.AddParagraph(text, b.Style, b.Bold, b.Italic);
    }

    private void ExecuteWordFragment(WordFragmentBlock b)
    {
        if (!b.HasContent)
        {
            Warn("A 'Word paragraphs' block has no content yet (use 'Edit in Word') — skipped.");
            return;
        }

        FragmentContent content;
        string plainText;
        IReadOnlyList<Fragments.FragmentRepeatRow> repeatRows;
        if (!string.IsNullOrWhiteSpace(b.FragmentFile))
        {
            var path = _options.ResolvePath(b.FragmentFile);
            if (!File.Exists(path))
                throw new MosRuntimeException(
                    $"Fragment file not found: {path} (referenced by a 'Word paragraphs' block).");
            content = new FragmentContent(path, null);
            plainText = FragmentText(path);
            repeatRows = Fragments.FragmentRepeats.Read(path);
        }
        else
        {
            content = new FragmentContent(null, b.LegacyFragmentXml);
            plainText = Interop.WordFragmentText.Extract(b.LegacyFragmentXml ?? "");
            repeatRows = [];   // pre-sidecar fragments predate repeating rows
        }

        var repeats = repeatRows.Select(BuildRepeat).Where(r => r is not null).ToList()!;

        // Evaluate each {placeholder} found in the fragment's text; the raw
        // placeholder (exactly as it appears in the document, smart quotes and
        // all) becomes the literal find-text for substitution. Placeholders that
        // only ever appear inside a repeated row are skipped here — they belong to
        // a record, and were already resolved per clone.
        var repeatedText = string.Concat(repeatRows.Select(r => r.RowText));
        var replacements = new List<KeyValuePair<string, string>>();
        foreach (var (expression, raw) in TemplateEngine.ExtractPlaceholders(plainText)
                     .DistinctBy(p => p.RawPlaceholder))
        {
            if (repeatRows.Any(r => r.Marker == raw)) continue;                 // the marker itself
            if (Occurrences(plainText, raw) <= Occurrences(repeatedText, raw)) continue;

            string value;
            try
            {
                value = Values.ToDisplayString(_compiler.Compile(expression).Eval(_ctx));
            }
            catch (MosExpressionException ex)
            {
                value = "{!" + ex.Message + "}";
                Warn($"In '{raw}': {ex.Message}");
            }
            replacements.Add(new KeyValuePair<string, string>(raw, value));
        }

        _writer.AddFragment(content, plainText, replacements, repeats!);
    }

    /// <summary>
    /// Works out what one repeated row becomes: its placeholders evaluated once per
    /// record of the source it names, each set applied to one clone of the row.
    /// </summary>
    private FragmentRowRepeat? BuildRepeat(Fragments.FragmentRepeatRow row)
    {
        var source = _ctx.ResolveSource(row.SourceName.Trim());
        if (source is null)
        {
            Warn($"'{row.Marker}' in a 'Word paragraphs' block: data source '{row.SourceName}' " +
                 "is not defined at this point — the row was dropped.");
            return new FragmentRowRepeat(row.Marker, row.CellTemplates, []);
        }

        var placeholders = TemplateEngine.ExtractPlaceholders(row.RowText)
            .DistinctBy(p => p.RawPlaceholder)
            .Where(p => p.RawPlaceholder != row.Marker)
            .ToList();

        var rows = new List<IReadOnlyList<KeyValuePair<string, string>>>(source.Rows.Count);
        for (var i = 0; i < source.Rows.Count; i++)
        {
            _options.Cancellation.ThrowIfCancellationRequested();
            using var scope = _ctx.PushRowScope(source, i);

            var values = new List<KeyValuePair<string, string>>(placeholders.Count);
            foreach (var (expression, raw) in placeholders)
            {
                string value;
                try
                {
                    value = Values.ToDisplayString(_compiler.Compile(expression).Eval(_ctx));
                }
                catch (MosExpressionException ex)
                {
                    value = "{!" + ex.Message + "}";
                    Warn($"In '{raw}' (row {i + 1} of '{row.SourceName}'): {ex.Message}");
                }
                values.Add(new KeyValuePair<string, string>(raw, value));
            }
            rows.Add(values);
        }

        _log($"Fragment table: {rows.Count} row(s) from '{row.SourceName}'.");
        return new FragmentRowRepeat(row.Marker, row.CellTemplates, rows);
    }

    private static int Occurrences(string text, string value)
    {
        if (string.IsNullOrEmpty(value)) return 0;
        int count = 0, at = 0;
        while ((at = text.IndexOf(value, at, StringComparison.Ordinal)) >= 0) { count++; at += value.Length; }
        return count;
    }

    /// <summary>Fragment text read straight from the .docx (no Word needed), cached per run.</summary>
    private string FragmentText(string docxPath)
    {
        if (_fragmentTextCache.TryGetValue(docxPath, out var cached)) return cached;
        try
        {
            return _fragmentTextCache[docxPath] = Fragments.FragmentFiles.ReadPlainText(docxPath);
        }
        catch (Exception ex)
        {
            Warn($"Could not read fragment '{Path.GetFileName(docxPath)}': {ex.Message}");
            return _fragmentTextCache[docxPath] = "";
        }
    }

    private void ExecuteTable(TableBlock b)
    {
        var source = ResolveSourceOrThrow(b.SourceName, b.DisplayName);
        var columns = ParseColumnsSpec(b.ColumnsSpec, source);

        var headers = columns.Select(c => c.Header).ToList();
        var rows = new List<string[]>(source.Rows.Count);
        for (var i = 0; i < source.Rows.Count; i++)
        {
            using var scope = _ctx.PushRowScope(source, i);
            var cells = new string[columns.Count];
            for (var c = 0; c < columns.Count; c++)
            {
                try
                {
                    cells[c] = Values.ToDisplayString(columns[c].Expr.Eval(_ctx));
                }
                catch (MosExpressionException ex)
                {
                    cells[c] = "{!" + ex.Message + "}";
                    Warn($"Table column '{columns[c].Header}': {ex.Message}");
                }
            }
            rows.Add(cells);
        }
        _writer.AddTable(headers, rows, b.HeaderRow);
    }

    private List<(string Header, ExprNode Expr)> ParseColumnsSpec(string spec, DataTableLite source)
    {
        var result = new List<(string, ExprNode)>();
        if (string.IsNullOrWhiteSpace(spec))
        {
            foreach (var col in source.Columns)
                result.Add((col, _compiler.Compile($"[{col}]")));
            return result;
        }

        foreach (var part in spec.Split('|'))
        {
            var piece = part.Trim();
            if (piece.Length == 0) continue;
            var colon = piece.IndexOf(':');
            string header, expr;
            if (colon > 0)
            {
                header = piece[..colon].Trim();
                expr = piece[(colon + 1)..].Trim();
            }
            else
            {
                header = piece;
                expr = piece.Contains(' ') ? $"[{piece}]" : piece;
            }
            result.Add((header, _compiler.Compile(expr)));
        }
        if (result.Count == 0)
            throw new MosRuntimeException("Table columns spec produced no columns.");
        return result;
    }

    // -------------------------------------------------------------- helpers

    private static void RequireSourceName(SourceBlockBase b)
    {
        if (string.IsNullOrWhiteSpace(b.Name))
            throw new MosRuntimeException($"A '{b.DisplayName}' block needs a name.");
    }

    private DataTableLite ResolveSourceOrThrow(string name, string blockName)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new MosRuntimeException($"A '{blockName}' block has no data source selected.");
        return _ctx.ResolveSource(name.Trim())
            ?? throw new MosRuntimeException(
                $"Data source '{name}' is not defined at this point (check block order).");
    }

    private string Interpolate(string template) =>
        TemplateEngine.Interpolate(template ?? "", _compiler, _ctx, Warn);

    private void Warn(string message)
    {
        _result.Warnings++;
        _log($"warning: {message}");
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        var clean = new string(chars).Trim();
        return clean.Length == 0 ? "Document" : clean;
    }

    // ------------------------------------------------------- evaluation ctx

    private sealed class EvalContext : IEvalContext
    {
        private readonly Dictionary<string, DataTableLite> _sources = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, object?> _variables = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<(string? Alias, DataTableLite Table, int RowIndex)> _rowScopes = [];

        public void RegisterSource(string name, DataTableLite table) => _sources[name.Trim()] = table;

        public void SetVariable(string name, object? value) => _variables[name] = value;

        public bool TryGetVariable(string name, out object? value) => _variables.TryGetValue(name, out value);

        public bool TryGetBareField(string field, out object? value)
        {
            for (var i = _rowScopes.Count - 1; i >= 0; i--)
            {
                var (_, table, rowIndex) = _rowScopes[i];
                if (table.TryGetValue(rowIndex, field, out value)) return true;
            }
            value = null;
            return false;
        }

        public bool TryGetAliasField(string alias, string field, out object? value)
        {
            for (var i = _rowScopes.Count - 1; i >= 0; i--)
            {
                var (a, table, rowIndex) = _rowScopes[i];
                if (a is not null && a.Equals(alias, StringComparison.OrdinalIgnoreCase))
                    return table.TryGetValue(rowIndex, field, out value);
            }
            value = null;
            return false;
        }

        public bool IsAlias(string name) =>
            _rowScopes.Any(s => s.Alias?.Equals(name, StringComparison.OrdinalIgnoreCase) == true);

        public DataTableLite? ResolveSource(string name) =>
            _sources.TryGetValue(name.Trim(), out var t) ? t : null;

        public IDisposable PushRowScope(DataTableLite table, int rowIndex) => Push(null, table, rowIndex);

        public IDisposable PushAliasScope(string alias, DataTableLite table, int rowIndex) =>
            Push(alias, table, rowIndex);

        private IDisposable Push(string? alias, DataTableLite table, int rowIndex)
        {
            _rowScopes.Add((alias, table, rowIndex));
            return new PopOnDispose(_rowScopes);
        }

        private sealed class PopOnDispose(List<(string?, DataTableLite, int)> scopes) : IDisposable
        {
            public void Dispose() => scopes.RemoveAt(scopes.Count - 1);
        }
    }
}
