using System.Text;

namespace MergeOnSteroids.Core.Runtime;

/// <summary>
/// Renders documents as plain text files — lets you test a program quickly
/// (and headlessly) without launching Word.
/// </summary>
public sealed class PreviewWriter : IDocumentWriter
{
    private RunOptions _options = new();
    private StringBuilder? _doc;
    private int _docCounter;

    /// <summary>(fileName, content) for every document produced in the run.</summary>
    public List<(string Path, string Content)> Documents { get; } = [];

    public bool InDocument => _doc is not null;

    public void Begin(RunOptions options)
    {
        _options = options;
        Directory.CreateDirectory(options.OutputFolder);
    }

    public void BeginDocument(string? templatePath)
    {
        _doc = new StringBuilder();
        _docCounter++;
        if (!string.IsNullOrWhiteSpace(templatePath))
            _doc.AppendLine($"[based on template: {templatePath}]").AppendLine();
    }

    public void AddParagraph(string text, string style, bool bold, bool italic)
    {
        var d = RequireDoc();
        var prefix = style switch
        {
            "Title" => "======== ",
            "Subtitle" => "-------- ",
            "Heading 1" => "# ",
            "Heading 2" => "## ",
            "Heading 3" => "### ",
            "Quote" => "> ",
            "List Bullet" => "  • ",
            _ => ""
        };
        var deco = (bold ? "**" : "") + (italic ? "_" : "");
        var decoClose = (italic ? "_" : "") + (bold ? "**" : "");
        d.AppendLine($"{prefix}{deco}{text}{decoClose}");
    }

    public void AddFragment(FragmentContent fragment, string plainText,
        IReadOnlyList<KeyValuePair<string, string>> replacements)
    {
        var d = RequireDoc();
        var text = plainText ?? "";
        foreach (var (find, replace) in replacements)
            text = text.Replace(find, replace);
        d.AppendLine(text);
    }

    public void AddTable(IReadOnlyList<string> headers, IReadOnlyList<string[]> rows, bool headerRow)
    {
        var d = RequireDoc();
        var colCount = headers.Count;
        var widths = new int[colCount];
        for (var c = 0; c < colCount; c++)
        {
            widths[c] = headerRow ? headers[c].Length : 0;
            foreach (var row in rows)
                widths[c] = Math.Max(widths[c], (c < row.Length ? row[c] : "").Length);
        }

        void WriteRow(IReadOnlyList<string> cells)
        {
            d.Append('|');
            for (var c = 0; c < colCount; c++)
                d.Append(' ').Append((c < cells.Count ? cells[c] : "").PadRight(widths[c])).Append(" |");
            d.AppendLine();
        }

        if (headerRow)
        {
            WriteRow(headers);
            d.Append('|');
            for (var c = 0; c < colCount; c++)
                d.Append(new string('-', widths[c] + 2)).Append('|');
            d.AppendLine();
        }
        foreach (var row in rows) WriteRow(row);
    }

    public void PageBreak() => RequireDoc().AppendLine().AppendLine("· · · · · · · · page break · · · · · · · ·").AppendLine();

    public string EndDocument(string fileNameWithoutExtension)
    {
        var d = RequireDoc();
        var name = string.IsNullOrWhiteSpace(fileNameWithoutExtension)
            ? $"Document_{_docCounter}"
            : fileNameWithoutExtension;
        var path = Path.Combine(_options.OutputFolder, name + ".txt");
        File.WriteAllText(path, d.ToString(), Encoding.UTF8);
        Documents.Add((path, d.ToString()));
        _doc = null;
        return path;
    }

    public void End() { }

    public void Dispose() { }

    private StringBuilder RequireDoc() =>
        _doc ?? throw new InvalidOperationException(
            "No document is open — content blocks must be inside a 'new document' block.");
}
