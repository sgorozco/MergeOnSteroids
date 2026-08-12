namespace MergeOnSteroids.Core.Data;

/// <summary>One column of a file data source, as the editor shows it on the block.</summary>
public sealed record SourceColumnInfo(string Name, string? SampleText)
{
    /// <summary>
    /// How this column is written inside an expression: bare when the name is a
    /// plain identifier, [bracketed] when it has spaces, punctuation, starts with
    /// a digit, or collides with a keyword.
    /// </summary>
    public string Reference => NeedsBrackets(Name) ? $"[{Name}]" : Name;

    private static readonly string[] Keywords = ["AND", "OR", "NOT", "TRUE", "FALSE", "NULL"];

    private static bool NeedsBrackets(string name)
    {
        if (string.IsNullOrEmpty(name)) return true;
        if (!char.IsLetter(name[0]) && name[0] != '_') return true;
        foreach (var c in name)
            if (!char.IsLetterOrDigit(c) && c != '_') return true;
        return Keywords.Contains(name, StringComparer.OrdinalIgnoreCase);
    }
}

/// <summary>
/// What a data file looks like, read without loading it into a table: the editor
/// uses it to show the column names (and, for a workbook, the sheet list) on the block.
/// </summary>
public sealed record SourceSchema
{
    /// <summary>Worksheet name, or the file name for a CSV.</summary>
    public required string Label { get; init; }

    public required IReadOnlyList<SourceColumnInfo> Columns { get; init; }

    /// <summary>The row the column names actually came from; 0 when there is no header row.</summary>
    public required int HeaderRow { get; init; }

    /// <summary>Data rows, or -1 when the file was too big to count during a peek.</summary>
    public required int RowCount { get; init; }

    /// <summary>Sheet names of the workbook — empty for sources that have no sheets.</summary>
    public IReadOnlyList<string> SheetNames { get; init; } = [];

    /// <summary>Anything else worth showing, e.g. the delimiter detected in a CSV.</summary>
    public string? Note { get; init; }

    /// <summary>One-line summary shown under the block.</summary>
    public string Summary
    {
        get
        {
            var rows = RowCount < 0 ? "many rows" : Plural(RowCount, "row");
            var text = $"{Label} — {Plural(Columns.Count, "column")}, {rows}" +
                       (HeaderRow > 0 ? $", headers in row {HeaderRow}" : ", no header row");
            return Note is { Length: > 0 } ? $"{text}, {Note}" : text;
        }
    }

    private static string Plural(int n, string noun) => $"{n} {noun}{(n == 1 ? "" : "s")}";
}
