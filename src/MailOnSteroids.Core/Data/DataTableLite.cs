namespace MailOnSteroids.Core.Data;

/// <summary>
/// Lightweight in-memory table shared by every data source provider.
/// Column lookup is case-insensitive.
/// </summary>
public sealed class DataTableLite
{
    private readonly Dictionary<string, int> _columnIndex = new(StringComparer.OrdinalIgnoreCase);

    public string Name { get; set; } = "";
    public List<string> Columns { get; } = [];
    public List<object?[]> Rows { get; } = [];

    public DataTableLite(IEnumerable<string> columns)
    {
        foreach (var c in columns)
        {
            var name = string.IsNullOrWhiteSpace(c) ? $"Column{Columns.Count + 1}" : c.Trim();
            // De-duplicate repeated header names
            var unique = name;
            var n = 2;
            while (_columnIndex.ContainsKey(unique)) unique = $"{name}_{n++}";
            _columnIndex[unique] = Columns.Count;
            Columns.Add(unique);
        }
    }

    public int ColumnIndex(string name) =>
        _columnIndex.TryGetValue(name?.Trim() ?? "", out var i) ? i : -1;

    public bool TryGetValue(int rowIndex, string column, out object? value)
    {
        var i = ColumnIndex(column);
        if (i < 0 || rowIndex < 0 || rowIndex >= Rows.Count) { value = null; return false; }
        value = Rows[rowIndex][i];
        return true;
    }

    public void AddRow(object?[] values)
    {
        if (values.Length != Columns.Count)
            Array.Resize(ref values, Columns.Count);
        Rows.Add(values);
    }

    /// <summary>New table with the same columns containing only rows matching the predicate.</summary>
    public DataTableLite Where(Func<int, bool> rowPredicate, string newName)
    {
        var result = new DataTableLite(Columns) { Name = newName };
        for (var i = 0; i < Rows.Count; i++)
            if (rowPredicate(i))
                result.Rows.Add(Rows[i]);
        return result;
    }
}
