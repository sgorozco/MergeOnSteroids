using System.Data.Common;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;

namespace MergeOnSteroids.Core.Data;

/// <summary>Loads external data (CSV, Excel, databases) into <see cref="DataTableLite"/>.</summary>
public static class DataSourceLoaders
{
    // ------------------------------------------------------------------ CSV

    /// <summary>
    /// Reads a delimited text file into a table.
    /// <paramref name="headerRow"/> is the 1-based line holding the column names
    /// (anything above it is skipped); 0 means the file has no header line and the
    /// columns are then called A, B, C…
    /// </summary>
    public static DataTableLite LoadCsv(string path, string name, int headerRow = 1)
    {
        var records = ParseCsv(ReadText(path), out _);
        var layout = ReadCsvLayout(records, headerRow);

        var table = new DataTableLite(layout.Headers) { Name = name };
        for (var i = layout.FirstDataRecord; i < records.Count; i++)
        {
            var rec = records[i];
            if (IsBlank(rec)) continue;                    // blank / trailing lines
            var row = new object?[table.Columns.Count];
            for (var c = 0; c < table.Columns.Count && c < rec.Count; c++)
                row[c] = InferValue(rec[c]);
            table.AddRow(row);
        }
        return table;
    }

    /// <summary>
    /// Column names of a delimited file, read from at most the first few megabytes —
    /// what the editor needs to show the columns on the block.
    /// </summary>
    public static SourceSchema PeekCsv(string path, int headerRow = 1)
    {
        const int peekLimit = 4 * 1024 * 1024;
        var text = ReadText(path, peekLimit, out var complete);
        var records = ParseCsv(text, out var delimiter);
        var layout = ReadCsvLayout(records, headerRow);

        var sample = records.Skip(layout.FirstDataRecord).FirstOrDefault(r => !IsBlank(r));
        var columns = layout.Headers
            .Select((h, i) => new SourceColumnInfo(
                h, sample is not null && i < sample.Count && sample[i].Trim() is { Length: > 0 } s ? s : null))
            .ToList();

        var rowCount = complete
            ? records.Skip(layout.FirstDataRecord).Count(r => !IsBlank(r))
            : -1;

        return new SourceSchema
        {
            Label = Path.GetFileName(path),
            Columns = columns,
            HeaderRow = layout.HeaderRow,
            RowCount = rowCount,
            Note = delimiter switch
            {
                '\t' => "tab separated",
                ';' => "semicolon separated",
                _ => "comma separated"
            }
        };
    }

    private readonly record struct CsvLayout(List<string> Headers, int HeaderRow, int FirstDataRecord);

    /// <summary>Works out which record holds the column names and where the data starts.</summary>
    private static CsvLayout ReadCsvLayout(List<List<string>> records, int headerRow)
    {
        var firstUsed = records.FindIndex(r => !IsBlank(r));
        if (firstUsed < 0) return new CsvLayout([], headerRow > 0 ? headerRow : 0, records.Count);

        var width = records.Max(r => r.Count);

        if (headerRow <= 0)
        {
            // No header line: name the columns A, B, C… and keep every line as data.
            var letters = Enumerable.Range(1, width).Select(ColumnLetter).ToList();
            return new CsvLayout(letters, 0, firstUsed);
        }

        if (headerRow - 1 > records.FindLastIndex(r => !IsBlank(r)))
            throw new InvalidOperationException(
                $"Header row {headerRow} is past the last line with data (line {records.FindLastIndex(r => !IsBlank(r)) + 1}).");

        // Like a worksheet: "row 1" on a file that starts with blank lines simply
        // means its first line that has anything on it.
        var index = Math.Max(headerRow - 1, firstUsed);
        var header = records[index];
        var headers = Enumerable.Range(0, width)
            .Select(c => c < header.Count && header[c].Trim() is { Length: > 0 } text ? text : ColumnLetter(c + 1))
            .ToList();

        return new CsvLayout(headers, index + 1, index + 1);
    }

    private static bool IsBlank(List<string> record) => record.All(string.IsNullOrWhiteSpace);

    /// <summary>Excel-style column name for a 1-based position: A, B, … Z, AA, AB…</summary>
    private static string ColumnLetter(int index)
    {
        var name = "";
        while (index > 0)
        {
            var rem = (index - 1) % 26;
            name = (char)('A' + rem) + name;
            index = (index - 1) / 26;
        }
        return name;
    }

    private static string ReadText(string path) => ReadText(path, int.MaxValue, out _);

    private static string ReadText(string path, int maxChars, out bool complete)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        if (maxChars == int.MaxValue)
        {
            complete = true;
            return reader.ReadToEnd();
        }

        var buffer = new char[Math.Min(maxChars, 64 * 1024)];
        var sb = new StringBuilder();
        int read;
        while (sb.Length < maxChars && (read = reader.Read(buffer, 0, buffer.Length)) > 0)
            sb.Append(buffer, 0, read);

        complete = reader.EndOfStream;
        return sb.ToString();
    }

    private static List<List<string>> ParseCsv(string text, out char delimiterUsed)
    {
        var delimiter = DetectDelimiter(text);
        delimiterUsed = delimiter;
        var records = new List<List<string>>();
        var current = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;

        void EndField() { current.Add(field.ToString()); field.Clear(); }
        void EndRecord() { EndField(); records.Add(current); current = []; }

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else inQuotes = false;
                }
                else field.Append(ch);
            }
            else if (ch == '"') inQuotes = true;
            else if (ch == delimiter) EndField();
            else if (ch == '\r') { /* swallow, \n ends the record */ }
            else if (ch == '\n') EndRecord();
            else field.Append(ch);
        }
        if (field.Length > 0 || current.Count > 0) EndRecord();
        return records;
    }

    /// <summary>
    /// Guesses the separator from the first few non-empty lines — not just the first
    /// one, so a report title above the header row cannot throw the guess off.
    /// </summary>
    private static char DetectDelimiter(string text)
    {
        int commas = 0, semis = 0, tabs = 0, lines = 0;
        for (var start = 0; start < text.Length && lines < 10;)
        {
            var newline = text.IndexOf('\n', start);
            var line = text.AsSpan(start, (newline < 0 ? text.Length : newline) - start);
            if (!line.IsWhiteSpace())
            {
                foreach (var c in line)
                {
                    if (c == ',') commas++;
                    else if (c == ';') semis++;
                    else if (c == '\t') tabs++;
                }
                lines++;
            }
            if (newline < 0) break;
            start = newline + 1;
        }
        if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
        if (semis > commas) return ';';
        return ',';
    }

    private static object? InferValue(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return null;
        var s = raw.Trim();
        if (s.Length == 0) return null;

        if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var dec))
            return dec;
        if (bool.TryParse(s, out var b))
            return b;
        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return dt;
        if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
            return dt;
        return raw;
    }

    // ----------------------------------------------------------------- Excel

    /// <summary>
    /// Reads a worksheet into a table.
    /// <paramref name="headerRow"/> is the 1-based worksheet row holding the column
    /// names (rows above it — titles, logos, blank rows — are ignored); 0 means the
    /// sheet has no header row and the columns are named after their Excel letters.
    /// </summary>
    public static DataTableLite LoadExcel(string path, string sheetName, string name, int headerRow = 1)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = ResolveSheet(workbook, sheetName);

        var used = sheet.RangeUsed();
        if (used is null)
            return new DataTableLite([]) { Name = name };

        var layout = ReadLayout(sheet, used, headerRow);
        var table = new DataTableLite(layout.Headers) { Name = name };

        for (var r = layout.FirstDataRow; r <= layout.LastRow; r++)
        {
            var values = new object?[table.Columns.Count];
            var hasValue = false;
            for (var c = 0; c < table.Columns.Count; c++)
            {
                values[c] = CellValue(sheet.Cell(r, layout.FirstColumn + c));
                hasValue |= values[c] is not null;
            }
            if (hasValue) table.AddRow(values);
        }
        return table;
    }

    /// <summary>
    /// Sheet names and column names of a workbook, without building a table —
    /// what the editor needs to show the columns on the block.
    /// </summary>
    public static SourceSchema PeekExcel(string path, string sheetName, int headerRow = 1)
    {
        using var workbook = new XLWorkbook(path);
        var sheetNames = workbook.Worksheets.Select(w => w.Name).ToList();
        var sheet = ResolveSheet(workbook, sheetName);

        var used = sheet.RangeUsed();
        if (used is null)
            return new SourceSchema
            {
                Label = sheet.Name, Columns = [], HeaderRow = headerRow, RowCount = 0, SheetNames = sheetNames
            };

        var layout = ReadLayout(sheet, used, headerRow);

        var columns = new List<SourceColumnInfo>(layout.Headers.Count);
        for (var c = 0; c < layout.Headers.Count; c++)
        {
            var sample = layout.FirstDataRow <= layout.LastRow
                ? sheet.Cell(layout.FirstDataRow, layout.FirstColumn + c).GetFormattedString()
                : null;
            columns.Add(new SourceColumnInfo(layout.Headers[c], string.IsNullOrWhiteSpace(sample) ? null : sample));
        }

        return new SourceSchema
        {
            Label = sheet.Name,
            Columns = columns,
            HeaderRow = layout.HeaderRow,
            RowCount = Math.Max(0, layout.LastRow - layout.FirstDataRow + 1),
            SheetNames = sheetNames
        };
    }

    private static IXLWorksheet ResolveSheet(XLWorkbook workbook, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            return workbook.Worksheets.First();

        return workbook.Worksheets.FirstOrDefault(
                   w => w.Name.Equals(sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
               ?? throw new InvalidOperationException(
                   $"Sheet '{sheetName.Trim()}' not found. This workbook has: " +
                   $"{string.Join(", ", workbook.Worksheets.Select(w => w.Name))}.");
    }

    private readonly record struct SheetLayout(
        List<string> Headers, int HeaderRow, int FirstDataRow, int LastRow, int FirstColumn);

    /// <summary>Works out where the headers and the data actually are inside the used range.</summary>
    private static SheetLayout ReadLayout(IXLWorksheet sheet, IXLRange used, int headerRow)
    {
        var first = used.RangeAddress.FirstAddress;
        var last = used.RangeAddress.LastAddress;
        var columnNumbers = Enumerable.Range(first.ColumnNumber, last.ColumnNumber - first.ColumnNumber + 1).ToList();

        if (headerRow <= 0)
        {
            // No header row: name the columns after their Excel letters, keep every row as data.
            var letters = columnNumbers.Select(c => XLHelper.GetColumnLetterFromNumber(c)).ToList();
            return new SheetLayout(letters, 0, first.RowNumber, last.RowNumber, first.ColumnNumber);
        }

        if (headerRow > last.RowNumber)
            throw new InvalidOperationException(
                $"Header row {headerRow} is past the last row with data in sheet '{sheet.Name}' " +
                $"(rows {first.RowNumber}–{last.RowNumber}).");

        // Rows above the used range are empty, so "row 1" on a sheet that starts
        // lower simply means its first row that has anything on it.
        var actualHeaderRow = Math.Max(headerRow, first.RowNumber);

        var headers = columnNumbers
            .Select(c => sheet.Cell(actualHeaderRow, c).GetString().Trim() is { Length: > 0 } text
                ? text
                : XLHelper.GetColumnLetterFromNumber(c))
            .ToList();

        return new SheetLayout(headers, actualHeaderRow, actualHeaderRow + 1, last.RowNumber, first.ColumnNumber);
    }

    private static object? CellValue(IXLCell cell)
    {
        var v = cell.Value;
        if (v.IsBlank) return null;
        if (v.IsBoolean) return v.GetBoolean();
        if (v.IsNumber) return (decimal)v.GetNumber();
        if (v.IsDateTime) return v.GetDateTime();
        if (v.IsTimeSpan) return v.GetTimeSpan();
        var text = v.GetText();
        return string.IsNullOrEmpty(text) ? null : text;
    }

    // -------------------------------------------------------------- Database

    public static DataTableLite LoadDatabase(string provider, string connectionString, string query, string name)
    {
        using DbConnection connection = provider.Trim().ToLowerInvariant() switch
        {
            "sqlserver" or "mssql" or "sql" => new SqlConnection(connectionString),
            "sqlite" => new SqliteConnection(connectionString),
            _ => throw new InvalidOperationException(
                $"Unknown database provider '{provider}'. Supported: SqlServer, Sqlite.")
        };

        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = query;
        using var reader = command.ExecuteReader();

        var columns = new List<string>();
        for (var i = 0; i < reader.FieldCount; i++)
            columns.Add(reader.GetName(i));

        var table = new DataTableLite(columns) { Name = name };
        while (reader.Read())
        {
            var row = new object?[reader.FieldCount];
            for (var i = 0; i < reader.FieldCount; i++)
                row[i] = reader.IsDBNull(i) ? null : Normalize(reader.GetValue(i));
            table.AddRow(row);
        }
        return table;
    }

    private static object? Normalize(object value) => value switch
    {
        byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
            => Convert.ToDecimal(value, CultureInfo.InvariantCulture),
        _ => value
    };
}
