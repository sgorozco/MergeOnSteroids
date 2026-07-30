using System.Data.Common;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;

namespace MailOnSteroids.Core.Data;

/// <summary>Loads external data (CSV, Excel, databases) into <see cref="DataTableLite"/>.</summary>
public static class DataSourceLoaders
{
    // ------------------------------------------------------------------ CSV

    public static DataTableLite LoadCsv(string path, string name)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var text = reader.ReadToEnd();
        var records = ParseCsv(text);
        if (records.Count == 0)
            return new DataTableLite([]) { Name = name };

        var table = new DataTableLite(records[0]) { Name = name };
        for (var i = 1; i < records.Count; i++)
        {
            var rec = records[i];
            if (rec.Count == 1 && string.IsNullOrWhiteSpace(rec[0])) continue; // trailing blank line
            var row = new object?[table.Columns.Count];
            for (var c = 0; c < table.Columns.Count && c < rec.Count; c++)
                row[c] = InferValue(rec[c]);
            table.AddRow(row);
        }
        return table;
    }

    private static List<List<string>> ParseCsv(string text)
    {
        var delimiter = DetectDelimiter(text);
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

    private static char DetectDelimiter(string text)
    {
        var firstLine = text.AsSpan(0, text.IndexOf('\n') is var n and >= 0 ? n : text.Length);
        int commas = 0, semis = 0, tabs = 0;
        foreach (var c in firstLine)
        {
            if (c == ',') commas++;
            else if (c == ';') semis++;
            else if (c == '\t') tabs++;
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

    public static DataTableLite LoadExcel(string path, string sheetName, string name)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = string.IsNullOrWhiteSpace(sheetName)
            ? workbook.Worksheets.First()
            : workbook.Worksheet(sheetName);

        var used = sheet.RangeUsed();
        if (used is null)
            return new DataTableLite([]) { Name = name };

        var headerRow = used.FirstRow();
        var headers = headerRow.Cells().Select(c => c.GetString()).ToList();
        var table = new DataTableLite(headers) { Name = name };

        foreach (var row in used.Rows().Skip(1))
        {
            var values = new object?[table.Columns.Count];
            var i = 0;
            foreach (var cell in row.Cells(1, table.Columns.Count))
            {
                values[i++] = CellValue(cell);
                if (i >= table.Columns.Count) break;
            }
            if (values.All(v => v is null)) continue;
            table.AddRow(values);
        }
        return table;
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
