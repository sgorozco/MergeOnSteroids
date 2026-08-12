using System.Text.Json.Serialization;
using MergeOnSteroids.Core.Data;

namespace MergeOnSteroids.Core.Blocks;

// ---------------------------------------------------------------------------
// Data source blocks
// ---------------------------------------------------------------------------

/// <summary>Base for blocks that register a named data source.</summary>
public abstract class SourceBlockBase : Block
{
    private string _name = "";
    /// <summary>Name other blocks use to refer to this source.</summary>
    public string Name { get => _name; set => Set(ref _name, value); }
}

/// <summary>
/// Base for the data sources that read a file the user can point at: they share a
/// path, a header row, and the columns the editor reads back from the file so the
/// block can show them.
/// </summary>
public abstract class FileSourceBlockBase : SourceBlockBase
{
    private string _filePath = "";
    private int _headerRow = 1;
    private IReadOnlyList<SourceColumnInfo> _detectedColumns = [];
    private string _schemaStatus = "";

    /// <summary>File path, relative to the program file or absolute. Supports {expressions}.</summary>
    public string FilePath { get => _filePath; set => Set(ref _filePath, value); }

    /// <summary>
    /// 1-based row holding the column names — anything above it (report titles,
    /// blank rows) is skipped. 0 means the file has no header row, and the columns
    /// are then named after their spreadsheet letters (A, B, C…).
    /// </summary>
    public int HeaderRow { get => _headerRow; set => Set(ref _headerRow, Math.Max(0, value)); }

    /// <summary>
    /// Columns read from the file so the editor can show them on the block.
    /// Design-time only — re-read from the file, never saved with the program.
    /// </summary>
    [JsonIgnore]
    public IReadOnlyList<SourceColumnInfo> DetectedColumns
    {
        get => _detectedColumns;
        set { Set(ref _detectedColumns, value); Raise(nameof(HasDetectedColumns)); }
    }

    /// <summary>What the last read of the file found (or why it failed). Design-time only.</summary>
    [JsonIgnore]
    public string SchemaStatus { get => _schemaStatus; set => Set(ref _schemaStatus, value); }

    [JsonIgnore] public bool HasDetectedColumns => DetectedColumns.Count > 0;
}

public sealed class CsvSourceBlock : FileSourceBlockBase
{
    [JsonIgnore] public override string DisplayName => "open CSV data source";
}

public sealed class ExcelSourceBlock : FileSourceBlockBase
{
    private string _sheetName = "";
    private IReadOnlyList<string> _detectedSheets = [];

    /// <summary>Worksheet name; empty = first sheet.</summary>
    public string SheetName { get => _sheetName; set => Set(ref _sheetName, value); }

    /// <summary>Sheet names of the workbook, offered as a drop-down. Design-time only.</summary>
    [JsonIgnore]
    public IReadOnlyList<string> DetectedSheets
    {
        get => _detectedSheets;
        set => Set(ref _detectedSheets, value);
    }

    [JsonIgnore] public override string DisplayName => "open Excel data source";
}

public sealed class DatabaseSourceBlock : SourceBlockBase
{
    private string _provider = "SqlServer";
    private string _connectionString = "";
    private string _query = "";

    /// <summary>"SqlServer" or "Sqlite".</summary>
    public string Provider { get => _provider; set => Set(ref _provider, value); }
    public string ConnectionString { get => _connectionString; set => Set(ref _connectionString, value); }
    /// <summary>SQL query. Supports {expression} interpolation against the current record scope.</summary>
    public string Query { get => _query; set => Set(ref _query, value); }

    [JsonIgnore] public override string DisplayName => "open database data source";
}

/// <summary>
/// Creates a new in-memory source by filtering another source with an expression.
/// Inside the condition, the candidate row's fields are available bare ("Amount"),
/// and outer loop records via their alias ("c.Id").
/// </summary>
public sealed class FilterSourceBlock : SourceBlockBase
{
    private string _sourceName = "";
    private string _condition = "";

    public string SourceName { get => _sourceName; set => Set(ref _sourceName, value); }
    public string Condition { get => _condition; set => Set(ref _condition, value); }

    [JsonIgnore] public override string DisplayName => "filter data source";
}

// ---------------------------------------------------------------------------
// Control blocks
// ---------------------------------------------------------------------------

public sealed class ForEachBlock : Block
{
    private string _alias = "rec";
    private string _sourceName = "";

    /// <summary>Alias for the current record, e.g. "c" → "c.Name".</summary>
    public string Alias { get => _alias; set => Set(ref _alias, value); }
    public string SourceName { get => _sourceName; set => Set(ref _sourceName, value); }

    [JsonObjectCreationHandling(JsonObjectCreationHandling.Populate)]
    public BlockCollection Children { get; }

    public ForEachBlock() => Children = new BlockCollection(this, nameof(Children));

    [JsonIgnore] public override string DisplayName => "for each record";
    public override IEnumerable<BlockCollection> ChildLists() => [Children];
}

public sealed class IfBlock : Block
{
    private string _condition = "";
    public string Condition { get => _condition; set => Set(ref _condition, value); }

    [JsonObjectCreationHandling(JsonObjectCreationHandling.Populate)]
    public BlockCollection Children { get; }

    [JsonObjectCreationHandling(JsonObjectCreationHandling.Populate)]
    public BlockCollection Else { get; }

    public IfBlock()
    {
        Children = new BlockCollection(this, nameof(Children));
        Else = new BlockCollection(this, nameof(Else));
    }

    [JsonIgnore] public override string DisplayName => "if";
    public override IEnumerable<BlockCollection> ChildLists() => [Children, Else];
}

public sealed class SetVariableBlock : Block
{
    private string _variableName = "x";
    private string _valueExpression = "";

    public string VariableName { get => _variableName; set => Set(ref _variableName, value); }
    public string ValueExpression { get => _valueExpression; set => Set(ref _valueExpression, value); }

    [JsonIgnore] public override string DisplayName => "set variable";
}

// ---------------------------------------------------------------------------
// Document blocks
// ---------------------------------------------------------------------------

/// <summary>
/// Starts a new Word document; children add content; the document is saved when
/// the block finishes. Typically placed inside a "for each record" loop to get
/// one document per record.
/// </summary>
public sealed class NewDocumentBlock : Block
{
    private string _fileNameTemplate = "Document_{1}";
    private string _templatePath = "";

    /// <summary>File name (no extension); supports {expression} interpolation.</summary>
    public string FileNameTemplate { get => _fileNameTemplate; set => Set(ref _fileNameTemplate, value); }
    /// <summary>Optional .dotx/.docx template the document is based on.</summary>
    public string TemplatePath { get => _templatePath; set => Set(ref _templatePath, value); }

    [JsonObjectCreationHandling(JsonObjectCreationHandling.Populate)]
    public BlockCollection Children { get; }

    public NewDocumentBlock() => Children = new BlockCollection(this, nameof(Children));

    [JsonIgnore] public override string DisplayName => "new document";
    public override IEnumerable<BlockCollection> ChildLists() => [Children];
}

public static class ParagraphStyles
{
    public const string Normal = "Normal";
    public const string Heading1 = "Heading 1";
    public const string Heading2 = "Heading 2";
    public const string Heading3 = "Heading 3";
    public const string Title = "Title";
    public const string Subtitle = "Subtitle";
    public const string Quote = "Quote";
    public const string ListBullet = "List Bullet";

    public static readonly string[] All =
        [Normal, Heading1, Heading2, Heading3, Title, Subtitle, Quote, ListBullet];
}

public sealed class ParagraphBlock : Block
{
    private string _textTemplate = "";
    private string _style = ParagraphStyles.Normal;
    private bool _bold;
    private bool _italic;

    /// <summary>Paragraph text; supports {expression} interpolation.</summary>
    public string TextTemplate { get => _textTemplate; set => Set(ref _textTemplate, value); }
    public string Style { get => _style; set => Set(ref _style, value); }
    public bool Bold { get => _bold; set => Set(ref _bold, value); }
    public bool Italic { get => _italic; set => Set(ref _italic, value); }

    [JsonIgnore] public override string DisplayName => "add paragraph";
}

/// <summary>
/// A fragment of real Word content, authored and formatted directly in Word.
/// The content lives in its own small .docx next to the program (a reusable
/// "paragraph library" file you can also open straight from Explorer); the
/// program only stores the relative path. At run time the fragment is inserted
/// into the output document and {expressions} in its text are substituted,
/// keeping every bit of Word formatting.
/// </summary>
public sealed class WordFragmentBlock : Block
{
    private string _fragmentFile = "";
    private string? _legacyFragmentXml;
    private string _plainText = "";
    private byte[]? _previewImage;

    /// <summary>
    /// Path of the fragment .docx, relative to the program file (or absolute).
    /// Several blocks may point at the same file to reuse a fragment.
    /// </summary>
    public string FragmentFile
    {
        get => _fragmentFile;
        set { if (Set(ref _fragmentFile, value)) Raise(nameof(HasContent)); }
    }

    /// <summary>
    /// Inline Flat OPC content written by the first (pre-sidecar) version of the
    /// fragment block. Still honoured so old programs keep working; it is
    /// migrated to a file the next time the fragment is edited.
    /// </summary>
    [JsonPropertyName("fragmentXml")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LegacyFragmentXml
    {
        get => _legacyFragmentXml;
        set { if (Set(ref _legacyFragmentXml, string.IsNullOrWhiteSpace(value) ? null : value)) Raise(nameof(HasContent)); }
    }

    /// <summary>Fragment text, read from the .docx when the program loads. Not persisted.</summary>
    [JsonIgnore]
    public string PlainText { get => _plainText; set => Set(ref _plainText, value); }

    /// <summary>
    /// Word's own rendering of the fragment, shown inside the block. Loaded from
    /// the ".preview.png" file that sits beside the .docx, so it is not persisted here.
    /// </summary>
    [JsonIgnore]
    public byte[]? PreviewImage { get => _previewImage; set => Set(ref _previewImage, value); }

    [JsonIgnore]
    public bool HasContent =>
        !string.IsNullOrWhiteSpace(FragmentFile) || !string.IsNullOrWhiteSpace(LegacyFragmentXml);

    [JsonIgnore] public override string DisplayName => "Word paragraphs";
}

/// <summary>
/// Inserts a table filled from a data source. Columns spec syntax:
///   Header: expression | Header2: expression2 | ...
/// Empty spec = one column per source field, raw values.
/// </summary>
public sealed class TableBlock : Block
{
    private string _sourceName = "";
    private string _columnsSpec = "";
    private bool _headerRow = true;

    public string SourceName { get => _sourceName; set => Set(ref _sourceName, value); }
    public string ColumnsSpec { get => _columnsSpec; set => Set(ref _columnsSpec, value); }
    public bool HeaderRow { get => _headerRow; set => Set(ref _headerRow, value); }

    [JsonIgnore] public override string DisplayName => "add table";
}

public sealed class PageBreakBlock : Block
{
    [JsonIgnore] public override string DisplayName => "page break";
}
