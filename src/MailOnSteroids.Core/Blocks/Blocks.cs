using System.Text.Json.Serialization;

namespace MailOnSteroids.Core.Blocks;

// ---------------------------------------------------------------------------
// Data source blocks
// ---------------------------------------------------------------------------

/// <summary>Base for blocks that register a named data source.</summary>
public abstract class SourceBlockBase : Block
{
    private string _name = "";
    /// <summary>Name other blocks use to refer to this source.</summary>
    public string Name { get => _name; set => Set(ref _name, value); }

    public override BlockCategory Category => BlockCategory.Data;
}

public sealed class CsvSourceBlock : SourceBlockBase
{
    private string _filePath = "";
    public string FilePath { get => _filePath; set => Set(ref _filePath, value); }

    public override string DisplayName => "open CSV data source";
}

public sealed class ExcelSourceBlock : SourceBlockBase
{
    private string _filePath = "";
    private string _sheetName = "";

    public string FilePath { get => _filePath; set => Set(ref _filePath, value); }
    /// <summary>Worksheet name; empty = first sheet.</summary>
    public string SheetName { get => _sheetName; set => Set(ref _sheetName, value); }

    public override string DisplayName => "open Excel data source";
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

    public override string DisplayName => "open database data source";
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

    public override string DisplayName => "filter data source";
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

    public override string DisplayName => "for each record";
    public override BlockCategory Category => BlockCategory.Control;
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

    public override string DisplayName => "if";
    public override BlockCategory Category => BlockCategory.Control;
    public override IEnumerable<BlockCollection> ChildLists() => [Children, Else];
}

public sealed class SetVariableBlock : Block
{
    private string _variableName = "x";
    private string _valueExpression = "";

    public string VariableName { get => _variableName; set => Set(ref _variableName, value); }
    public string ValueExpression { get => _valueExpression; set => Set(ref _valueExpression, value); }

    public override string DisplayName => "set variable";
    public override BlockCategory Category => BlockCategory.Variables;
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

    public override string DisplayName => "new document";
    public override BlockCategory Category => BlockCategory.Document;
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

    public override string DisplayName => "add paragraph";
    public override BlockCategory Category => BlockCategory.Document;
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

    public override string DisplayName => "add table";
    public override BlockCategory Category => BlockCategory.Document;
}

public sealed class PageBreakBlock : Block
{
    public override string DisplayName => "page break";
    public override BlockCategory Category => BlockCategory.Document;
}
