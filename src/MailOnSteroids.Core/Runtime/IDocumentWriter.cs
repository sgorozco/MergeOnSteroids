namespace MailOnSteroids.Core.Runtime;

/// <summary>Options for a single run of a program.</summary>
public sealed class RunOptions
{
    /// <summary>Folder used to resolve relative data-source / template paths (usually the program file's folder).</summary>
    public string BaseFolder { get; set; } = Environment.CurrentDirectory;

    /// <summary>Absolute output folder for generated documents.</summary>
    public string OutputFolder { get; set; } = Environment.CurrentDirectory;

    /// <summary>Show the Word window while generating (Word writer only).</summary>
    public bool ShowWord { get; set; }

    public CancellationToken Cancellation { get; set; } = CancellationToken.None;

    public string ResolvePath(string path) =>
        Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(BaseFolder, path));
}

/// <summary>
/// Target that receives document content produced by the interpreter.
/// Implementations: Word COM automation, plain-text preview.
/// </summary>
public interface IDocumentWriter : IDisposable
{
    bool InDocument { get; }

    void Begin(RunOptions options);
    void BeginDocument(string? templatePath);
    void AddParagraph(string text, string style, bool bold, bool italic);
    /// <summary>
    /// Inserts a formatted Word fragment (Flat OPC XML), then applies the given
    /// literal find→replace substitutions inside the inserted content only.
    /// <paramref name="plainText"/> is the fragment's text for non-Word writers.
    /// </summary>
    void AddFragment(string fragmentXml, string plainText,
        IReadOnlyList<KeyValuePair<string, string>> replacements);
    void AddTable(IReadOnlyList<string> headers, IReadOnlyList<string[]> rows, bool headerRow);
    void PageBreak();
    /// <summary>Finish the current document. Returns the path it was saved to.</summary>
    string EndDocument(string fileNameWithoutExtension);
    void End();
}
