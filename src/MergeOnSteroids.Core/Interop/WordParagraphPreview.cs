using System.Runtime.InteropServices;

namespace MergeOnSteroids.Core.Interop;

/// <summary>Everything that decides how a paragraph looks on the page.</summary>
public sealed record ParagraphPreviewRequest(
    string? TemplatePath, string Text, string Style, bool Bold, bool Italic)
{
    /// <summary>Two requests with the same key must render identically.</summary>
    public string Key => $"{TemplatePath}\u0001{Style}\u0001{(Bold ? 'B' : '-')}{(Italic ? 'I' : '-')}\u0001{Text}";
}

/// <summary>
/// Renders single paragraphs the way Word will, by laying each one out in a hidden
/// document and asking Word for its own picture of it. The document is built from
/// the same template the run would use, so a block shows that template's real
/// styles — font, size, colour, spacing, bullets — rather than an approximation.
///
/// One hidden Word instance serves every preview and stays open between them, since
/// starting Word costs far more than laying out a paragraph. Not thread-safe: drive
/// it from a single STA thread.
/// </summary>
public sealed class WordParagraphPreview : IDisposable
{
    private const int WdDoNotSaveChanges = 0;

    private readonly Dictionary<string, dynamic> _scratchDocuments = new(StringComparer.OrdinalIgnoreCase);
    private dynamic? _app;

    /// <summary>Word's rendering of one paragraph as PNG bytes; empty when there is nothing to draw.</summary>
    public byte[] Render(ParagraphPreviewRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Text)) return [];

        var document = ScratchDocument(request.TemplatePath);

        dynamic content = document.Content;
        try
        {
            content.Delete();
        }
        finally
        {
            Release(content);
        }

        dynamic range = document.Content;
        try
        {
            // Exactly what WordComWriter.AddParagraph does, so the preview and the
            // generated document cannot disagree about what this block produces.
            range.Text = request.Text;
            try
            {
                range.Style = WordStyleIds.Of(request.Style);
            }
            catch (COMException)
            {
                // style missing from this template — Word leaves it as body text, and so do we
            }
            if (request.Bold) range.Font.Bold = 1;
            if (request.Italic) range.Font.Italic = 1;
        }
        finally
        {
            Release(range);
        }

        return Capture(document);
    }

    /// <summary>Word's picture of the document, minus the trailing paragraph mark.</summary>
    private static byte[] Capture(dynamic document)
    {
        var end = Math.Max(0, (int)document.Content.End - 1);
        if (end == 0) return [];

        dynamic range = document.Range(0, end);
        try
        {
            return WordRendering.EmfToPng((byte[])range.EnhMetaFileBits);
        }
        catch (Exception)
        {
            return [];   // a preview is never worth breaking the editor over
        }
        finally
        {
            Release(range);
        }
    }

    /// <summary>
    /// The hidden document paragraphs are laid out in — one per template, kept open
    /// so a template's styles are loaded once rather than per keystroke.
    /// </summary>
    private dynamic ScratchDocument(string? templatePath)
    {
        var key = templatePath ?? "";
        if (_scratchDocuments.TryGetValue(key, out var existing)) return existing;

        var app = RequireApp();
        dynamic document = string.IsNullOrWhiteSpace(templatePath)
            ? app.Documents.Add()
            : app.Documents.Add(templatePath);

        _scratchDocuments[key] = document;
        return document;
    }

    private dynamic RequireApp()
    {
        if (_app is not null) return _app;

        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed (Word.Application COM class not found).");

        _app = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Could not start Microsoft Word.");
        _app!.Visible = false;
        _app!.DisplayAlerts = 0;
        return _app;
    }

    public void Dispose()
    {
        foreach (var document in _scratchDocuments.Values)
        {
            try { document.Close(WdDoNotSaveChanges); } catch (Exception) { /* best effort */ }
            Release(document);
        }
        _scratchDocuments.Clear();

        if (_app is null) return;
        try { _app.Quit(WdDoNotSaveChanges); } catch (Exception) { /* best effort */ }
        Release(_app);
        _app = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static void Release(object? o)
    {
        if (o is not null && Marshal.IsComObject(o))
            Marshal.ReleaseComObject(o);
    }
}
