using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace MergeOnSteroids.Core.Interop;

/// <summary>Helpers for Flat OPC WordprocessingML fragments.</summary>
public static class WordFragmentText
{
    private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Plain text of a Flat OPC fragment: paragraphs joined with newlines.</summary>
    public static string Extract(string flatOpcXml)
    {
        if (string.IsNullOrWhiteSpace(flatOpcXml)) return "";
        try
        {
            var doc = XDocument.Parse(flatOpcXml);
            var documentPart = doc.Descendants(Pkg + "part")
                .FirstOrDefault(p => (string?)p.Attribute(Pkg + "name") == "/word/document.xml");
            var body = documentPart?.Descendants(W + "body").FirstOrDefault();
            if (body is null) return "";

            var paragraphs = body.Descendants(W + "p")
                .Select(p => string.Concat(p.Descendants(W + "t").Select(t => (string)t)));
            return string.Join("\n", paragraphs);
        }
        catch
        {
            return "";
        }
    }
}

public sealed record FragmentCapture(string Xml, byte[] PreviewPng, string PlainText);

/// <summary>
/// A live Word window holding one fragment. Opens the fragment's own .docx (or a
/// blank document for a new fragment) so the user edits the real file, then saves
/// it back and captures a preview rendered by Word itself.
/// </summary>
public sealed class WordFragmentEditSession : IDisposable
{
    private const int WdFormatXMLDocument = 12;
    private const int WdDoNotSaveChanges = 0;

    private dynamic? _app;
    private dynamic? _doc;

    private WordFragmentEditSession() { }

    /// <summary>New fragment: a blank document, optionally seeded with Flat OPC content.</summary>
    public static WordFragmentEditSession Start(string? seedFragmentXml, bool visible = true)
    {
        var session = new WordFragmentEditSession();
        session.StartWord(visible);
        session._doc = session._app!.Documents.Add();
        if (!string.IsNullOrWhiteSpace(seedFragmentXml))
        {
            dynamic range = session._doc!.Content;
            range.InsertXML(seedFragmentXml);
            Release(range);
        }
        session.Activate(visible);
        return session;
    }

    /// <summary>Existing fragment: opens its .docx for editing in place.</summary>
    public static WordFragmentEditSession OpenFile(string docxPath, bool visible = true)
    {
        var session = new WordFragmentEditSession();
        session.StartWord(visible);
        session._doc = session._app!.Documents.Open(docxPath, false, false, false);
        session.Activate(visible);
        return session;
    }

    private void StartWord(bool visible)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed (Word.Application COM class not found).");

        _app = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Could not start Microsoft Word.");
        _app!.DisplayAlerts = 0;
        _app!.Visible = visible;
    }

    private void Activate(bool visible)
    {
        if (!visible) return;
        try { _app!.Activate(); } catch { /* focus is best effort */ }
    }

    /// <summary>Direct access for programmatic authoring (sample generation).</summary>
    public dynamic Document => _doc ?? throw new InvalidOperationException("Session is closed.");

    /// <summary>Saves the fragment to <paramref name="docxPath"/> and captures it.</summary>
    public FragmentCapture SaveAs(string docxPath)
    {
        var doc = _doc ?? throw new InvalidOperationException("Session is closed.");
        Directory.CreateDirectory(Path.GetDirectoryName(docxPath)!);
        doc.SaveAs2(docxPath, WdFormatXMLDocument);
        return Capture();
    }

    /// <summary>
    /// Captures the current fragment: Flat OPC XML, Word's rendering as PNG, and
    /// the plain text. The document's final paragraph mark is excluded so a
    /// fragment does not carry a trailing empty paragraph into every document.
    /// </summary>
    public FragmentCapture Capture()
    {
        var doc = _doc ?? throw new InvalidOperationException("Session is closed.");

        int end = Math.Max(0, (int)doc.Content.End - 1);
        if (end == 0)
            return new FragmentCapture("", [], "");

        dynamic range = doc.Range(0, end);
        try
        {
            string xml = range.WordOpenXML;
            string text = ((string)range.Text ?? "").Replace("\r", "\n").TrimEnd('\n');
            byte[] png = [];
            try
            {
                png = WordRendering.EmfToPng((byte[])range.EnhMetaFileBits);
            }
            catch
            {
                // preview is optional — content capture must not fail because of it
            }
            return new FragmentCapture(xml, png, text);
        }
        finally
        {
            Release(range);
        }
    }

    public void Dispose()
    {
        if (_doc is not null)
        {
            try { _doc.Close(WdDoNotSaveChanges); } catch { /* user may have closed it already */ }
            Release(_doc);
            _doc = null;
        }
        if (_app is not null)
        {
            try { _app.Quit(WdDoNotSaveChanges); } catch { /* best effort */ }
            Release(_app);
            _app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void Release(object? o)
    {
        if (o is not null && Marshal.IsComObject(o))
            Marshal.ReleaseComObject(o);
    }
}
