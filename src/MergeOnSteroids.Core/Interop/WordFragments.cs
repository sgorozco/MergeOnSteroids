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
/// One fragment open in Word. The document belongs to the shared
/// <see cref="WordApplication"/>, so a session is cheap: disposing it closes the
/// document and leaves Word running for the next one.
/// </summary>
public sealed class WordFragmentEditSession : IDisposable
{
    private const int WdFormatXMLDocument = 12;
    private const int WdDoNotSaveChanges = 0;

    private readonly WordApplication _word;
    private readonly bool _visible;
    private dynamic? _doc;

    private WordFragmentEditSession(WordApplication word, bool visible)
    {
        _word = word;
        _visible = visible;
    }

    /// <summary>New fragment: a blank document, optionally seeded with Flat OPC content.</summary>
    public static WordFragmentEditSession Start(WordApplication word, string? seedFragmentXml, bool visible = true)
    {
        var session = new WordFragmentEditSession(word, visible);
        session._doc = word.Instance.Documents.Add(Type.Missing, Type.Missing, Type.Missing, visible);
        if (!string.IsNullOrWhiteSpace(seedFragmentXml))
        {
            dynamic range = session._doc!.Content;
            range.InsertXML(seedFragmentXml);
            Release(range);
        }
        session.Reveal();
        return session;
    }

    /// <summary>Existing fragment: opens its .docx for editing in place.</summary>
    public static WordFragmentEditSession OpenFile(WordApplication word, string docxPath, bool visible = true)
    {
        var session = new WordFragmentEditSession(word, visible);
        // positional: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, … Visible (12th)
        session._doc = word.Instance.Documents.Open(docxPath, false, false, false,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, visible);
        session.Reveal();
        return session;
    }

    private void Reveal()
    {
        if (_visible) _word.Show();
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

    /// <summary>
    /// Closes the fragment and puts Word away — but leaves it running. Word itself is
    /// only quit when the editor exits, because quitting and restarting it is what made
    /// every edit cost half a minute.
    /// </summary>
    public void Dispose()
    {
        if (_doc is null) return;

        var doc = _doc;
        _doc = null;
        try { doc.Close(WdDoNotSaveChanges); } catch (Exception) { /* user may have closed it already */ }
        Release(doc);        // while Word is still alive, so this is instant

        if (_visible) _word.Hide();
    }

    private static void Release(object? o)
    {
        if (o is not null && Marshal.IsComObject(o))
            Marshal.ReleaseComObject(o);
    }
}
