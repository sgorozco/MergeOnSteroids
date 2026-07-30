using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace MailOnSteroids.Core.Interop;

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
/// Opens a fragment in a visible Word window for direct editing, then captures
/// it back: formatted XML (Range.WordOpenXML), a PNG preview rendered by Word
/// itself (Range.EnhMetaFileBits), and the plain text.
/// </summary>
public sealed class WordFragmentEditSession : IDisposable
{
    private dynamic? _app;
    private dynamic? _doc;

    private WordFragmentEditSession() { }

    /// <summary>Starts Word (visible unless told otherwise) with the fragment loaded.</summary>
    public static WordFragmentEditSession Start(string? fragmentXml, bool visible = true)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed (Word.Application COM class not found).");

        var session = new WordFragmentEditSession();
        session._app = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Could not start Microsoft Word.");
        session._app!.DisplayAlerts = 0;
        session._app!.Visible = visible;

        session._doc = session._app!.Documents.Add();
        if (!string.IsNullOrWhiteSpace(fragmentXml))
        {
            dynamic range = session._doc!.Content;
            range.InsertXML(fragmentXml);
        }
        if (visible)
        {
            try { session._app!.Activate(); } catch { /* focus is best effort */ }
        }
        return session;
    }

    /// <summary>Direct access for programmatic authoring (sample generation).</summary>
    public dynamic Document => _doc ?? throw new InvalidOperationException("Session is closed.");

    /// <summary>Captures the current state of the fragment document.</summary>
    public FragmentCapture Capture()
    {
        var doc = _doc ?? throw new InvalidOperationException("Session is closed.");

        // Exclude the structural final paragraph mark so fragments don't carry
        // a trailing empty paragraph into every document.
        int end = Math.Max(0, (int)doc.Content.End - 1);
        if (end == 0)
            return new FragmentCapture("", [], "");

        dynamic range = doc.Range(0, end);
        string xml = range.WordOpenXML;
        string text = ((string)range.Text ?? "").Replace("\r", "\n").TrimEnd('\n');
        byte[] png = [];
        try
        {
            byte[] emf = (byte[])range.EnhMetaFileBits;
            png = EmfToPng(emf);
        }
        catch
        {
            // preview is optional — content capture must not fail because of it
        }
        return new FragmentCapture(xml, png, text);
    }

    public void Dispose()
    {
        if (_doc is not null)
        {
            try { _doc.Close(0); } catch { /* user may have closed it already */ }
            Release(_doc);
            _doc = null;
        }
        if (_app is not null)
        {
            try { _app.Quit(0); } catch { /* best effort */ }
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

    // ------------------------------------------------------- EMF → PNG

    private static byte[] EmfToPng(byte[] emfBytes)
    {
        using var input = new MemoryStream(emfBytes);
        using var metafile = new System.Drawing.Imaging.Metafile(input);

        const int maxWidth = 640;
        var scale = Math.Min(1.0, (double)maxWidth / metafile.Width);
        var w = Math.Max(1, (int)(metafile.Width * scale));
        var h = Math.Max(1, (int)(metafile.Height * scale));

        using var bitmap = new System.Drawing.Bitmap(w, h);
        using (var g = System.Drawing.Graphics.FromImage(bitmap))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.Clear(System.Drawing.Color.White);
            g.DrawImage(metafile, 0, 0, w, h);
        }
        using var output = new MemoryStream();
        bitmap.Save(output, System.Drawing.Imaging.ImageFormat.Png);
        return output.ToArray();
    }
}
