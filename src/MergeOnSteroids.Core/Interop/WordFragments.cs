using System.Drawing;
using System.Drawing.Imaging;
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
                png = EmfToPng((byte[])range.EnhMetaFileBits);
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

    // ------------------------------------------------------- EMF → PNG

    /// <summary>
    /// Rasterizes Word's enhanced metafile at the resolution it was recorded at,
    /// then resamples down to display size. Playing a 600-DPI metafile straight
    /// into a small bitmap rounds every glyph position to a whole pixel, which
    /// makes the text look condensed and blotchy; rendering big first avoids that.
    /// </summary>
    private static byte[] EmfToPng(byte[] emfBytes, double displayDpi = 150, int maxDisplayWidth = 1000)
    {
        using var input = new MemoryStream(emfBytes);
        using var metafile = new Metafile(input);

        var (inchesWide, inchesHigh, nativeDpi) = PhysicalSize(metafile);

        var displayWidth = (int)Math.Round(inchesWide * displayDpi);
        if (displayWidth > maxDisplayWidth)
        {
            displayDpi *= maxDisplayWidth / (double)displayWidth;
            displayWidth = maxDisplayWidth;
        }
        displayWidth = Math.Max(1, displayWidth);
        var displayHeight = Math.Max(1, (int)Math.Round(inchesHigh * displayDpi));

        // Render at the metafile's own resolution, bounded so a long fragment
        // cannot allocate an enormous intermediate bitmap.
        var renderDpi = Math.Clamp(nativeDpi, displayDpi, 600);
        const int maxRenderWidth = 4000;
        if (inchesWide * renderDpi > maxRenderWidth)
            renderDpi = maxRenderWidth / inchesWide;
        var renderWidth = Math.Max(displayWidth, (int)Math.Round(inchesWide * renderDpi));
        var renderHeight = Math.Max(displayHeight, (int)Math.Round(inchesHigh * renderDpi));

        using var rendered = new Bitmap(renderWidth, renderHeight);
        using (var g = Graphics.FromImage(rendered))
        {
            g.Clear(Color.White);
            g.DrawImage(metafile, new Rectangle(0, 0, renderWidth, renderHeight));
        }

        using var display = new Bitmap(displayWidth, displayHeight);
        display.SetResolution((float)displayDpi, (float)displayDpi);
        using (var g = Graphics.FromImage(display))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.DrawImage(rendered, new Rectangle(0, 0, displayWidth, displayHeight));
        }

        using var trimmed = TrimTrailingWhitespace(display);
        using var output = new MemoryStream();
        trimmed.Save(output, ImageFormat.Png);
        return output.ToArray();
    }

    /// <summary>
    /// Crops the empty area Word leaves to the right of and below the text, so a
    /// fragment block is only as tall as its content. Never crops the top or left,
    /// which would break the visual alignment of indents.
    /// </summary>
    private static Bitmap TrimTrailingWhitespace(Bitmap source, int padding = 6)
    {
        var data = source.LockBits(
            new Rectangle(0, 0, source.Width, source.Height),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        int lastRow = -1, lastColumn = -1;
        try
        {
            var stride = data.Stride;
            var buffer = new byte[stride * source.Height];
            Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);

            for (var y = 0; y < source.Height; y++)
            {
                var row = y * stride;
                for (var x = 0; x < source.Width; x++)
                {
                    var i = row + x * 4;
                    // anything visibly darker than the white page counts as content
                    if (buffer[i] < 245 || buffer[i + 1] < 245 || buffer[i + 2] < 245)
                    {
                        if (y > lastRow) lastRow = y;
                        if (x > lastColumn) lastColumn = x;
                    }
                }
            }
        }
        finally
        {
            source.UnlockBits(data);
        }

        if (lastRow < 0) return (Bitmap)source.Clone();   // blank fragment

        var width = Math.Min(source.Width, lastColumn + 1 + padding);
        var height = Math.Min(source.Height, lastRow + 1 + padding);
        if (width == source.Width && height == source.Height) return (Bitmap)source.Clone();

        var cropped = new Bitmap(width, height);
        cropped.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        using var g = Graphics.FromImage(cropped);
        g.DrawImage(source, new Rectangle(0, 0, width, height), new Rectangle(0, 0, width, height),
            GraphicsUnit.Pixel);
        return cropped;
    }

    /// <summary>Physical size of the metafile in inches, plus the DPI it was recorded at.</summary>
    private static (double Width, double Height, double NativeDpi) PhysicalSize(Metafile metafile)
    {
        try
        {
            var header = metafile.GetMetafileHeader();
            var dpiX = header.DpiX > 1 ? header.DpiX : 96f;
            var dpiY = header.DpiY > 1 ? header.DpiY : 96f;
            var w = Math.Abs(header.Bounds.Width) / dpiX;
            var h = Math.Abs(header.Bounds.Height) / dpiY;
            if (w > 0.01 && h > 0.01) return (w, h, dpiX);
        }
        catch
        {
            // fall through to the pixel-size estimate
        }
        return (Math.Max(1, metafile.Width) / 96.0, Math.Max(1, metafile.Height) / 96.0, 96.0);
    }
}
