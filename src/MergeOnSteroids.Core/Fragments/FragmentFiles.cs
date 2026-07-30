using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MergeOnSteroids.Core.Fragments;

/// <summary>
/// Conventions and file access for fragment sidecars: each fragment is a small
/// .docx, accompanied by a ".preview.png" that holds Word's own rendering of it.
/// </summary>
public static class FragmentFiles
{
    /// <summary>Folder (relative to the program file) where new fragments are created.</summary>
    public const string FolderName = "fragments";

    public static string DefaultRelativePath(string? blockId)
    {
        var suffix = blockId is { Length: >= 8 } ? blockId[..8] : Guid.NewGuid().ToString("N")[..8];
        return $"{FolderName}/fragment-{suffix}.docx";
    }

    /// <summary>The preview image that belongs to a fragment document.</summary>
    public static string PreviewPathFor(string docxPath) =>
        Path.ChangeExtension(docxPath, null) + ".preview.png";

    /// <summary>
    /// Paragraph text of a fragment .docx, without needing Word. Used for
    /// placeholder scanning and for the text-only preview run.
    /// </summary>
    public static string ReadPlainText(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body is null) return "";
        return string.Join("\n", body.Descendants<Paragraph>().Select(p => p.InnerText));
    }

    public static byte[]? TryReadPreview(string docxPath)
    {
        try
        {
            var path = PreviewPathFor(docxPath);
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }
        catch (IOException)
        {
            return null;
        }
    }

    public static void WritePreview(string docxPath, byte[] png)
    {
        if (png.Length == 0) return;
        File.WriteAllBytes(PreviewPathFor(docxPath), png);
    }

    /// <summary>Program-relative path with forward slashes, so programs stay portable.</summary>
    public static string MakeRelative(string baseFolder, string fullPath)
    {
        var relative = Path.GetRelativePath(baseFolder, fullPath);
        return Path.IsPathRooted(relative) ? relative : relative.Replace('\\', '/');
    }
}
