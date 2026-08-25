using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MergeOnSteroids.Core.Fragments;

/// <summary>
/// A row of a table inside a fragment that carries a {repeat source} marker: at run
/// time it is cloned once per record of that source, keeping every bit of the
/// formatting the row was drawn with.
/// </summary>
/// <param name="Marker">The marker exactly as it reads in the document, so it can be found and removed.</param>
/// <param name="SourceName">The data source named in the marker.</param>
/// <param name="CellTemplates">The row's cells, for writers that cannot clone a Word row.</param>
/// <param name="RowText">Everything the row says, used to tell its placeholders from the rest.</param>
public sealed record FragmentRepeatRow(
    string Marker, string SourceName, IReadOnlyList<string> CellTemplates, string RowText);

/// <summary>Finds the repeating rows in a fragment .docx.</summary>
public static class FragmentRepeats
{
    private static readonly Regex MarkerPattern =
        new(@"\{\s*repeat\s+([^}]+?)\s*\}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static IReadOnlyList<FragmentRepeatRow> Read(string docxPath)
    {
        try
        {
            using var document = WordprocessingDocument.Open(docxPath, isEditable: false);
            var body = document.MainDocumentPart?.Document?.Body;
            if (body is null) return [];

            var repeats = new List<FragmentRepeatRow>();
            foreach (var row in body.Descendants<TableRow>())
            {
                var cells = row.Elements<TableCell>().Select(c => c.InnerText).ToList();
                var rowText = string.Join("\n", cells);

                var match = MarkerPattern.Match(rowText);
                if (!match.Success) continue;

                repeats.Add(new FragmentRepeatRow(
                    Marker: match.Value,
                    SourceName: match.Groups[1].Value.Trim(),
                    CellTemplates: cells,
                    RowText: rowText));
            }
            return repeats;
        }
        catch (Exception)
        {
            return [];   // an unreadable fragment is reported elsewhere; it simply has no repeats
        }
    }
}
