namespace MergeOnSteroids.Core.Interop;

/// <summary>
/// Word's built-in style constants (WdBuiltinStyle). Shared by the writer and the
/// preview renderer so what a block shows cannot drift from what a run produces.
/// </summary>
public static class WordStyleIds
{
    public const int Normal = -1;
    public const int Heading1 = -2;
    public const int Heading2 = -3;
    public const int Heading3 = -4;
    public const int Title = -63;
    public const int Subtitle = -75;
    public const int Quote = -181;
    public const int ListBullet = -49;   // -12 is "Index 2": the old value produced no bullet

    /// <summary>The constant for one of <see cref="Blocks.ParagraphStyles"/>.</summary>
    public static int Of(string style) => style switch
    {
        "Heading 1" => Heading1,
        "Heading 2" => Heading2,
        "Heading 3" => Heading3,
        "Title" => Title,
        "Subtitle" => Subtitle,
        "Quote" => Quote,
        "List Bullet" => ListBullet,
        _ => Normal
    };
}
