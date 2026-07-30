using System.Globalization;
using System.Windows;
using System.Windows.Data;
using MailOnSteroids.Core.Blocks;

namespace MailOnSteroids.App.Common;

/// <summary>Font size preview for a Word paragraph style.</summary>
public sealed class StyleFontSizeConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        (value as string) switch
        {
            ParagraphStyles.Title => 22.0,
            ParagraphStyles.Subtitle => 14.0,
            ParagraphStyles.Heading1 => 18.0,
            ParagraphStyles.Heading2 => 15.5,
            ParagraphStyles.Heading3 => 13.5,
            _ => 12.5
        };

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>FontWeight preview: heading styles or the Bold flag.</summary>
public sealed class StyleWeightConverter : IMultiValueConverter
{
    public object Convert(object?[] values, Type targetType, object? parameter, CultureInfo culture)
    {
        var style = values.ElementAtOrDefault(0) as string ?? "";
        var bold = values.ElementAtOrDefault(1) as bool? ?? false;
        var styleBold = style is ParagraphStyles.Title or ParagraphStyles.Heading1
            or ParagraphStyles.Heading2 or ParagraphStyles.Heading3;
        return bold || styleBold ? FontWeights.Bold : FontWeights.Normal;
    }

    public object[] ConvertBack(object? value, Type[] targetTypes, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>FontStyle preview: quote/subtitle styles or the Italic flag.</summary>
public sealed class StyleItalicConverter : IMultiValueConverter
{
    public object Convert(object?[] values, Type targetType, object? parameter, CultureInfo culture)
    {
        var style = values.ElementAtOrDefault(0) as string ?? "";
        var italic = values.ElementAtOrDefault(1) as bool? ?? false;
        var styleItalic = style is ParagraphStyles.Quote or ParagraphStyles.Subtitle;
        return italic || styleItalic ? FontStyles.Italic : FontStyles.Normal;
    }

    public object[] ConvertBack(object? value, Type[] targetTypes, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
