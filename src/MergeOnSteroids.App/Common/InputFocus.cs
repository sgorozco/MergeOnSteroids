using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MergeOnSteroids.App.Common;

/// <summary>How the text in a block input is interpreted, which decides whether an
/// inserted column reference needs {braces} around it.</summary>
public enum InputKind
{
    /// <summary>The whole box is one expression — "c.Balance * 1.16".</summary>
    Expression,

    /// <summary>Text with {expression} placeholders — paragraph text, file names, SQL.</summary>
    Template,
}

/// <summary>
/// Remembers the last block input the user typed in, so clicking a column on a data
/// source block drops the reference straight into it instead of making the user
/// retype the column name.
/// </summary>
public static class InputFocus
{
    public static readonly DependencyProperty KindProperty = DependencyProperty.RegisterAttached(
        "Kind", typeof(InputKind), typeof(InputFocus), new PropertyMetadata(InputKind.Expression));

    public static void SetKind(DependencyObject element, InputKind value) => element.SetValue(KindProperty, value);
    public static InputKind GetKind(DependencyObject element) => (InputKind)element.GetValue(KindProperty);

    private static WeakReference<TextBox>? _last;

    /// <summary>Starts remembering focused text boxes (called once at startup).</summary>
    public static void Track() =>
        EventManager.RegisterClassHandler(typeof(TextBox), UIElement.GotKeyboardFocusEvent,
            new KeyboardFocusChangedEventHandler((sender, _) => _last = new WeakReference<TextBox>((TextBox)sender)));

    /// <summary>
    /// Inserts <paramref name="reference"/> at the caret of the last focused input,
    /// wrapping it in braces when that input holds a template and the caret is not
    /// already inside a placeholder. False when there is nowhere to insert it.
    /// </summary>
    public static bool TryInsert(string reference)
    {
        if (_last is null || !_last.TryGetTarget(out var box)) return false;
        if (!box.IsVisible || PresentationSource.FromVisual(box) is null) return false;

        var caret = box.SelectionStart;
        var text = GetKind(box) == InputKind.Template && !InsidePlaceholder(box.Text, caret)
            ? $"{{{reference}}}"
            : reference;

        box.SelectedText = text;                 // replaces the selection, or inserts at the caret
        box.CaretIndex = caret + text.Length;
        box.Focus();
        return true;
    }

    /// <summary>True when the caret sits between a '{' and its '}'.</summary>
    private static bool InsidePlaceholder(string text, int caret)
    {
        for (var i = Math.Min(caret, text.Length) - 1; i >= 0; i--)
        {
            if (text[i] == '}') return false;
            if (text[i] == '{') return true;
        }
        return false;
    }
}
