using System.Windows;

namespace MergeOnSteroids.App.Views;

/// <summary>
/// The little "Save &amp; return" window shown while a fragment is open in Word.
///
/// Deliberately ownerless and modeless: an owned window drags its owner to the front
/// every time you touch it, so nudging this box aside to reach Word buried Word behind
/// the editor. Topmost keeps it findable instead, and it parks itself out of the way.
/// </summary>
public partial class WordEditDialog : Window
{
    private readonly TaskCompletionSource<bool> _completion =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    public WordEditDialog() => InitializeComponent();

    /// <summary>True when the user chose "Save &amp; return"; false if cancelled or closed.</summary>
    public Task<bool> Result => _completion.Task;

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        // top-right of the working area: visible, but not over the text being edited
        var work = SystemParameters.WorkArea;
        Left = work.Right - ActualWidth - 24;
        Top = work.Top + 24;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        _completion.TrySetResult(true);
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    protected override void OnClosed(EventArgs e)
    {
        _completion.TrySetResult(false);   // closing it any other way means "do not save"
        base.OnClosed(e);
    }
}
