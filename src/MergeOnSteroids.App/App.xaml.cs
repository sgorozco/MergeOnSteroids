using System.Windows;
using MergeOnSteroids.App.Common;

namespace MergeOnSteroids.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        ThemeManager.Initialize();
        InputFocus.Track();
    }

    /// <summary>
    /// Last chance to close the hidden Word instance the paragraph previews use —
    /// it is a separate process, and nothing else would ever shut it down.
    /// </summary>
    protected override void OnExit(ExitEventArgs e)
    {
        ViewModels.MainViewModel.Current?.Shutdown();
        base.OnExit(e);
    }
}
