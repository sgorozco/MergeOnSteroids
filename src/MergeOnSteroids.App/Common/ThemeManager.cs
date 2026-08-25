using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Win32;

namespace MergeOnSteroids.App.Common;

public enum AppTheme
{
    Light,
    Dark,
}

/// <summary>
/// Swaps the theme resource dictionary (slot 0 of the application resources) and
/// remembers the choice between runs. Everything themed refers to those brushes
/// with {DynamicResource}, so the whole UI repaints when the slot is replaced.
/// </summary>
public static class ThemeManager
{
    private const int ThemeSlot = 0;

    /// <summary>Raised after the theme changed, so code-built visuals can re-read their brushes.</summary>
    public static event Action? ThemeChanged;

    public static AppTheme Current { get; private set; } = AppTheme.Light;

    public static bool IsDark => Current == AppTheme.Dark;

    /// <summary>Applies the remembered theme, or the one Windows itself is using on first run.</summary>
    public static void Initialize()
    {
        // every window gets the matching (dark or light) title bar, whenever it opens
        EventManager.RegisterClassHandler(typeof(Window), FrameworkElement.LoadedEvent,
            new RoutedEventHandler((sender, _) => ApplyTitleBar((Window)sender)));

        Apply(LoadSaved() ?? SystemTheme(), save: false);
    }

    public static void Toggle() => Apply(IsDark ? AppTheme.Light : AppTheme.Dark);

    public static void Apply(AppTheme theme, bool save = true)
    {
        var app = Application.Current;
        if (app is null) return;

        var dictionary = new ResourceDictionary
        {
            Source = new Uri($"Themes/{theme}.xaml", UriKind.Relative),
        };

        var merged = app.Resources.MergedDictionaries;
        if (merged.Count > ThemeSlot) merged[ThemeSlot] = dictionary;
        else merged.Insert(ThemeSlot, dictionary);

        Current = theme;

        foreach (Window window in app.Windows) ApplyTitleBar(window);
        ThemeChanged?.Invoke();

        if (save) Save(theme);
    }

    // ------------------------------------------------------------- title bar

    private const int DwmwaUseImmersiveDarkMode = 20;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    /// <summary>Paints the window's caption bar to match the theme (Windows 10 20H1+).</summary>
    private static void ApplyTitleBar(Window window)
    {
        var handle = new WindowInteropHelper(window).Handle;
        if (handle == IntPtr.Zero) return;

        var dark = IsDark ? 1 : 0;
        try
        {
            DwmSetWindowAttribute(handle, DwmwaUseImmersiveDarkMode, ref dark, sizeof(int));
        }
        catch (DllNotFoundException)
        {
            // no desktop composition — the caption simply stays light
        }
    }

    // ----------------------------------------------------------- persistence

    private static AppTheme? LoadSaved() =>
        Enum.TryParse<AppTheme>(UserSettings.Load().Theme, ignoreCase: true, out var theme) ? theme : null;

    private static void Save(AppTheme theme) =>
        UserSettings.Update(s => s with { Theme = theme.ToString() });

    /// <summary>What Windows' own "app mode" setting is, used the very first time we run.</summary>
    private static AppTheme SystemTheme()
    {
        try
        {
            var value = Registry.GetValue(
                @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                "AppsUseLightTheme", 1);
            return value is int light && light == 0 ? AppTheme.Dark : AppTheme.Light;
        }
        catch (Exception)
        {
            return AppTheme.Light;
        }
    }
}
