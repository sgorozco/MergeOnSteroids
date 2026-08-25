using System.IO;
using System.Text.Json;

namespace MergeOnSteroids.App.Common;

/// <summary>
/// The handful of editor preferences that outlive a session, kept in
/// %APPDATA%\MergeOnSteroids\settings.json. Losing them is never worth an error,
/// so every failure here is swallowed and the defaults stand.
/// </summary>
public sealed record UserSettings
{
    /// <summary>"Light" or "Dark"; null on the first run, when Windows decides.</summary>
    public string? Theme { get; init; }

    /// <summary>Render paragraph previews through Word instead of approximating them.</summary>
    public bool WordPreviews { get; init; }

    private static string FilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "MergeOnSteroids", "settings.json");

    public static UserSettings Load()
    {
        try
        {
            return File.Exists(FilePath)
                ? JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(FilePath)) ?? new UserSettings()
                : new UserSettings();
        }
        catch (Exception)
        {
            return new UserSettings();
        }
    }

    /// <summary>Reads, changes and writes back — so one setting cannot drop another.</summary>
    public static void Update(Func<UserSettings, UserSettings> change)
    {
        try
        {
            var updated = change(Load());
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath,
                JsonSerializer.Serialize(updated, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception)
        {
            // a remembered preference is a nicety, not worth an error dialog
        }
    }
}
