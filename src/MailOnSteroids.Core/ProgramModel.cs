using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using MailOnSteroids.Core.Blocks;

namespace MailOnSteroids.Core;

/// <summary>Root of a Mail-on-Steroids program (the whole "scratch script").</summary>
public sealed class ProgramModel : INotifyPropertyChanged
{
    private string _name = "Untitled program";
    private string _outputFolder = "output";

    public string Name { get => _name; set => Set(ref _name, value); }

    /// <summary>Where generated documents go. Relative paths resolve against the program file's folder.</summary>
    public string OutputFolder { get => _outputFolder; set => Set(ref _outputFolder, value); }

    [JsonObjectCreationHandling(JsonObjectCreationHandling.Populate)]
    public BlockCollection Blocks { get; } = new(null, "Program");

    public event PropertyChangedEventHandler? PropertyChanged;

    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public IEnumerable<Block> AllBlocks()
    {
        foreach (var b in Blocks)
            foreach (var d in b.SelfAndDescendants())
                yield return d;
    }
}

public static class ProgramSerializer
{
    public const string FileExtension = ".mos.json";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };

    public static string ToJson(ProgramModel program) => JsonSerializer.Serialize(program, Options);

    public static ProgramModel FromJson(string json) =>
        JsonSerializer.Deserialize<ProgramModel>(json, Options)
        ?? throw new InvalidOperationException("Program file is empty or invalid.");

    public static void Save(ProgramModel program, string path) => File.WriteAllText(path, ToJson(program));

    public static ProgramModel Load(string path) => FromJson(File.ReadAllText(path));

    /// <summary>Deep-clone a block subtree (used by copy/duplicate).</summary>
    public static Block CloneBlock(Block block)
    {
        var json = JsonSerializer.Serialize(block, Options);
        var clone = JsonSerializer.Deserialize<Block>(json, Options)!;
        foreach (var b in clone.SelfAndDescendants())
            b.Id = Guid.NewGuid().ToString("N");
        return clone;
    }
}
