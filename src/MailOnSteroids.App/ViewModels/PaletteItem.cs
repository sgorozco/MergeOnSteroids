using System.Windows.Media;
using MailOnSteroids.Core.Blocks;

namespace MailOnSteroids.App.ViewModels;

/// <summary>One draggable entry in the block palette.</summary>
public sealed class PaletteItem(
    string category, string title, string description, Brush brush, Func<Block> factory)
{
    public string Category { get; } = category;
    public string Title { get; } = title;
    public string Description { get; } = description;
    public Brush Brush { get; } = brush;
    public Func<Block> Factory { get; } = factory;

    public Block CreateBlock() => Factory();
}
