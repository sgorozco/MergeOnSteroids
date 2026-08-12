using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using MergeOnSteroids.Core.Blocks;

namespace MergeOnSteroids.App.ViewModels;

/// <summary>One draggable entry in the block palette.</summary>
public sealed class PaletteItem(
    string category, string title, string description, string brushKey, Func<Block> factory)
    : INotifyPropertyChanged
{
    public string Category { get; } = category;
    public string Title { get; } = title;
    public string Description { get; } = description;
    public Func<Block> Factory { get; } = factory;

    /// <summary>The block color of the current theme (see Themes/Light.xaml and Themes/Dark.xaml).</summary>
    public Brush Brush =>
        Application.Current?.TryFindResource(brushKey) as Brush ?? Brushes.Gray;

    public Block CreateBlock() => Factory();

    /// <summary>Called after a theme switch so the chip repaints in the new palette.</summary>
    internal void RefreshBrush() =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Brush)));

    public event PropertyChangedEventHandler? PropertyChanged;
}
