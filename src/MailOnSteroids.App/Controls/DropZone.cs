using System.Windows;
using System.Windows.Controls;
using MailOnSteroids.App.ViewModels;
using MailOnSteroids.Core.Blocks;

namespace MailOnSteroids.App.Controls;

public enum DropMode
{
    /// <summary>Insert before the block this zone's DataContext points at.</summary>
    InsertBefore,
    /// <summary>Append to the BlockCollection this zone's DataContext points at.</summary>
    Append
}

/// <summary>
/// Thin drop target rendered before every block and at the end of every block list.
/// Accepts palette items (create) and existing blocks (move).
/// </summary>
public sealed class DropZone : Control
{
    public const string BlockFormat = "mos-block";
    public const string PaletteFormat = "mos-palette";

    public static readonly DependencyProperty ModeProperty = DependencyProperty.Register(
        nameof(Mode), typeof(DropMode), typeof(DropZone), new PropertyMetadata(DropMode.InsertBefore));

    public static readonly DependencyProperty IsActiveProperty = DependencyProperty.Register(
        nameof(IsActive), typeof(bool), typeof(DropZone), new PropertyMetadata(false));

    public DropMode Mode
    {
        get => (DropMode)GetValue(ModeProperty);
        set => SetValue(ModeProperty, value);
    }

    public bool IsActive
    {
        get => (bool)GetValue(IsActiveProperty);
        set => SetValue(IsActiveProperty, value);
    }

    public DropZone()
    {
        AllowDrop = true;
        Focusable = false;
    }

    private (BlockCollection List, int Index)? ResolveTarget()
    {
        if (Mode == DropMode.Append && DataContext is BlockCollection list)
            return (list, list.Count);
        if (Mode == DropMode.InsertBefore && DataContext is Block block && block.ParentCollection is { } parent)
            return (parent, parent.IndexOf(block));
        return null;
    }

    protected override void OnDragOver(DragEventArgs e)
    {
        base.OnDragOver(e);
        e.Effects = DragDropEffects.None;
        if (ResolveTarget() is { } target)
        {
            if (e.Data.GetDataPresent(PaletteFormat))
                e.Effects = DragDropEffects.Copy;
            else if (e.Data.GetDataPresent(BlockFormat) &&
                     e.Data.GetData(BlockFormat) is Block dragged &&
                     !WouldCreateCycle(dragged, target.List))
                e.Effects = DragDropEffects.Move;
        }
        IsActive = e.Effects != DragDropEffects.None;
        e.Handled = true;
    }

    protected override void OnDragLeave(DragEventArgs e)
    {
        base.OnDragLeave(e);
        IsActive = false;
    }

    protected override void OnDrop(DragEventArgs e)
    {
        base.OnDrop(e);
        IsActive = false;
        e.Handled = true;
        if (ResolveTarget() is not { } target) return;
        var (list, index) = target;

        if (e.Data.GetDataPresent(PaletteFormat) &&
            e.Data.GetData(PaletteFormat) is PaletteItem palette)
        {
            var block = palette.CreateBlock();
            list.Insert(Math.Min(index, list.Count), block);
            if (MainViewModel.Current is { } vm) vm.SelectedBlock = block;
            return;
        }

        if (e.Data.GetDataPresent(BlockFormat) &&
            e.Data.GetData(BlockFormat) is Block dragged)
        {
            if (WouldCreateCycle(dragged, list)) return;
            var oldList = dragged.ParentCollection;
            if (oldList is null) return;
            var oldIndex = oldList.IndexOf(dragged);
            if (ReferenceEquals(oldList, list) && (oldIndex == index || oldIndex == index - 1))
                return; // dropped where it already is

            oldList.RemoveAt(oldIndex);
            if (ReferenceEquals(oldList, list) && oldIndex < index) index--;
            list.Insert(Math.Min(index, list.Count), dragged);
            if (MainViewModel.Current is { } vm) vm.SelectedBlock = dragged;
        }
    }

    /// <summary>True when the target list lives inside the dragged block itself.</summary>
    private static bool WouldCreateCycle(Block dragged, BlockCollection targetList)
    {
        var owner = targetList.Owner;
        while (owner is not null)
        {
            if (ReferenceEquals(owner, dragged)) return true;
            owner = owner.ParentCollection?.Owner;
        }
        return false;
    }
}
