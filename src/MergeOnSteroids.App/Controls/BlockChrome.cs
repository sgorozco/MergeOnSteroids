using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MergeOnSteroids.App.ViewModels;
using MergeOnSteroids.Core;
using MergeOnSteroids.Core.Blocks;

namespace MergeOnSteroids.App.Controls;

/// <summary>
/// Shared visual shell of every block: colored rounded body, selection halo,
/// drag-to-move, and a context menu (delete / duplicate / move).
/// </summary>
public sealed class BlockChrome : ContentControl
{
    public static readonly DependencyProperty BlockBrushProperty = DependencyProperty.Register(
        nameof(BlockBrush), typeof(Brush), typeof(BlockChrome), new PropertyMetadata(Brushes.Gray));

    public static readonly DependencyProperty BlockBorderBrushProperty = DependencyProperty.Register(
        nameof(BlockBorderBrush), typeof(Brush), typeof(BlockChrome), new PropertyMetadata(Brushes.DimGray));

    public static readonly DependencyProperty IsSelectedProperty = DependencyProperty.Register(
        nameof(IsSelected), typeof(bool), typeof(BlockChrome), new PropertyMetadata(false));

    public Brush BlockBrush
    {
        get => (Brush)GetValue(BlockBrushProperty);
        set => SetValue(BlockBrushProperty, value);
    }

    public Brush BlockBorderBrush
    {
        get => (Brush)GetValue(BlockBorderBrushProperty);
        set => SetValue(BlockBorderBrushProperty, value);
    }

    public bool IsSelected
    {
        get => (bool)GetValue(IsSelectedProperty);
        set => SetValue(IsSelectedProperty, value);
    }

    private Point? _dragStart;

    public BlockChrome()
    {
        Loaded += (_, _) =>
        {
            if (MainViewModel.Current is { } vm)
            {
                vm.SelectionChanged += OnSelectionChanged;
                OnSelectionChanged();
            }
        };
        Unloaded += (_, _) =>
        {
            if (MainViewModel.Current is { } vm)
                vm.SelectionChanged -= OnSelectionChanged;
        };
    }

    private Block? Block => DataContext as Block;

    private void OnSelectionChanged() =>
        IsSelected = Block is not null && ReferenceEquals(MainViewModel.Current?.SelectedBlock, Block);

    // ------------------------------------------------------ select and drag

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        if (Block is null) return;
        if (MainViewModel.Current is { } vm) vm.SelectedBlock = Block;
        _dragStart = e.GetPosition(this);
        e.Handled = true; // keep outer container blocks from stealing the click
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonUp(e);
        _dragStart = null;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_dragStart is not { } start || e.LeftButton != MouseButtonState.Pressed || Block is null)
            return;

        var pos = e.GetPosition(this);
        if (Math.Abs(pos.X - start.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(pos.Y - start.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        _dragStart = null;
        var data = new DataObject(DropZone.BlockFormat, Block);
        DragDrop.DoDragDrop(this, data, DragDropEffects.Move);
        e.Handled = true;
    }

    // ----------------------------------------------------------- context menu

    protected override void OnContextMenuOpening(ContextMenuEventArgs e)
    {
        base.OnContextMenuOpening(e);
        if (Block is null) return;

        var menu = new ContextMenu();
        menu.Items.Add(MenuItem("Duplicate", DuplicateBlock));
        menu.Items.Add(MenuItem("Move up", () => MoveBlock(-1)));
        menu.Items.Add(MenuItem("Move down", () => MoveBlock(+1)));
        menu.Items.Add(new Separator());
        menu.Items.Add(MenuItem("Delete", DeleteBlock));
        ContextMenu = menu;
    }

    private static MenuItem MenuItem(string header, Action action)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) => action();
        return item;
    }

    private void DuplicateBlock()
    {
        if (Block?.ParentCollection is not { } parent) return;
        var index = parent.IndexOf(Block);
        parent.Insert(index + 1, ProgramSerializer.CloneBlock(Block));
    }

    private void MoveBlock(int delta)
    {
        if (Block?.ParentCollection is not { } parent) return;
        var index = parent.IndexOf(Block);
        var target = index + delta;
        if (target >= 0 && target < parent.Count)
            parent.Move(index, target);
    }

    private void DeleteBlock()
    {
        if (Block?.ParentCollection is { } parent && Block is { } b)
        {
            parent.Remove(b);
            if (MainViewModel.Current is { } vm && ReferenceEquals(vm.SelectedBlock, b))
                vm.SelectedBlock = null;
        }
    }
}
