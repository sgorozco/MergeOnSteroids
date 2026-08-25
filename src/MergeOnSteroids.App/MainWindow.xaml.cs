using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using MergeOnSteroids.App.Controls;
using MergeOnSteroids.App.ViewModels;

namespace MergeOnSteroids.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm;
    private Point _paletteDragStart;

    public MainWindow()
    {
        _vm = new MainViewModel();
        DataContext = _vm;
        InitializeComponent();

        var args = Environment.GetCommandLineArgs();
        if (args.Length > 1 && File.Exists(args[1]))
            _vm.TryLoadFile(args[1]);
    }

    // ------------------------------------------------------------ palette

    private void PaletteList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) =>
        _paletteDragStart = e.GetPosition(null);

    private void PaletteList_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed) return;
        var pos = e.GetPosition(null);
        if (Math.Abs(pos.X - _paletteDragStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(pos.Y - _paletteDragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        if (FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject) is not { DataContext: PaletteItem item } container)
            return;

        DragDrop.DoDragDrop(container, new DataObject(DropZone.PaletteFormat, item), DragDropEffects.Copy);
    }

    private void PaletteList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject) is { DataContext: PaletteItem item })
        {
            var block = item.CreateBlock();
            _vm.Program.Blocks.Add(block);
            _vm.SelectedBlock = block;
        }
    }

    // ----------------------------------------------------------- shortcuts

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Delete) return;
        if (Keyboard.FocusedElement is TextBoxBase or ComboBox) return; // typing, not deleting a block
        if (_vm.SelectedBlock is { ParentCollection: { } parent } block)
        {
            parent.Remove(block);
            _vm.SelectedBlock = null;
            e.Handled = true;
        }
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
        if (!_vm.ConfirmDiscard()) e.Cancel = true;
        else _vm.Shutdown();   // closes the hidden Word instance previews were using
    }

    // ------------------------------------------------------------- run panel

    private void LogBox_TextChanged(object sender, TextChangedEventArgs e) => LogBox.ScrollToEnd();

    private void OutputFiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is ListBox { SelectedItem: string path })
        {
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open file:\n{ex.Message}", "Merge on Steroids",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // --------------------------------------------------------------- helpers

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T match) return match;
            current = VisualTreeHelper.GetParent(current);
        }
        return null;
    }
}
