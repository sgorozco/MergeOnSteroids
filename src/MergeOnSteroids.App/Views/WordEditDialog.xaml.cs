using System.Windows;

namespace MergeOnSteroids.App.Views;

public partial class WordEditDialog : Window
{
    public WordEditDialog() => InitializeComponent();

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
