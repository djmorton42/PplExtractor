using Avalonia.Controls;
using Avalonia.Interactivity;
using PplExtractor.ViewModels;

namespace PplExtractor.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void CloseButton_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }

    private async void SelectExcelFile_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            await viewModel.SelectExcelFileAsync(this);
        }
    }

    private async void SelectOutputDirectory_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            await viewModel.SelectOutputDirectoryAsync(this);
        }
    }
}

