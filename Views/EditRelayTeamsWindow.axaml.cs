using System.Collections.Generic;
using Avalonia.Controls;
using Avalonia.Interactivity;
using PplExtractor.Models;
using PplExtractor.ViewModels;

namespace PplExtractor.Views;

public partial class EditRelayTeamsWindow : Window
{
    public EditRelayTeamsWindow()
    {
        InitializeComponent();
    }

    private void OkButton_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is EditRelayTeamsViewModel vm)
        {
            Close(vm.GetTeamsCopy());
        }
        else
        {
            Close(null);
        }
    }

    private void CancelButton_Click(object? sender, RoutedEventArgs e)
    {
        Close(null);
    }
}
