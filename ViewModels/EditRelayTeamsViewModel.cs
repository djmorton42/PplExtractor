using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PplExtractor.Models;

namespace PplExtractor.ViewModels;

public partial class EditRelayTeamsViewModel : ViewModelBase
{
    [ObservableProperty]
    private RelayTeamEntry? _selectedTeam;

    public ObservableCollection<RelayTeamEntry> Teams { get; } = new();

    public EditRelayTeamsViewModel(IEnumerable<RelayTeamEntry> initial)
    {
        foreach (var entry in initial)
        {
            Teams.Add(entry.CreateCopy());
        }
    }

    [RelayCommand]
    private void AddTeam()
    {
        Teams.Add(new RelayTeamEntry("", "", "", "SSO"));
    }

    [RelayCommand(CanExecute = nameof(CanRemoveTeam))]
    private void RemoveTeam()
    {
        if (SelectedTeam != null)
        {
            Teams.Remove(SelectedTeam);
            SelectedTeam = null;
        }
    }

    private bool CanRemoveTeam() => SelectedTeam != null;

    partial void OnSelectedTeamChanged(RelayTeamEntry? value)
    {
        RemoveTeamCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void ResetToDefaults()
    {
        Teams.Clear();
        foreach (var entry in RelayTeamEntry.GetDefaults())
        {
            Teams.Add(entry.CreateCopy());
        }
    }

    /// <summary>
    /// Returns a copy of the current teams list (for dialog result).
    /// </summary>
    public List<RelayTeamEntry> GetTeamsCopy()
    {
        return Teams.Select(t => t.CreateCopy()).ToList();
    }
}
