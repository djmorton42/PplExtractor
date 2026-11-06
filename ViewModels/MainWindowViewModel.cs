using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelDataReader;
using System.Data;
using PplExtractor.Models;
using PplExtractor.Views;

namespace PplExtractor.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty]
    private string _selectedExcelFile = string.Empty;

    [ObservableProperty]
    private string _selectedOutputDirectory = string.Empty;

    [ObservableProperty]
    private ObservableCollection<string> _availableSheets = new();

    [ObservableProperty]
    private string? _selectedSheet;

    [ObservableProperty]
    private bool _isExtracting;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private ObservableCollection<ParticipantRecord> _extractedRacers = new();

    [ObservableProperty]
    private bool _addRelayTeams;

    [ObservableProperty]
    private FirstNameFormat _firstNameFormat = FirstNameFormat.All;

    public bool IsFirstNameAll
    {
        get => FirstNameFormat == FirstNameFormat.All;
        set { if (value) FirstNameFormat = FirstNameFormat.All; }
    }

    public bool IsFirstNameCommon
    {
        get => FirstNameFormat == FirstNameFormat.Common;
        set { if (value) FirstNameFormat = FirstNameFormat.Common; }
    }

    public bool IsFirstNameGiven
    {
        get => FirstNameFormat == FirstNameFormat.Given;
        set { if (value) FirstNameFormat = FirstNameFormat.Given; }
    }

    partial void OnFirstNameFormatChanged(FirstNameFormat value)
    {
        OnPropertyChanged(nameof(IsFirstNameAll));
        OnPropertyChanged(nameof(IsFirstNameCommon));
        OnPropertyChanged(nameof(IsFirstNameGiven));
    }

    [ObservableProperty]
    private ObservableCollection<RelayTeamEntry> _relayTeams = new(RelayTeamEntry.GetDefaults());

    public bool CanExtract => !string.IsNullOrEmpty(SelectedExcelFile) 
        && !string.IsNullOrEmpty(SelectedOutputDirectory)
        && !string.IsNullOrEmpty(SelectedSheet)
        && !IsExtracting;

    partial void OnSelectedExcelFileChanged(string value)
    {
        OnPropertyChanged(nameof(CanExtract));
        // Clear extracted racers when a new file is selected
        ExtractedRacers.Clear();
        LoadSheets();
    }

    partial void OnSelectedOutputDirectoryChanged(string value)
    {
        OnPropertyChanged(nameof(CanExtract));
    }

    partial void OnSelectedSheetChanged(string? value)
    {
        OnPropertyChanged(nameof(CanExtract));
    }

    partial void OnIsExtractingChanged(bool value)
    {
        OnPropertyChanged(nameof(CanExtract));
    }

    private void LoadSheets()
    {
        if (string.IsNullOrEmpty(SelectedExcelFile) || !File.Exists(SelectedExcelFile))
        {
            AvailableSheets.Clear();
            SelectedSheet = null;
            ExtractedRacers.Clear();
            return;
        }

        try
        {
            AvailableSheets.Clear();
            SelectedSheet = null;

            using var stream = File.Open(SelectedExcelFile, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var dataset = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            });

            foreach (DataTable table in dataset.Tables)
            {
                AvailableSheets.Add(table.TableName);
            }

            // Try to find and select the default sheet (one containing "Club")
            var defaultSheet = AvailableSheets.FirstOrDefault(s => 
                s.Contains("Club", StringComparison.OrdinalIgnoreCase));
            
            if (defaultSheet != null)
            {
                SelectedSheet = defaultSheet;
            }
            // If no default sheet found, user must manually select
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error loading sheets: {ex.Message}";
        }
    }

    public async Task SelectExcelFileAsync(TopLevel topLevel)
    {
        try
        {
            var options = new FilePickerOpenOptions
            {
                Title = "Select Excel File",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("Excel Files")
                    {
                        Patterns = new[] { "*.xls", "*.xlsx" }
                    },
                    new FilePickerFileType("All Files")
                    {
                        Patterns = new[] { "*.*" }
                    }
                }
            };

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(options);
            
            if (files.Count > 0 && files[0].Path != null)
            {
                string? localPath = null;
                try
                {
                    localPath = files[0].Path.LocalPath;
                }
                catch
                {
                    localPath = files[0].Path.ToString();
                }

                if (!string.IsNullOrEmpty(localPath))
                {
                    SelectedExcelFile = localPath;
                }
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error selecting file: {ex.Message}";
        }
    }

    public async Task SelectOutputDirectoryAsync(TopLevel topLevel)
    {
        try
        {
            var options = new FolderPickerOpenOptions
            {
                Title = "Select Output Directory",
                AllowMultiple = false
            };

            var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(options);
            
            if (folders.Count > 0 && folders[0].Path != null)
            {
                string? localPath = null;
                try
                {
                    localPath = folders[0].Path.LocalPath;
                }
                catch
                {
                    localPath = folders[0].Path.ToString();
                }

                if (!string.IsNullOrEmpty(localPath))
                {
                    SelectedOutputDirectory = localPath;
                }
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error selecting directory: {ex.Message}";
        }
    }

    [RelayCommand]
    public async Task ExtractRacersAsync()
    {
        if (!CanExtract)
            return;

        try
        {
            IsExtracting = true;
            StatusMessage = "Extracting racers...";

            var outputFile = Path.Combine(SelectedOutputDirectory, "Lynx.ppl");
            var extractor = new LynxExtractor();
            var result = await extractor.ExtractAsync(
                SelectedExcelFile,
                SelectedSheet,
                outputFile,
                addRelayTeams: AddRelayTeams,
                relayTeams: AddRelayTeams ? RelayTeams : null,
                firstNameFormat: FirstNameFormat);
            
            // Update the extracted racers collection
            ExtractedRacers.Clear();
            foreach (var record in result)
            {
                ExtractedRacers.Add(record);
            }
            
            var entryCount = result.Count + (AddRelayTeams ? RelayTeams.Count : 0);
            StatusMessage = $"✓ Created {Path.GetFileName(outputFile)} with {entryCount} entries";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
        }
        finally
        {
            IsExtracting = false;
        }
    }

    /// <summary>
    /// Opens the Edit Relay Teams dialog. Call from view with the main window as owner.
    /// </summary>
    public async Task OpenEditRelayTeamsAsync(Window owner)
    {
        var initialCopy = RelayTeams.Select(t => t.CreateCopy()).ToList();
        var vm = new EditRelayTeamsViewModel(initialCopy);
        var dialog = new EditRelayTeamsWindow { DataContext = vm };
        var result = await dialog.ShowDialog<List<RelayTeamEntry>?>(owner);
        if (result != null)
        {
            RelayTeams.Clear();
            foreach (var entry in result)
            {
                RelayTeams.Add(entry);
            }
        }
    }
}

