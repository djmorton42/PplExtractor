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

    public bool CanExtract => !string.IsNullOrEmpty(SelectedExcelFile) 
        && !string.IsNullOrEmpty(SelectedOutputDirectory)
        && !string.IsNullOrEmpty(SelectedSheet)
        && !IsExtracting;

    partial void OnSelectedExcelFileChanged(string value)
    {
        OnPropertyChanged(nameof(CanExtract));
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
            var result = await extractor.ExtractAsync(SelectedExcelFile, SelectedSheet, outputFile);
            
            StatusMessage = $"✓ Created {Path.GetFileName(outputFile)} with {result.Count} entries";
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
}

