using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

// Register code page provider for ExcelDataReader
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var parser = new ArgumentParser();
var arguments = parser.Parse(args);

if (arguments.ShowHelp)
{
    parser.ShowHelp();
    return 0;
}

if (arguments.InputFile == null)
{
    Console.WriteLine();
    parser.ShowHelp();
    return 1;
}

try
{
    var extractor = new LynxExtractor();
    var result = await extractor.ExtractAsync(arguments.InputFile, arguments.SheetName, arguments.OutputFile);
    
    Console.WriteLine($"\n✓ Created {arguments.OutputFile} with {result.Count} entries");
    Console.WriteLine("\nFirst 10 entries:");
    foreach (var entry in result.Take(10))
    {
        Console.WriteLine($"  {entry.Helmet}, {entry.LastName}, {entry.FirstName}, {entry.Club}");
    }
    
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

public class ArgumentParser
{
    public class Arguments
    {
        public string? InputFile { get; set; }
        public string? SheetName { get; set; }
        public string OutputFile { get; set; } = "Lynx.ppl";
        public bool ShowHelp { get; set; }
    }

    public Arguments Parse(string[] args)
    {
        var result = new Arguments();
        
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLower())
            {
                case "-h":
                case "--help":
                case "/?":
                    result.ShowHelp = true;
                    break;
                case "-s":
                case "--sheet":
                    if (i + 1 < args.Length)
                    {
                        result.SheetName = args[++i];
                    }
                    break;
                case "-o":
                case "--output":
                    if (i + 1 < args.Length)
                    {
                        result.OutputFile = args[++i];
                    }
                    break;
                default:
                    if (!args[i].StartsWith("-") && result.InputFile == null)
                    {
                        result.InputFile = args[i];
                    }
                    break;
            }
        }
        
        // Auto-detect input file if not provided
        if (result.InputFile == null)
        {
            var xlsFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xls")
                .Concat(Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx"))
                .ToList();
            
            if (xlsFiles.Count == 0)
            {
                // No file found - return null to trigger help message
                result.InputFile = null;
            }
            else if (xlsFiles.Count == 1)
            {
                result.InputFile = xlsFiles[0];
                Console.WriteLine($"Using found file: {Path.GetFileName(result.InputFile)}");
            }
            else
            {
                Console.WriteLine("Multiple Excel files found. Please specify which one to use:");
                for (int i = 0; i < xlsFiles.Count; i++)
                {
                    Console.WriteLine($"  {i + 1}. {Path.GetFileName(xlsFiles[i])}");
                }
                throw new Exception("Multiple Excel files found. Please specify input file.");
            }
        }
        
        return result;
    }

    public void ShowHelp()
    {
        Console.WriteLine("PplExtractor - Extract data from Excel to create Lynx.ppl files");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  PplExtractor [input_file] [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -s, --sheet <name>    Sheet name to read from (default: auto-detect 'Club' sheet)");
        Console.WriteLine("  -o, --output <file>   Output file name (default: Lynx.ppl)");
        Console.WriteLine("  -h, --help            Show this help message");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  PplExtractor \"OPC W2 - Hamilton - Distribution 2.xls\"");
        Console.WriteLine("  PplExtractor input.xls --sheet \"Club\"");
        Console.WriteLine("  PplExtractor input.xls --output custom_name.ppl");
    }
}

public class ParticipantRecord
{
    public int Helmet { get; set; }
    public string LastName { get; set; } = string.Empty;
    public string FirstName { get; set; } = string.Empty;
    public string Club { get; set; } = string.Empty;
}

public class LynxExtractor
{
    public async Task<List<ParticipantRecord>> ExtractAsync(string inputFile, string? sheetName, string outputFile)
    {
        if (!File.Exists(inputFile))
        {
            throw new FileNotFoundException($"File '{inputFile}' not found.");
        }

        Console.WriteLine($"\nReading from: {Path.GetFileName(inputFile)}");

        // Read Excel file - read without headers first, then find the header row automatically
        DataTable finalTable;
        
        using var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        
        // Read without headers to get all rows
        var dataset = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = false
            }
        });

        var sheet = FindSheet(dataset, sheetName);
        Console.WriteLine($"Sheet: {sheet.TableName}");
        
        // Find the header row by searching for rows containing expected column names
        int headerRowIndex = FindHeaderRow(sheet);
        
        if (headerRowIndex >= 0)
        {
            Console.WriteLine($"Found header row at index {headerRowIndex}");
            
            // Create new table with proper column names from the detected header row
            finalTable = new DataTable(sheet.TableName);
            
            // Set column names from the header row
            var headerRow = sheet.Rows[headerRowIndex];
            foreach (DataColumn col in sheet.Columns)
            {
                var headerName = headerRow[col]?.ToString()?.Trim() ?? $"Column{col.Ordinal + 1}";
                finalTable.Columns.Add(headerName, typeof(string));
            }
            
            // Copy data rows (skip rows up to and including the header row)
            for (int i = headerRowIndex + 1; i < sheet.Rows.Count; i++)
            {
                var newRow = finalTable.NewRow();
                for (int j = 0; j < sheet.Columns.Count && j < finalTable.Columns.Count; j++)
                {
                    newRow[j] = sheet.Rows[i][j]?.ToString() ?? "";
                }
                finalTable.Rows.Add(newRow);
            }
        }
        else
        {
            // Fallback: use row 0 as headers
            Console.WriteLine("Could not auto-detect header row, using row 0");
            stream.Position = 0;
            using var reader2 = ExcelReaderFactory.CreateReader(stream);
            var dataset2 = reader2.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });
            finalTable = FindSheet(dataset2, sheetName);
        }
        
        Console.WriteLine($"Columns found: {string.Join(", ", finalTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}");

        // Find columns by trying various possible names (case-insensitive)
        var helmetCol = FindColumn(finalTable, new[] { "Helmet", "Helmet Number", "Helmet #", "Number" });
        var lastNameCol = FindColumn(finalTable, new[] { "Last Name", "LastName", "Last", "Surname" });
        var firstNameCol = FindColumn(finalTable, new[] { "First Name", "FirstName", "First", "Given Name" });
        var clubCol = FindColumn(finalTable, new[] { "Club", "Club Name", "Team", "Organization" });

        // Check if we found all required columns
        var missing = new List<string>();
        if (helmetCol == null) missing.Add("Helmet/Number");
        if (lastNameCol == null) missing.Add("Last Name");
        if (firstNameCol == null) missing.Add("First Name");
        if (clubCol == null) missing.Add("Club");

        if (missing.Any())
        {
            Console.WriteLine($"\nError: Could not find required columns: {string.Join(", ", missing)}");
            Console.WriteLine($"Available columns: {string.Join(", ", finalTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}");
            Console.WriteLine("\nPlease check the spreadsheet and update the column name mappings if needed.");
            throw new Exception($"Missing required columns: {string.Join(", ", missing)}");
        }

        Console.WriteLine($"\nUsing columns:");
        Console.WriteLine($"  Helmet: {helmetCol!.ColumnName}");
        Console.WriteLine($"  Last Name: {lastNameCol!.ColumnName}");
        Console.WriteLine($"  First Name: {firstNameCol!.ColumnName}");
        Console.WriteLine($"  Club: {clubCol!.ColumnName}");

        // Extract data
        var records = new List<ParticipantRecord>();
        foreach (DataRow row in finalTable.Rows)
        {
            var helmetStr = row[helmetCol]?.ToString()?.Trim();
            var lastName = row[lastNameCol]?.ToString()?.Trim();
            var firstName = row[firstNameCol]?.ToString()?.Trim();
            var club = row[clubCol]?.ToString()?.Trim();

            // Skip blank rows
            if (string.IsNullOrWhiteSpace(helmetStr) && 
                string.IsNullOrWhiteSpace(lastName) && 
                string.IsNullOrWhiteSpace(firstName) && 
                string.IsNullOrWhiteSpace(club))
            {
                continue;
            }

            // Skip rows missing required data
            if (string.IsNullOrWhiteSpace(helmetStr) || 
                string.IsNullOrWhiteSpace(lastName) || 
                string.IsNullOrWhiteSpace(firstName) || 
                string.IsNullOrWhiteSpace(club))
            {
                continue;
            }

            // Parse helmet number
            if (!int.TryParse(helmetStr, out int helmet))
            {
                continue; // Skip rows with invalid helmet numbers
            }

            records.Add(new ParticipantRecord
            {
                Helmet = helmet,
                LastName = lastName!,
                FirstName = firstName!,
                Club = club!
            });
        }

        // Write CSV with UTF-16 encoding
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = false
        };

        using var writer = new StreamWriter(outputFile, false, new UnicodeEncoding(bigEndian: false, byteOrderMark: true));
        using var csv = new CsvWriter(writer, config);
        
        foreach (var record in records)
        {
            csv.WriteField(record.Helmet.ToString().Trim());
            csv.WriteField(record.LastName.Trim());
            csv.WriteField(record.FirstName.Trim());
            csv.WriteField(record.Club.Trim());
            await csv.NextRecordAsync();
        }

        return records;
    }

    private DataTable FindSheet(DataSet dataset, string? sheetName)
    {
        if (sheetName != null)
        {
            if (dataset.Tables.Contains(sheetName))
            {
                return dataset.Tables[sheetName]!;
            }
            Console.WriteLine($"Warning: Sheet '{sheetName}' not found. Available sheets: {string.Join(", ", dataset.Tables.Cast<DataTable>().Select(t => t.TableName))}");
            Console.WriteLine("Trying to find a sheet containing 'Club'...");
        }

        // Try to find a sheet containing "Club" (case-insensitive)
        foreach (DataTable table in dataset.Tables)
        {
            if (table.TableName.Contains("Club", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Found sheet: '{table.TableName}'");
                return table;
            }
        }

        // If no "Club" sheet found, use the first sheet
        if (dataset.Tables.Count > 0)
        {
            var firstSheet = dataset.Tables[0];
            Console.WriteLine($"No 'Club' sheet found. Using first sheet: '{firstSheet.TableName}'");
            return firstSheet;
        }

        throw new Exception("No sheets found in Excel file.");
    }

    private DataColumn? FindColumn(DataTable table, string[] possibleNames)
    {
        var columnsLower = table.Columns.Cast<DataColumn>()
            .ToDictionary(col => col.ColumnName.ToLowerInvariant(), col => col);

        foreach (var name in possibleNames)
        {
            if (columnsLower.TryGetValue(name.ToLowerInvariant(), out var column))
            {
                return column;
            }
        }

        return null;
    }

    private int FindHeaderRow(DataTable sheet)
    {
        // Expected column name patterns (case-insensitive)
        var helmetPatterns = new[] { "helmet", "number", "#" };
        var lastNamePatterns = new[] { "last", "surname" };
        var firstNamePatterns = new[] { "first", "given" };
        var clubPatterns = new[] { "club", "team", "organization" };

        // Search through rows (up to first 20 rows to avoid scanning entire large files)
        int maxRowsToCheck = Math.Min(20, sheet.Rows.Count);
        
        for (int rowIndex = 0; rowIndex < maxRowsToCheck; rowIndex++)
        {
            var row = sheet.Rows[rowIndex];
            var rowValues = new List<string>();
            
            // Collect all non-empty values from this row
            foreach (DataColumn col in sheet.Columns)
            {
                var value = row[col]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(value))
                {
                    rowValues.Add(value.ToLowerInvariant());
                }
            }

            // Check if this row contains patterns matching our expected columns
            bool hasHelmet = rowValues.Any(v => helmetPatterns.Any(p => v.Contains(p, StringComparison.OrdinalIgnoreCase)));
            bool hasLastName = rowValues.Any(v => lastNamePatterns.Any(p => v.Contains(p, StringComparison.OrdinalIgnoreCase)));
            bool hasFirstName = rowValues.Any(v => firstNamePatterns.Any(p => v.Contains(p, StringComparison.OrdinalIgnoreCase)));
            bool hasClub = rowValues.Any(v => clubPatterns.Any(p => v.Contains(p, StringComparison.OrdinalIgnoreCase)));

            // If we find at least 3 out of 4 expected columns, this is likely the header row
            int matchCount = (hasHelmet ? 1 : 0) + (hasLastName ? 1 : 0) + (hasFirstName ? 1 : 0) + (hasClub ? 1 : 0);
            
            if (matchCount >= 3)
            {
                return rowIndex;
            }
        }

        // If not found, return -1 to indicate header row not found
        return -1;
    }
}
