# PplExtractor (.NET Version)

A .NET console application that extracts data from Excel spreadsheets to create `Lynx.ppl` files in the correct format for speed skating competitions.

## Features

- Cross-platform (Windows, macOS, Linux)
- Single-file executable support
- Auto-detects Excel files and "Club" sheet
- Flexible column name matching
- UTF-16 CSV output

## Requirements

- .NET 9.0 SDK (or .NET 8.0 SDK)
- Windows 10/11 or macOS

## Building

### Development Build

```bash
dotnet build
```

### Single-File Executable (Windows)

```bash
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
```

The executable will be in: `bin/Release/net9.0/win-x64/publish/PplExtractor.exe`

### Single-File Executable (macOS)

```bash
dotnet publish -c Release -r osx-arm64 --self-contained false -p:PublishSingleFile=true
```

For Intel Macs:
```bash
dotnet publish -c Release -r osx-x64 --self-contained false -p:PublishSingleFile=true
```

The executable will be in: `bin/Release/net9.0/osx-arm64/publish/PplExtractor` (or `osx-x64` for Intel)

### Self-Contained (No .NET Runtime Required)

To create a self-contained executable that doesn't require .NET to be installed:

**Windows:**
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

**macOS:**
```bash
dotnet publish -c Release -r osx-arm64 --self-contained true -p:PublishSingleFile=true
```

Note: Self-contained executables are larger (typically 50-100 MB) but don't require .NET to be installed on the target machine.

## Usage

### Basic Usage

```bash
PplExtractor "OPC W2 - Hamilton - Distribution 2.xls"
```

### Auto-Detect Excel File

If there's only one `.xls` or `.xlsx` file in the current directory:

```bash
PplExtractor
```

### Specify Sheet Name

```bash
PplExtractor input.xls --sheet "Club"
```

### Custom Output File

```bash
PplExtractor input.xls --output custom_name.ppl
```

### Help

```bash
PplExtractor --help
```

## Command-Line Options

- `[input_file]` - Input Excel file (.xls or .xlsx). If not provided, auto-detects Excel files in current directory.
- `-s, --sheet <name>` - Sheet name to read from (default: auto-detect "Club" sheet)
- `-o, --output <file>` - Output file name (default: Lynx.ppl)
- `-h, --help` - Show help message

## Output Format

The `Lynx.ppl` file is a CSV file (no header row) with UTF-16 encoding:

```
helmet number, last name, first name, club
```

Example:
```
466,Broad,Penelope,Barrie
1367,Mandal,Ishaan,Barrie
```

## Column Name Matching

The application automatically finds columns by trying multiple name variations (case-insensitive):

- **Helmet**: "Helmet", "Helmet Number", "Helmet #", "Number"
- **Last Name**: "Last Name", "LastName", "Last", "Surname"
- **First Name**: "First Name", "FirstName", "First", "Given Name"
- **Club**: "Club", "Club Name", "Team", "Organization"

## Dependencies

- **ExcelDataReader** (3.6.0) - For reading Excel files (.xls and .xlsx)
- **ExcelDataReader.DataSet** (3.6.0) - Dataset support for ExcelDataReader
- **CsvHelper** (33.0.1) - For writing CSV files
- **System.Text.Encoding.CodePages** (8.0.0) - Required for ExcelDataReader to handle .xls files

## Troubleshooting

### "Could not find required columns"

The application couldn't find the expected column names. It will show you what columns are available. You may need to:
- Check that your Excel file has columns for Helmet, First Name, Last Name, and Club
- The application tries to match common variations, but if your columns are named very differently, you may need to modify the code

### "No 'Club' sheet found"

The application will use the first sheet in your Excel file. If you want to use a different sheet, specify it with the `--sheet` option.

### File Not Found

- Make sure you're in the correct directory
- Check that the Excel file name is spelled correctly
- If the file name has spaces, make sure to put it in quotes: `"File Name.xls"`

## Notes

- The output file `Lynx.ppl` will be created in the current working directory
- If a file named `Lynx.ppl` already exists, it will be overwritten
- The application filters out blank rows automatically
- The output is saved in UTF-16 encoding as required


