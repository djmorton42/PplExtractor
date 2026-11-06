namespace PplExtractor.Models;

/// <summary>
/// How a first-name cell that may contain a parenthetical given name is written to the PPL file.
/// </summary>
public enum FirstNameFormat
{
    /// <summary>Write the cell exactly as it appears in the spreadsheet.</summary>
    All,

    /// <summary>Write the name outside the parentheses (the name they commonly use).</summary>
    Common,

    /// <summary>Write only the name inside the parentheses (their given name).</summary>
    Given
}
