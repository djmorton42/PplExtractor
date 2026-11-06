using System;
using System.Text;

namespace PplExtractor.Models;

/// <summary>
/// Splits first names such as "Kal (Kalliana)" or "(Kalliana) Kal".
/// The text in parentheses is the given name; the text outside is the common name.
/// Parentheses may appear before or after the common name.
/// When the selected part is missing, the cell is kept as provided so the racer is still written.
/// </summary>
public static class FirstNameFormatter
{
    public static string Apply(string firstName, FirstNameFormat format)
    {
        if (format == FirstNameFormat.All || string.IsNullOrEmpty(firstName))
            return firstName;

        var open = firstName.IndexOf('(');
        var close = open >= 0 ? firstName.IndexOf(')', open + 1) : -1;
        if (open < 0 || close < 0)
            return firstName;

        if (format == FirstNameFormat.Given)
        {
            var given = firstName[(open + 1)..close].Trim();
            return string.IsNullOrEmpty(given) ? firstName : given;
        }

        var withoutBrackets = new StringBuilder(firstName.Length);
        var index = 0;
        while (index < firstName.Length)
        {
            var nextOpen = firstName.IndexOf('(', index);
            if (nextOpen < 0)
            {
                withoutBrackets.Append(firstName.AsSpan(index));
                break;
            }

            var nextClose = firstName.IndexOf(')', nextOpen + 1);
            if (nextClose < 0)
            {
                withoutBrackets.Append(firstName.AsSpan(index));
                break;
            }

            withoutBrackets.Append(firstName.AsSpan(index, nextOpen - index));
            withoutBrackets.Append(' ');
            index = nextClose + 1;
        }

        var common = string.Join(' ', withoutBrackets.ToString().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return string.IsNullOrEmpty(common) ? firstName : common;
    }
}
