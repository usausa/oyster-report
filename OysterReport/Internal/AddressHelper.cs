namespace OysterReport.Internal;

using System.Globalization;

internal static class AddressHelper
{
    public static string ToAddress(int row, int column)
    {
        var current = column;
        var result = string.Empty;
        while (current > 0)
        {
            current--;
            result = String.Concat((char)('A' + (current % 26)), result);
            current /= 26;
        }

        return String.Create(CultureInfo.InvariantCulture, $"{result}{row}");
    }

    public static (int Row, int Column) ParseAddress(string address)
    {
        var letters = string.Empty;
        var digits = string.Empty;

        foreach (var character in address.Trim().ToUpperInvariant())
        {
            if (char.IsLetter(character))
            {
                letters += character;
            }
            else if (char.IsDigit(character))
            {
                digits += character;
            }
        }

        if (String.IsNullOrEmpty(letters) || String.IsNullOrEmpty(digits))
        {
            throw new FormatException(String.Create(CultureInfo.InvariantCulture, $"Invalid cell address '{address}'."));
        }

        var column = 0;
        foreach (var character in letters)
        {
            column = (column * 26) + (character - 'A' + 1);
        }

        return (Int32.Parse(digits, CultureInfo.InvariantCulture), column);
    }
}
