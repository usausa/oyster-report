namespace OysterReport.Internal;

using System.Globalization;

internal static class AddressHelper
{
    public static string ToAddress(int row, int column)
    {
        Span<char> colBuffer = stackalloc char[8];
        var colStart = colBuffer.Length;
        var current = column;
        while (current > 0)
        {
            current--;
            colBuffer[--colStart] = (char)('A' + (current % 26));
            current /= 26;
        }

        using var sb = new ValueStringBuilder(stackalloc char[16]);
        sb.Append(colBuffer[colStart..]);
        sb.Append(row.ToString(CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    public static (int Row, int Column) ParseAddress(string address)
    {
        using var letters = new ValueStringBuilder(stackalloc char[16]);
        using var digits = new ValueStringBuilder(stackalloc char[8]);

        foreach (var character in address.Trim().ToUpperInvariant())
        {
            if (Char.IsLetter(character))
            {
                letters.Append(character);
            }
            else if (Char.IsDigit(character))
            {
                digits.Append(character);
            }
        }

        var lettersStr = letters.ToString();
        var digitsStr = digits.ToString();

        if (String.IsNullOrEmpty(lettersStr) || String.IsNullOrEmpty(digitsStr))
        {
            throw new FormatException(String.Create(CultureInfo.InvariantCulture, $"Invalid cell address. address=[{address}]"));
        }

        var column = 0;
        foreach (var character in lettersStr)
        {
            column = (column * 26) + (character - 'A' + 1);
        }

        return (Int32.Parse(digitsStr, CultureInfo.InvariantCulture), column);
    }
}
