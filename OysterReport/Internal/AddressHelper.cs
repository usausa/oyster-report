namespace OysterReport.Internal;

using System.Globalization;
using System.Runtime.CompilerServices;

internal static class AddressHelper
{
    [SkipLocalsInit]
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
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

        Span<char> rowBuffer = stackalloc char[11];
        row.TryFormat(rowBuffer, out var rowLength, provider: CultureInfo.InvariantCulture);

        var columnLength = colBuffer.Length - colStart;
        Span<char> addressBuffer = stackalloc char[columnLength + rowLength];
        colBuffer[colStart..].CopyTo(addressBuffer);
        rowBuffer[..rowLength].CopyTo(addressBuffer[columnLength..]);
        return new string(addressBuffer);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static void ParseAddress(string address, out int row, out int column)
    {
        var source = address.AsSpan().Trim();
        var hasLetters = false;
        var hasDigits = false;
        var seenDigits = false;

        column = 0;
        row = 0;

        foreach (var character in source)
        {
            if (Char.IsLetter(character))
            {
                if (seenDigits)
                {
                    throw new FormatException(String.Create(CultureInfo.InvariantCulture, $"Invalid cell address. address=[{address}]"));
                }

                column = (column * 26) + (Char.ToUpperInvariant(character) - 'A' + 1);
                hasLetters = true;
            }
            else if (Char.IsDigit(character))
            {
                row = (row * 10) + (character - '0');
                hasDigits = true;
                seenDigits = true;
            }
        }

        if (!hasLetters || !hasDigits)
        {
            throw new FormatException(String.Create(CultureInfo.InvariantCulture, $"Invalid cell address. address=[{address}]"));
        }
    }
}
