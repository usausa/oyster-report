namespace OysterReport.Internal;

using System.Globalization;
using System.Runtime.CompilerServices;

internal static class AddressHelper
{
    private const string Digits2 =
        "00010203040506070809101112131415161718192021222324252627282930313233343536373839" +
        "40414243444546474849505152535455565758596061626364656667686970717273747576777879" +
        "8081828384858687888990919293949596979899";

    [SkipLocalsInit]
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
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
        var table = Digits2.AsSpan();
        var rowStart = rowBuffer.Length;
        var value = row;
        while (value >= 100)
        {
            var quotient = value / 100;
            var remainder = value - (quotient * 100);
            rowStart -= 2;
            table.Slice(remainder * 2, 2).CopyTo(rowBuffer[rowStart..]);
            value = quotient;
        }

        if (value >= 10)
        {
            rowStart -= 2;
            table.Slice(value * 2, 2).CopyTo(rowBuffer[rowStart..]);
        }
        else
        {
            rowBuffer[--rowStart] = (char)('0' + value);
        }

        var columnLength = colBuffer.Length - colStart;
        var rowLength = rowBuffer.Length - rowStart;
        Span<char> addressBuffer = stackalloc char[columnLength + rowLength];
        colBuffer[colStart..].CopyTo(addressBuffer);
        rowBuffer[rowStart..].CopyTo(addressBuffer[columnLength..]);
        return new string(addressBuffer);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
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
