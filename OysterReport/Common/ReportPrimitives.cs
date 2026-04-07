namespace OysterReport.Common;

using System.Globalization;

public enum ReportDumpFormat
{
    Json,
    Markdown
}

public enum ReportDiagnosticSeverity
{
    Info,
    Warning,
    Error
}

public enum ReportPaperSize
{
    Custom,
    A4,
    Letter,
    Legal
}

public enum ReportPageOrientation
{
    Portrait,
    Landscape
}

public enum ReportAnchorType
{
    MoveAndSizeWithCells,
    MoveWithCells,
    Absolute
}

public enum ReportHorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Justify
}

public enum ReportVerticalAlignment
{
    Top,
    Center,
    Bottom,
    Justify
}

public enum ReportBorderStyle
{
    None,
    Thin,
    Medium,
    Thick,
    DoubleLine,
    Dashed,
    Dotted,
    Hair,
    DashDot
}

public enum ReportCellValueKind
{
    Blank,
    Text,
    Number,
    DateTime,
    Boolean,
    Error
}

public readonly record struct ReportRange
{
    public ReportRange(int startRow, int startColumn, int endRow, int endColumn)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(startRow);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(startColumn);
        ArgumentOutOfRangeException.ThrowIfLessThan(endRow, startRow);
        ArgumentOutOfRangeException.ThrowIfLessThan(endColumn, startColumn);
        StartRow = startRow;
        StartColumn = startColumn;
        EndRow = endRow;
        EndColumn = endColumn;
    }

    public int StartRow { get; init; } // 開始行番号

    public int StartColumn { get; init; } // 開始列番号

    public int EndRow { get; init; } // 終了行番号

    public int EndColumn { get; init; } // 終了列番号

    public int RowCount => EndRow - StartRow + 1; // 範囲の行数

    public int ColumnCount => EndColumn - StartColumn + 1; // 範囲の列数

    public bool Contains(int row, int column) =>
        row >= StartRow &&
        row <= EndRow &&
        column >= StartColumn &&
        column <= EndColumn;

    public ReportRange ShiftRows(int offset) =>
        new(StartRow + offset, StartColumn, EndRow + offset, EndColumn);

    public override string ToString()
    {
        var startAddress = AddressHelper.ToAddress(StartRow, StartColumn);
        var endAddress = AddressHelper.ToAddress(EndRow, EndColumn);
        return startAddress == endAddress ? startAddress : string.Create(CultureInfo.InvariantCulture, $"{startAddress}:{endAddress}");
    }
}

public readonly record struct ReportOffset
{
    public double X { get; init; } // X オフセット(point)

    public double Y { get; init; } // Y オフセット(point)
}

internal static class AddressHelper
{
    public static string ToAddress(int row, int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(row);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(column);

        var current = column;
        var result = string.Empty;
        while (current > 0)
        {
            current--;
            result = string.Concat((char)('A' + (current % 26)), result);
            current /= 26;
        }

        return string.Create(CultureInfo.InvariantCulture, $"{result}{row}");
    }

    public static (int Row, int Column) ParseAddress(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

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

        if (string.IsNullOrEmpty(letters) || string.IsNullOrEmpty(digits))
        {
            throw new FormatException(string.Create(CultureInfo.InvariantCulture, $"Invalid cell address '{address}'."));
        }

        var column = 0;
        foreach (var character in letters)
        {
            column = (column * 26) + (character - 'A' + 1);
        }

        return (int.Parse(digits, CultureInfo.InvariantCulture), column);
    }
}
