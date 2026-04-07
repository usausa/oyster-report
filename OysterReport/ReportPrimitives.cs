namespace OysterReport;

using System.Globalization;

using OysterReport.Internal;

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
