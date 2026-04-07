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
        StartRow = startRow;
        StartColumn = startColumn;
        EndRow = endRow;
        EndColumn = endColumn;
    }

    public int StartRow { get; init; } // Start row number (1-based)

    public int StartColumn { get; init; } // Start column number (1-based)

    public int EndRow { get; init; } // End row number (1-based)

    public int EndColumn { get; init; } // End column number (1-based)

    public int RowCount => EndRow - StartRow + 1; // Number of rows in range

    public int ColumnCount => EndColumn - StartColumn + 1; // Number of columns in range

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
    public double X { get; init; } // X offset (points)

    public double Y { get; init; } // Y offset (points)
}
