namespace OysterReport;

using System.Globalization;

using OysterReport.Helpers;

// ── Geometry ──────────────────────────────────────────────────────────────────

public readonly record struct ReportRect
{
    public double X { get; init; } // Upper-left X coordinate (points)

    public double Y { get; init; } // Upper-left Y coordinate (points)

    public double Width { get; init; } // Width (points)

    public double Height { get; init; } // Height (points)

    public double Right => X + Width; // Right edge X coordinate (points)

    public double Bottom => Y + Height; // Bottom edge Y coordinate (points)

    public ReportRect Deflate(ReportThickness thickness) =>
        new()
        {
            X = X + thickness.Left,
            Y = Y + thickness.Top,
            Width = Math.Max(0, Width - thickness.Left - thickness.Right),
            Height = Math.Max(0, Height - thickness.Top - thickness.Bottom)
        };

    public static ReportRect Union(ReportRect first, ReportRect second)
    {
        var x = Math.Min(first.X, second.X);
        var y = Math.Min(first.Y, second.Y);
        var right = Math.Max(first.Right, second.Right);
        var bottom = Math.Max(first.Bottom, second.Bottom);
        return new ReportRect
        {
            X = x,
            Y = y,
            Width = right - x,
            Height = bottom - y
        };
    }
}

public readonly record struct ReportThickness
{
    public double Left { get; init; } // Left margin (points)

    public double Top { get; init; } // Top margin (points)

    public double Right { get; init; } // Right margin (points)

    public double Bottom { get; init; } // Bottom margin (points)

    public static ReportThickness Uniform(double value) =>
        new() { Left = value, Top = value, Right = value, Bottom = value };
}

public readonly record struct ReportLine
{
    public double X1 { get; init; } // Start point X coordinate (points)

    public double Y1 { get; init; } // Start point Y coordinate (points)

    public double X2 { get; init; } // End point X coordinate (points)

    public double Y2 { get; init; } // End point Y coordinate (points)
}

// ── Cell addressing ───────────────────────────────────────────────────────────

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

// ── Cell style ────────────────────────────────────────────────────────────────

public sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new(); // Text font settings

    public ReportFill Fill { get; init; } = new(); // Background fill settings

    public ReportBorders Borders { get; init; } = new(); // Border settings for all four sides

    public ReportAlignment Alignment { get; init; } = new(); // Horizontal and vertical alignment settings

    public string? NumberFormat { get; init; } // Excel number format string

    public bool WrapText { get; init; } // Text wrap flag

    public double Rotation { get; init; } // Text rotation angle

    public bool ShrinkToFit { get; init; } // Whether to shrink text to fit the cell
}

public sealed record ReportFont
{
    public string Name { get; init; } = "Arial"; // Font name

    public double Size { get; init; } = 11d; // Font size (points)

    public bool Bold { get; init; } // Bold flag

    public bool Italic { get; init; } // Italic flag

    public bool Underline { get; init; } // Underline flag

    public bool Strikeout { get; init; } // Strikethrough flag

    public string ColorHex { get; init; } = "#FF000000"; // Text color
}

public sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000"; // Background color
}

public sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new(); // Left border

    public ReportBorder Top { get; init; } = new(); // Top border

    public ReportBorder Right { get; init; } = new(); // Right border

    public ReportBorder Bottom { get; init; } = new(); // Bottom border
}

public sealed record ReportBorder
{
    public ReportBorderStyle Style { get; init; } // Border style

    public string ColorHex { get; init; } = "#FF000000"; // Border color

    public double Width { get; init; } = 0.5d; // Border width (points)
}

public sealed record ReportAlignment
{
    public ReportHorizontalAlignment Horizontal { get; init; } = ReportHorizontalAlignment.General; // Horizontal alignment

    public ReportVerticalAlignment Vertical { get; init; } = ReportVerticalAlignment.Top; // Vertical alignment
}

// ── Page / print settings ─────────────────────────────────────────────────────

public sealed record ReportPageSetup
{
    public ReportPaperSize PaperSize { get; init; } = ReportPaperSize.A4; // Paper size

    public ReportPageOrientation Orientation { get; init; } = ReportPageOrientation.Portrait; // Page orientation

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d }; // Page body margins

    public double HeaderMarginPoint { get; init; } = 18d; // Header margin (points)

    public double FooterMarginPoint { get; init; } = 18d; // Footer margin (points)

    public int ScalePercent { get; init; } = 100; // Print scale percentage

    public int? FitToPagesWide { get; init; } // Target page count in horizontal direction

    public int? FitToPagesTall { get; init; } // Target page count in vertical direction

    public bool CenterHorizontally { get; init; } // Center horizontally on page flag

    public bool CenterVertically { get; init; } // Center vertically on page flag
}

public sealed record ReportHeaderFooter
{
    public bool AlignWithMargins { get; init; } = true; // Whether to align with page margins

    public bool DifferentFirst { get; init; } // Whether the first page uses a different header/footer

    public bool DifferentOddEven { get; init; } // Whether odd and even pages use different headers/footers

    public bool ScaleWithDocument { get; init; } = true; // Whether to scale with the document

    public string? OddHeader { get; init; } // Header text for odd pages

    public string? OddFooter { get; init; } // Footer text for odd pages

    public string? EvenHeader { get; init; } // Header text for even pages

    public string? EvenFooter { get; init; } // Footer text for even pages

    public string? FirstHeader { get; init; } // Header text for the first page

    public string? FirstFooter { get; init; } // Footer text for the first page
}

public sealed record ReportPrintArea
{
    public ReportRange Range { get; init; } // Print area range
}

public sealed record ReportPageBreak
{
    public int Index { get; init; } // Row or column index of the page break

    public bool IsHorizontal { get; init; } // Whether this is a horizontal page break
}

// ── Workbook metadata ─────────────────────────────────────────────────────────

public sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty; // Template name

    public string? SourceFilePath { get; init; } // Source file path

    public DateTimeOffset? SourceLastWriteTime { get; init; } // Last write time of the source file

    public string? Author { get; init; } // Template author
}

public sealed record ReportMeasurementProfile
{
    public double Dpi { get; init; } = 96d; // DPI assumed during measurement

    public double MaxDigitWidth { get; init; } = 7d; // Maximum digit width in the default font

    public string DefaultFontName { get; init; } = "Arial"; // Default font name

    public double DefaultFontSize { get; init; } = 11d; // Default font size (points)

    public double ColumnWidthAdjustment { get; init; } = 1d; // Adjustment factor for column width conversion
}

// ── Cell value ────────────────────────────────────────────────────────────────

public sealed record ReportCellValue
{
    public ReportCellValueKind Kind { get; init; } // Original data type

    public object? RawValue { get; init; } // Raw value retrieved from Excel
}

public sealed record ReportMergeInfo
{
    public string OwnerCellAddress { get; init; } = string.Empty; // Address of the owner cell in the merged range

    public bool IsOwner { get; init; } // Whether this cell is the owner of the merge

    public ReportRange Range { get; init; } // Merged cell range
}
