namespace OysterReport.Generator;

using ClosedXML.Excel;

using OysterReport.Helpers;

// ---- Workbook metadata ----

internal sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty;

    public string? SourceFilePath { get; init; }

    public DateTimeOffset? SourceLastWriteTime { get; init; }
}

internal sealed record ReportMeasurementProfile
{
    public double MaxDigitWidth { get; init; } = 7d;

    public string DefaultFontName { get; init; } = "Arial";

    public double DefaultFontSize { get; init; } = 11d;

    public double ColumnWidthAdjustment { get; init; } = 1d;
}

// ---- Cell value and style ----

internal sealed record ReportCellValue
{
    public XLDataType Kind { get; init; } = XLDataType.Blank;

    public object? RawValue { get; init; }
}

internal sealed record ReportFont
{
    public string Name { get; init; } = "Arial";

    public double Size { get; init; } = 11d;

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool Underline { get; init; }

    public bool Strikeout { get; init; }

    public string ColorHex { get; init; } = "#FF000000";
}

internal sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000";
}

internal sealed record ReportBorder
{
    public XLBorderStyleValues Style { get; init; } = XLBorderStyleValues.None;

    public string ColorHex { get; init; } = "#FF000000";

    public double Width { get; init; } = 0.5d;
}

internal sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new();

    public ReportBorder Top { get; init; } = new();

    public ReportBorder Right { get; init; } = new();

    public ReportBorder Bottom { get; init; } = new();
}

internal sealed record ReportAlignment
{
    public XLAlignmentHorizontalValues Horizontal { get; init; } = XLAlignmentHorizontalValues.General;

    public XLAlignmentVerticalValues Vertical { get; init; } = XLAlignmentVerticalValues.Top;
}

internal sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new();

    public ReportFill Fill { get; init; } = new();

    public ReportBorders Borders { get; init; } = new();

    public ReportAlignment Alignment { get; init; } = new();

    public bool WrapText { get; init; }
}

internal sealed record ReportMergeInfo
{
    public string OwnerCellAddress { get; init; } = string.Empty;

    public bool IsOwner { get; init; }

    public ReportRange Range { get; init; }
}

// ---- Page setup ----

internal sealed record ReportPageBreak
{
    public int Index { get; init; }

    public bool IsHorizontal { get; init; }
}

internal sealed record ReportPrintArea
{
    public ReportRange Range { get; init; }
}

internal sealed record ReportHeaderFooter
{
    public bool AlignWithMargins { get; init; } = true;

    public bool DifferentFirst { get; init; }

    public bool DifferentOddEven { get; init; }

    public bool ScaleWithDocument { get; init; } = true;

    public string? OddHeader { get; init; }

    public string? OddFooter { get; init; }

    public string? EvenHeader { get; init; }

    public string? EvenFooter { get; init; }

    public string? FirstHeader { get; init; }

    public string? FirstFooter { get; init; }
}

internal sealed record ReportPageSetup
{
    public XLPaperSize PaperSize { get; init; } = XLPaperSize.A4Paper;

    public XLPageOrientation Orientation { get; init; } = XLPageOrientation.Default;

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d };

    public double HeaderMarginPoint { get; init; } = 18d;

    public double FooterMarginPoint { get; init; } = 18d;

    public int ScalePercent { get; init; } = 100;

    public int? FitToPagesWide { get; init; }

    public int? FitToPagesTall { get; init; }

    public bool CenterHorizontally { get; init; }

    public bool CenterVertically { get; init; }
}

// ---- Sheet structure ----

internal sealed class ReportRow
{
    public ReportRow(int index, double heightPoint, bool isHidden = false, int outlineLevel = 0)
    {
        Index = index;
        HeightPoint = heightPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
    }

    public int Index { get; }

    public double HeightPoint { get; }

    public double TopPoint { get; set; }

    public bool IsHidden { get; }

    public int OutlineLevel { get; }
}

internal sealed class ReportColumn
{
    public ReportColumn(int index, double widthPoint, bool isHidden = false, int outlineLevel = 0, double originalExcelWidth = 0)
    {
        Index = index;
        WidthPoint = widthPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
        OriginalExcelWidth = originalExcelWidth;
    }

    public int Index { get; }

    public double WidthPoint { get; }

    public double LeftPoint { get; set; }

    public bool IsHidden { get; }

    public int OutlineLevel { get; }

    public double OriginalExcelWidth { get; }
}

internal sealed class ReportMergedRange
{
    public ReportMergedRange(ReportRange range)
    {
        Range = range;
        OwnerCellAddress = AddressHelper.ToAddress(range.StartRow, range.StartColumn);
    }

    public ReportRange Range { get; }

    public string OwnerCellAddress { get; }
}

internal sealed class ReportImage
{
    public ReportImage(
        string name,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;
    }

    public string Name { get; }

    public string FromCellAddress { get; }

    public string? ToCellAddress { get; }

    public ReportOffset Offset { get; }

    public double WidthPoint { get; }

    public double HeightPoint { get; }

    public ReadOnlyMemory<byte> ImageBytes { get; }
}

internal sealed class ReportCell
{
    public ReportCell(
        int row,
        int column,
        ReportCellValue value,
        string displayText,
        ReportCellStyle style)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
        Value = value;
        DisplayText = displayText;
        Style = style;
    }

    public int Row { get; }

    public int Column { get; }

    public string Address { get; }

    public ReportCellValue Value { get; }

    public string DisplayText { get; }

    public ReportCellStyle Style { get; set; }

    public ReportMergeInfo? Merge { get; set; }
}

internal sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public ReportSheet(string name)
    {
        Name = name;
        UsedRange = new ReportRange(1, 1, 1, 1);
    }

    public string Name { get; }

    public ReportRange UsedRange { get; set; }

    public IReadOnlyList<ReportRow> Rows => rows;

    public IReadOnlyList<ReportColumn> Columns => columns;

    public IReadOnlyList<ReportCell> Cells => cells;

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges;

    public IReadOnlyList<ReportImage> Images => images;

    public ReportPageSetup PageSetup { get; set; } = new();

    public ReportHeaderFooter HeaderFooter { get; set; } = new();

    public ReportPrintArea? PrintArea { get; set; }

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks;

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks;

    public bool ShowGridLines { get; set; }

    public void AddRowDefinition(ReportRow row) => rows.Add(row);

    public void AddColumnDefinition(ReportColumn column) => columns.Add(column);

    public void AddCell(ReportCell cell) => cells.Add(cell);

    public void AddMergedRange(ReportMergedRange range) => mergedRanges.Add(range);

    public void AddImage(ReportImage image) => images.Add(image);

    public void AddHorizontalPageBreak(ReportPageBreak pageBreak) => horizontalPageBreaks.Add(pageBreak);

    public void AddVerticalPageBreak(ReportPageBreak pageBreak) => verticalPageBreaks.Add(pageBreak);

    public void RecalculateLayout()
    {
        var top = 0d;
        foreach (var row in rows.OrderBy(static row => row.Index))
        {
            row.TopPoint = top;
            top += row.HeightPoint;
        }

        var left = 0d;
        foreach (var column in columns.OrderBy(static column => column.Index))
        {
            column.LeftPoint = left;
            left += column.WidthPoint;
        }
    }
}

internal sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];

    public ReportWorkbook(ReportMetadata? metadata = null, ReportMeasurementProfile? measurementProfile = null)
    {
        Metadata = metadata ?? new ReportMetadata();
        MeasurementProfile = measurementProfile ?? new ReportMeasurementProfile();
    }

    public IReadOnlyList<ReportSheet> Sheets => sheets;

    public ReportMetadata Metadata { get; }

    public ReportMeasurementProfile MeasurementProfile { get; }

    public void AddSheet(ReportSheet sheet) => sheets.Add(sheet);
}
