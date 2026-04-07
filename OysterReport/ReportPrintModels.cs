namespace OysterReport;

using OysterReport.Internal;

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

public sealed class ReportImage
{
    public ReportImage(
        string name,
        ReportAnchorType anchorType,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        AnchorType = anchorType;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;

        var (row, _) = AddressHelper.ParseAddress(fromCellAddress);
        FromRow = row;
    }

    public string Name { get; } // Image identifier name

    public ReportAnchorType AnchorType { get; } // Anchor type

    public string FromCellAddress { get; private set; } // Starting cell address

    public string? ToCellAddress { get; private set; } // Ending cell address

    public ReportOffset Offset { get; } // Offset within the starting cell

    public double WidthPoint { get; } // Image width (points)

    public double HeightPoint { get; } // Image height (points)

    public ReadOnlyMemory<byte> ImageBytes { get; } // Raw image data

    internal int FromRow { get; private set; } // Starting row number (1-based)

    internal ReportImage CloneShifted(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        var shiftedFrom = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        string? shiftedTo = null;
        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            shiftedTo = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }

        return new ReportImage(Name, AnchorType, shiftedFrom, shiftedTo, Offset, WidthPoint, HeightPoint, ImageBytes);
    }

    internal void ShiftRows(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        FromCellAddress = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        FromRow += rowOffset;

        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            ToCellAddress = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }
    }
}
