namespace OysterReport.Generator.Models;

using ClosedXML.Excel;

internal sealed record ReportPageSetup
{
    public XLPaperSize PaperSize { get; init; } = XLPaperSize.A4Paper;

    public XLPageOrientation Orientation { get; init; } = XLPageOrientation.Default;

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d }; // Page body margins

    public double HeaderMarginPoint { get; init; } = 18d; // Header margin (points)

    public double FooterMarginPoint { get; init; } = 18d; // Footer margin (points)

    public int ScalePercent { get; init; } = 100; // Print scale percentage

    public int? FitToPagesWide { get; init; } // Target page count in horizontal direction

    public int? FitToPagesTall { get; init; } // Target page count in vertical direction

    public bool CenterHorizontally { get; init; } // Center horizontally on page flag

    public bool CenterVertically { get; init; } // Center vertically on page flag
}
