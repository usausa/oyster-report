namespace OysterReport;

using System.Diagnostics.CodeAnalysis;

using ClosedXML.Excel;

[ExcludeFromCodeCoverage]
public sealed record ReportRenderOption
{
    // Resolves page size (pt) from the paper size
    public Func<XLPaperSize, (double Width, double Height)> PageSizeResolver { get; set; } = ResolveDefaultPageSize;

    // Horizontal padding for cell text (pt)
    public double HorizontalCellTextPadding { get; set; } = 2d;

    // Default font size for cell text (pt)
    public double DefaultCellFontSize { get; set; } = 11d;

    // Default font size for headers and footers (pt)
    public double HeaderFooterFontSize { get; set; } = 9d;

    // Fallback font candidates used when rendering headers and footers.
    public IReadOnlyList<string> HeaderFooterFallbackFonts { get; set; } =
    [
        "Arial",
        "Segoe UI",
        "Helvetica",
        "Liberation Sans",
        "DejaVu Sans"
    ];

    // Drawing width for Thick borders (pt)
    public double ThickBorderWidth { get; set; } = 2.25d;

    // Drawing width for Medium borders (pt)
    public double MediumBorderWidth { get; set; } = 1.5d;

    // Drawing width for regular borders (pt)
    public double NormalBorderWidth { get; set; } = 0.75d;

    // Drawing width for Hair borders (pt)
    public double HairBorderWidth { get; set; } = 0.25d;

    // Underline drawing width (pt)
    public double UnderlineWidth { get; set; } = 0.5d;

    // Strikeout drawing width (pt)
    public double StrikeoutWidth { get; set; } = 0.5d;

    // Adjustment factor used when converting column widths to points
    public double ColumnWidthAdjustment { get; set; } = 1d;

    // Fallback max digit width for unknown fonts (96-DPI reference pixels)
    public double FallbackMaxDigitWidth { get; set; } = 7d;

    private static (double Width, double Height) ResolveDefaultPageSize(XLPaperSize paperSize) =>
        paperSize switch
        {
            XLPaperSize.LetterPaper => (612d, 792d),
            XLPaperSize.LegalPaper => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
}
