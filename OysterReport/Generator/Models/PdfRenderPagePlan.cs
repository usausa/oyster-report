namespace OysterReport.Generator.Models;

internal sealed record PdfRenderPagePlan
{
    public int PageNumber { get; init; } // Page number (1-based)

    public ReportRect PageBounds { get; init; } // Rectangle of the entire page

    public ReportRect PrintableBounds { get; init; } // Printable area excluding margins

    public PdfHeaderFooterRenderInfo HeaderFooter { get; init; } = new(); // Header and footer render info for this page

    public IReadOnlyList<PdfCellRenderInfo> Cells { get; init; } = Array.Empty<PdfCellRenderInfo>(); // Cells to render on this page
}
