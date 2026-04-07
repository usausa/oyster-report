namespace OysterReport.Generator.Models;

internal sealed record PdfHeaderFooterRenderInfo
{
    public string? HeaderText { get; init; } // Header text for this page (null if none)

    public string? FooterText { get; init; } // Footer text for this page (null if none)

    public ReportRect HeaderBounds { get; init; } // Header drawing area

    public ReportRect FooterBounds { get; init; } // Footer drawing area
}
