namespace OysterReport.Generator.Models;

internal sealed record PdfRenderSheetPlan
{
    public string SheetName { get; init; } = string.Empty; // Target sheet name

    public IReadOnlyList<PdfRenderPagePlan> Pages { get; init; } = Array.Empty<PdfRenderPagePlan>(); // Pages after layout resolution

    public IReadOnlyList<PdfImageRenderInfo> Images { get; init; } = Array.Empty<PdfImageRenderInfo>(); // Final image placements
}
