namespace OysterReport.Generator.Models;

internal sealed record PdfRenderPlan
{
    public IReadOnlyList<PdfRenderSheetPlan> Sheets { get; init; } = Array.Empty<PdfRenderSheetPlan>(); // Resolved sheet render plans
}
