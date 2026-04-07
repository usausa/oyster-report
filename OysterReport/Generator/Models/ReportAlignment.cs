namespace OysterReport.Generator.Models;

internal sealed record ReportAlignment
{
    public ReportHorizontalAlignment Horizontal { get; init; } = ReportHorizontalAlignment.General; // Horizontal alignment

    public ReportVerticalAlignment Vertical { get; init; } = ReportVerticalAlignment.Top; // Vertical alignment
}
