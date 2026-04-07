namespace OysterReport.Generator.Models;

internal sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new(); // Left border

    public ReportBorder Top { get; init; } = new(); // Top border

    public ReportBorder Right { get; init; } = new(); // Right border

    public ReportBorder Bottom { get; init; } = new(); // Bottom border
}
