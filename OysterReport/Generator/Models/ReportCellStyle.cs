namespace OysterReport.Generator.Models;

internal sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new(); // Text font settings

    public ReportFill Fill { get; init; } = new(); // Background fill settings

    public ReportBorders Borders { get; init; } = new(); // Border settings for all four sides

    public ReportAlignment Alignment { get; init; } = new();

    public bool WrapText { get; init; }
}
