namespace OysterReport.Generator.Models;

internal sealed record ReportBorder
{
    public ReportBorderStyle Style { get; init; } // Border style

    public string ColorHex { get; init; } = "#FF000000"; // Border color

    public double Width { get; init; } = 0.5d; // Border width (points)
}
