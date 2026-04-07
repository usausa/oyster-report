namespace OysterReport.Generator.Models;

internal readonly record struct ReportOffset
{
    public double X { get; init; } // X offset (points)

    public double Y { get; init; } // Y offset (points)
}
