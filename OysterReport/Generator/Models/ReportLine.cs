namespace OysterReport.Generator.Models;

internal readonly record struct ReportLine
{
    public double X1 { get; init; } // Start point X coordinate (points)

    public double Y1 { get; init; } // Start point Y coordinate (points)

    public double X2 { get; init; } // End point X coordinate (points)

    public double Y2 { get; init; } // End point Y coordinate (points)
}
