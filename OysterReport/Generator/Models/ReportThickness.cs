namespace OysterReport.Generator.Models;

internal readonly record struct ReportThickness
{
    public double Left { get; init; } // Left margin (points)

    public double Top { get; init; } // Top margin (points)

    public double Right { get; init; } // Right margin (points)

    public double Bottom { get; init; } // Bottom margin (points)

    public static ReportThickness Uniform(double value) =>
        new() { Left = value, Top = value, Right = value, Bottom = value };
}
