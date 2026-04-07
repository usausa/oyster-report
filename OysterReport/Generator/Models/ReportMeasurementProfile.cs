namespace OysterReport.Generator.Models;

internal sealed record ReportMeasurementProfile
{
    public double MaxDigitWidth { get; init; } = 7d;

    public string DefaultFontName { get; init; } = "Arial";

    public double DefaultFontSize { get; init; } = 11d;

    public double ColumnWidthAdjustment { get; init; } = 1d;
}
