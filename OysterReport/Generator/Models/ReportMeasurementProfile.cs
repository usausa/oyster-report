namespace OysterReport.Generator.Models;

internal sealed record ReportMeasurementProfile
{
    public double Dpi { get; init; } = 96d; // DPI assumed during measurement

    public double MaxDigitWidth { get; init; } = 7d; // Maximum digit width in the default font

    public string DefaultFontName { get; init; } = "Arial"; // Default font name

    public double DefaultFontSize { get; init; } = 11d; // Default font size (points)

    public double ColumnWidthAdjustment { get; init; } = 1d; // Adjustment factor for column width conversion
}
