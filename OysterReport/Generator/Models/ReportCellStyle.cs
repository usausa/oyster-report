namespace OysterReport.Generator.Models;

internal sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new(); // Text font settings

    public ReportFill Fill { get; init; } = new(); // Background fill settings

    public ReportBorders Borders { get; init; } = new(); // Border settings for all four sides

    public ReportAlignment Alignment { get; init; } = new(); // Horizontal and vertical alignment settings

    public string? NumberFormat { get; init; } // Excel number format string

    public bool WrapText { get; init; } // Text wrap flag

    public double Rotation { get; init; } // Text rotation angle

    public bool ShrinkToFit { get; init; } // Whether to shrink text to fit the cell
}
