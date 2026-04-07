namespace OysterReport;

public sealed record ReportCellStyle
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

public sealed record ReportFont
{
    public string Name { get; init; } = "Arial"; // Font name

    public double Size { get; init; } = 11d; // Font size (points)

    public bool Bold { get; init; } // Bold flag

    public bool Italic { get; init; } // Italic flag

    public bool Underline { get; init; } // Underline flag

    public bool Strikeout { get; init; } // Strikethrough flag

    public string ColorHex { get; init; } = "#FF000000"; // Text color
}

public sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000"; // Background color
}

public sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new(); // Left border

    public ReportBorder Top { get; init; } = new(); // Top border

    public ReportBorder Right { get; init; } = new(); // Right border

    public ReportBorder Bottom { get; init; } = new(); // Bottom border
}

public sealed record ReportBorder
{
    public ReportBorderStyle Style { get; init; } // Border style

    public string ColorHex { get; init; } = "#FF000000"; // Border color

    public double Width { get; init; } = 0.5d; // Border width (points)
}

public sealed record ReportAlignment
{
    public ReportHorizontalAlignment Horizontal { get; init; } = ReportHorizontalAlignment.General; // Horizontal alignment

    public ReportVerticalAlignment Vertical { get; init; } = ReportVerticalAlignment.Top; // Vertical alignment
}
