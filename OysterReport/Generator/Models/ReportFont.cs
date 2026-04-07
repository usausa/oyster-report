namespace OysterReport.Generator.Models;

internal sealed record ReportFont
{
    public string Name { get; init; } = "Arial"; // Font name

    public double Size { get; init; } = 11d; // Font size (points)

    public bool Bold { get; init; } // Bold flag

    public bool Italic { get; init; } // Italic flag

    public bool Underline { get; init; } // Underline flag

    public bool Strikeout { get; init; } // Strikethrough flag

    public string ColorHex { get; init; } = "#FF000000"; // Text color
}
