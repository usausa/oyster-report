namespace OysterReport;

public sealed record ReportFontRequest
{
    public string FontName { get; init; } = string.Empty; // Requested font name

    public bool Bold { get; init; } // Bold request flag

    public bool Italic { get; init; } // Italic request flag
}
