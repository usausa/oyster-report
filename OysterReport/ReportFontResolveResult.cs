namespace OysterReport;

public sealed record ReportFontResolveResult
{
    public bool IsResolved { get; init; } // Whether the font was successfully resolved

    public string ResolvedFontName { get; init; } = string.Empty; // Resolved font name

    public string? Message { get; init; } // Diagnostic message (optional)
}
