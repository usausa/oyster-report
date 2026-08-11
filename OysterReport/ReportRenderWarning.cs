namespace OysterReport;

using System.Diagnostics.CodeAnalysis;

public enum ReportRenderWarningKind
{
    ImageDecodeFailed,
    FontResolverNotInstalled
}

[ExcludeFromCodeCoverage]
public sealed record ReportRenderWarning
{
    public ReportRenderWarningKind Kind { get; init; }

    public string SheetName { get; init; } = string.Empty;

    public string? Source { get; init; }

    public string Message { get; init; } = string.Empty;

    public Exception? Exception { get; init; }
}
