namespace OysterReport;

using System.Diagnostics.CodeAnalysis;

// Categorizes a non-fatal rendering problem. New kinds are added over time, so consumers
// should treat unknown values defensively rather than assuming an exhaustive switch.
public enum ReportRenderWarningKind
{
    // An embedded image could not be decoded and was skipped.
    ImageDecodeFailed
}

[ExcludeFromCodeCoverage]
public sealed record ReportRenderWarning
{
    public ReportRenderWarningKind Kind { get; init; }

    public string SheetName { get; init; } = string.Empty;

    // Identifies the element that triggered the warning (e.g. an image name); null when not applicable.
    public string? Source { get; init; }

    public string Message { get; init; } = string.Empty;

    public Exception? Exception { get; init; }
}
