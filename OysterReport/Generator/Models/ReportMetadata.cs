namespace OysterReport.Generator.Models;

internal sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty;

    public string? SourceFilePath { get; init; }

    public DateTimeOffset? SourceLastWriteTime { get; init; }
}
