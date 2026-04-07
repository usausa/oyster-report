namespace OysterReport.Generator.Models;

internal sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty; // Template name

    public string? SourceFilePath { get; init; } // Source file path

    public DateTimeOffset? SourceLastWriteTime { get; init; } // Last write time of the source file

    public string? Author { get; init; } // Template author
}
