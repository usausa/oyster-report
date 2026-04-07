namespace OysterReport.Generator.Models;

internal sealed record PdfImageRenderInfo
{
    public string Name { get; init; } = string.Empty; // Image identifier name

    public ReportRect Bounds { get; init; } // Final bounding rectangle for drawing

    public ReadOnlyMemory<byte> ImageBytes { get; init; } // Image byte data
}
