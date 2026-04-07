namespace OysterReport;

public sealed class PdfGeneratorOption
{
    public IReportFontResolver? FontResolver { get; set; } // Font resolver used during PDF rendering

    public bool EmbedDocumentMetadata { get; set; } = true; // Whether to embed document metadata into the PDF

    public bool CompressContentStreams { get; set; } = true; // Whether to compress PDF content streams
}

public interface IReportFontResolver
{
    ReportFontResolveResult Resolve(ReportFontRequest request);
}

public sealed record ReportFontRequest
{
    public string FontName { get; init; } = string.Empty; // Requested font name

    public bool Bold { get; init; } // Bold request flag

    public bool Italic { get; init; } // Italic request flag
}

public sealed record ReportFontResolveResult
{
    public bool IsResolved { get; init; } // Whether the font was successfully resolved

    public string ResolvedFontName { get; init; } = string.Empty; // Resolved font name

    public string? Message { get; init; } // Diagnostic message (optional)
}
