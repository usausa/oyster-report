namespace OysterReport;

public sealed class PdfGeneratorOption
{
    public IReportFontResolver? FontResolver { get; set; } // Font resolver used during PDF rendering

    public bool EmbedDocumentMetadata { get; set; } = true; // Whether to embed document metadata into the PDF

    public bool CompressContentStreams { get; set; } = true; // Whether to compress PDF content streams
}

