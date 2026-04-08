namespace OysterReport.Generator;

internal sealed record ReportRenderContext
{
    public ReportWorkbook Workbook { get; init; } = new();

    public IReadOnlyList<PdfRenderSheetPlan> SheetPlans { get; init; } = [];

    public IReportFontResolver? FontResolver { get; init; }

    public bool EmbedDocumentMetadata { get; init; } = true;

    public bool CompressContentStreams { get; init; } = true;
}
