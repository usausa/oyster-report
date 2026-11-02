namespace OysterReport;

using System.IO.Compression;

using OysterReport.Internal;

public sealed class OysterReportEngine
{
    public IReportFontResolver? FontResolver { get; set; }

    public ReportRenderOption RenderingOptions { get; set; } = new();

    public bool EmbedDocumentMetadata { get; set; } = true;

    public bool CompressContentStreams { get; set; } = true;

    public CompressionLevel CompressionLevel
    {
        get;
        set => field = Enum.IsDefined(value) ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown compression level.");
    }

    public void GeneratePdf(TemplateWorkbook workbook, Stream output)
    {
        var context = CreateRenderContext(workbook);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(TemplateWorkbook template)
    {
        var workbook = template.ReportWorkbook;
        return CreateRenderContext(workbook, PdfRenderPlanner.BuildPlan(workbook, RenderingOptions));
    }

    public void GeneratePdf(IEnumerable<TemplateWorkbook> workbooks, Stream output)
    {
        var context = CreateRenderContext(workbooks);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(IEnumerable<TemplateWorkbook> templates)
    {
        var documents = templates.Select(static x => x.ReportWorkbook).ToList();
        var first = documents.Count > 0 ? documents[0] : new ReportWorkbook();
        var combinedWorkbook = new ReportWorkbook
        {
            Metadata = first.Metadata,
            MeasurementProfile = first.MeasurementProfile
        };
        foreach (var sheet in documents.SelectMany(static x => x.Sheets))
        {
            combinedWorkbook.AddSheet(sheet);
        }

        return CreateRenderContext(combinedWorkbook, PdfRenderPlanner.BuildPlan(documents, RenderingOptions));
    }

    public void GeneratePdf(TemplateSheet sheet, Stream output)
    {
        var context = CreateRenderContext(sheet);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(TemplateSheet template)
    {
        var singleSheetWorkbook = new ReportWorkbook
        {
            Metadata = template.WorkbookMetadata with { TemplateName = template.Name },
            MeasurementProfile = template.WorkbookMeasurementProfile
        };
        singleSheetWorkbook.AddSheet(template.UnderlyingSheet);

        return CreateRenderContext(singleSheetWorkbook, PdfRenderPlanner.BuildPlan(singleSheetWorkbook, RenderingOptions));
    }

    private ReportRenderContext CreateRenderContext(ReportWorkbook workbook, IReadOnlyList<PdfRenderSheetPlan> sheetPlans) =>
        new()
        {
            Workbook = workbook,
            SheetPlans = sheetPlans,
            FontResolver = FontResolver,
            RenderingOptions = RenderingOptions,
            EmbedDocumentMetadata = EmbedDocumentMetadata,
            CompressContentStreams = CompressContentStreams,
            CompressionLevel = CompressionLevel
        };
}
