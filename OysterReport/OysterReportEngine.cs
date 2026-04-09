namespace OysterReport;

using OysterReport.Internal;

public sealed class OysterReportEngine
{
    public IReportFontResolver? FontResolver { get; set; }

    public ReportRenderOption RenderingOptions { get; set; } = new();

    public bool EmbedDocumentMetadata { get; set; } = true;

    public bool CompressContentStreams { get; set; } = true;

    public void GeneratePdf(TemplateWorkbook workbook, Stream output)
    {
        var context = CreateRenderContext(workbook);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(TemplateWorkbook template)
    {
        var workbook = ExcelReader.Read(template.UnderlyingWorkbook, RenderingOptions);
        var sheetPlans = PdfRenderPlanner.BuildPlan(workbook, RenderingOptions);

        return new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = sheetPlans,
            FontResolver = FontResolver,
            RenderingOptions = RenderingOptions,
            EmbedDocumentMetadata = EmbedDocumentMetadata,
            CompressContentStreams = CompressContentStreams
        };
    }

    public void GeneratePdf(TemplateSheet sheet, Stream output)
    {
        var context = CreateRenderContext(sheet);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(TemplateSheet template)
    {
        var workbook = ExcelReader.Read(template.UnderlyingWorksheet, RenderingOptions);
        var sheetPlans = PdfRenderPlanner.BuildPlan(workbook, RenderingOptions);

        return new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = sheetPlans,
            FontResolver = FontResolver,
            RenderingOptions = RenderingOptions,
            EmbedDocumentMetadata = EmbedDocumentMetadata,
            CompressContentStreams = CompressContentStreams
        };
    }
}
