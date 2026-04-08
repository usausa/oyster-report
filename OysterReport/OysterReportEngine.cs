namespace OysterReport;

using OysterReport.Internal;

public sealed class OysterReportEngine
{
    /// <summary>PDF レンダリング時のフォント解決に使用する。</summary>
    public IReportFontResolver? FontResolver { get; set; }

    /// <summary>PDF ドキュメントへメタデータを埋め込むかどうか。</summary>
    public bool EmbedDocumentMetadata { get; set; } = true;

    /// <summary>PDF のコンテンツストリームを圧縮するかどうか。</summary>
    public bool CompressContentStreams { get; set; } = true;

    /// <summary>ワークブック全体から PDF を生成する。</summary>
    public void GeneratePdf(TemplateWorkbook template, Stream output)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(output);

        var context = CreateRenderContext(template);
        PdfGenerator.WritePdf(context, output);
    }

    /// <summary>単一シートから PDF を生成する。</summary>
    public void GeneratePdf(TemplateSheet sheet, Stream output)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(output);

        var context = CreateRenderContext(sheet);
        PdfGenerator.WritePdf(context, output);
    }

    internal ReportRenderContext CreateRenderContext(TemplateWorkbook template)
    {
        ArgumentNullException.ThrowIfNull(template);

        var workbook = ExcelReader.Read(template.UnderlyingWorkbook);
        var sheetPlans = PdfRenderPlanner.BuildPlan(workbook);

        return new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = sheetPlans,
            FontResolver = FontResolver,
            EmbedDocumentMetadata = EmbedDocumentMetadata,
            CompressContentStreams = CompressContentStreams
        };
    }

    internal ReportRenderContext CreateRenderContext(TemplateSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        var workbook = ExcelReader.Read(sheet.UnderlyingWorksheet);
        var sheetPlans = PdfRenderPlanner.BuildPlan(workbook);

        return new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = sheetPlans,
            FontResolver = FontResolver,
            EmbedDocumentMetadata = EmbedDocumentMetadata,
            CompressContentStreams = CompressContentStreams
        };
    }
}
