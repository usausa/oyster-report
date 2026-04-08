namespace OysterReport;

using OysterReport.Generator;

public sealed class OysterReportEngine
{
    /// <summary>PDF レンダリング時のフォント解決に使用する。</summary>
    public IReportFontResolver? FontResolver { get; set; }

    /// <summary>PDF ドキュメントへメタデータを埋め込むかどうか。</summary>
    public bool EmbedDocumentMetadata { get; set; } = true;

    /// <summary>PDF のコンテンツストリームを圧縮するかどうか。</summary>
    public bool CompressContentStreams { get; set; } = true;

    /// <summary>テンプレートから PDF を生成する。</summary>
    public void GeneratePdf(TemplateWorkbook template, Stream output)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(output);

        var context = CreateRenderContext(template);
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
}
