namespace OysterReport;

using ClosedXML.Excel;

using OysterReport.Generator;

public sealed class OysterReportEngine
{
    private readonly ExcelReader excelReader = new();
    private readonly PdfGenerator pdfGenerator = new();

    /// <summary>テンプレート Excel ファイルを読み込む。</summary>
    public TemplateWorkbook Load(string filePath) =>
        new(new XLWorkbook(filePath));

    /// <summary>テンプレート Excel を Stream から読み込む。</summary>
    public TemplateWorkbook Load(Stream stream) =>
        new(new XLWorkbook(stream));

    /// <summary>テンプレートから PDF を生成する。</summary>
    public void GeneratePdf(TemplateWorkbook template, Stream output, PdfGeneratorOption? option = null)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(output);

        // ClosedXML ワークブックを保存して内部パイプラインで読み直す
        using var buffer = new MemoryStream();
        template.UnderlyingWorkbook.SaveAs(buffer);
        buffer.Position = 0;

        var reportWorkbook = excelReader.Read(buffer);
        pdfGenerator.Generate(reportWorkbook, output, option);
    }
}
