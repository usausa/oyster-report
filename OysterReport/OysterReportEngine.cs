namespace OysterReport;

public sealed class OysterReportEngine
{
    private readonly ExcelReader excelReader = new();
    private readonly PdfGenerator pdfGenerator = new();
    private readonly ReportDebugDumper debugDumper = new();

    public ReportWorkbook Read(string filePath, ExcelReadOptions? options = null) =>
        excelReader.Read(filePath, options);

    public void GeneratePdf(ReportWorkbook workbook, Stream output, PdfGenerateOptions? options = null) =>
        pdfGenerator.Generate(workbook, output, options);

    public void DumpWorkbook(ReportWorkbook workbook, Stream output, ReportDumpFormat format = ReportDumpFormat.Json) =>
        debugDumper.DumpWorkbook(workbook, output, format);

    public void DumpPdfPreparation(
        ReportWorkbook workbook,
        Stream output,
        ReportDumpFormat format = ReportDumpFormat.Json) =>
        debugDumper.DumpPdfPreparation(workbook, output, format);
}
