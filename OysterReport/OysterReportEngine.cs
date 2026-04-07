namespace OysterReport;

public sealed class OysterReportEngine
{
    private readonly ExcelReader excelReader = new();
    private readonly PdfGenerator pdfGenerator = new();
    private readonly ReportDebugDumper debugDumper = new();

    public ReportWorkbook Read(string filePath, ExcelReaderOption? option = null) =>
        excelReader.Read(filePath, option);

    public void GeneratePdf(ReportWorkbook workbook, Stream output, PdfGeneratorOption? option = null) =>
        pdfGenerator.Generate(workbook, output, option);

    public void DumpWorkbook(ReportWorkbook workbook, Stream output, ReportDumpFormat format = ReportDumpFormat.Json) =>
        debugDumper.DumpWorkbook(workbook, output, format);

    public void DumpPdfPreparation(
        ReportWorkbook workbook,
        Stream output,
        ReportDumpFormat format = ReportDumpFormat.Json) =>
        debugDumper.DumpPdfPreparation(workbook, output, format);
}
