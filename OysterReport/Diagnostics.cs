namespace OysterReport;

using System.Text;
using System.Text.Json;

using OysterReport.Internal;

public sealed class ReportDiagnostic
{
    public ReportDiagnosticSeverity Severity { get; init; } // 重大度

    public string Code { get; init; } = string.Empty; // 診断コード

    public string Message { get; init; } = string.Empty; // 利用者向け診断メッセージ

    public string? SheetName { get; init; } // 関連シート名

    public string? CellAddress { get; init; } // 関連セル番地
}

public sealed class ReportDebugDumper
{
    private readonly JsonSerializerOptions serializerOptions = DumpPayloadFactory.SerializerOptions;

    public void DumpWorkbook(ReportWorkbook workbook, Stream output, ReportDumpFormat format = ReportDumpFormat.Json)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(output);

        var payload = DumpPayloadFactory.CreateWorkbookPayload(workbook);
        WritePayload(output, payload, format, "Workbook");
    }

    public void DumpPdfPreparation(
        ReportWorkbook workbook,
        Stream output,
        PdfGenerateOptions? options = null,
        ReportDumpFormat format = ReportDumpFormat.Json)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(output);

        var renderPlan = PdfGenerator.BuildRenderPlan(workbook, options ?? new PdfGenerateOptions());
        var payload = DumpPayloadFactory.CreatePdfPreparationPayload(workbook, renderPlan);
        WritePayload(output, payload, format, "PdfPreparation");
    }

    private void WritePayload(Stream output, object payload, ReportDumpFormat format, string title)
    {
        var text = format switch
        {
            ReportDumpFormat.Markdown => BuildMarkdown(payload, title),
            _ => JsonSerializer.Serialize(payload, serializerOptions)
        };

        using var writer = new StreamWriter(output, Encoding.UTF8, 1024, leaveOpen: true);
        writer.Write(text);
        writer.Flush();
    }

    private string BuildMarkdown(object payload, string title)
    {
        var json = JsonSerializer.Serialize(payload, serializerOptions);
        var builder = new StringBuilder();
        builder.AppendLine(FormattableString.Invariant($"# {title}"));
        builder.AppendLine();
        builder.AppendLine("```json");
        builder.AppendLine(json);
        builder.AppendLine("```");
        return builder.ToString();
    }
}
