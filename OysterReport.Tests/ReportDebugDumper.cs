namespace OysterReport.Tests;

using System.Text;
using System.Text.Json;

using OysterReport.Internal;

internal enum ReportDumpFormat
{
    Json,
    Markdown
}

internal sealed class ReportDebugDumper
{
    private readonly JsonSerializerOptions serializerOptions = DumpPayloadFactory.SerializerOptions;

    public void DumpWorkbook(ReportRenderContext context, Stream output, ReportDumpFormat format = ReportDumpFormat.Json)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(output);

        var payload = DumpPayloadFactory.CreateWorkbookPayload(context.Workbook);
        WritePayload(output, payload, format, "Workbook");
    }

    public void DumpPdfPreparation(ReportRenderContext context, Stream output, ReportDumpFormat format = ReportDumpFormat.Json)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(output);

        var payload = DumpPayloadFactory.CreatePdfPreparationPayload(context.Workbook, context.SheetPlans);
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
