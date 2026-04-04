namespace OysterReport.Diagnostics;

using System.Globalization;
using System.Text;
using System.Text.Json;
using OysterReport.Common;
using OysterReport.Model;
using OysterReport.Writing.Pdf;

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
            _ => JsonSerializer.Serialize(payload, serializerOptions),
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

internal static class DumpPayloadFactory
{
    public static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
    };

    public static object CreateWorkbookPayload(ReportWorkbook workbook) =>
        new
        {
            workbook.Metadata,
            workbook.MeasurementProfile,
            workbook.Diagnostics,
            Sheets = workbook.Sheets.Select(sheet => new
            {
                sheet.Name,
                UsedRange = sheet.UsedRange.ToString(),
                sheet.ShowGridLines,
                Rows = sheet.Rows.Select(row => new
                {
                    row.Index,
                    row.HeightPoint,
                    row.TopPoint,
                    row.IsHidden,
                    row.OutlineLevel,
                }),
                Columns = sheet.Columns.Select(column => new
                {
                    column.Index,
                    column.WidthPoint,
                    column.LeftPoint,
                    column.IsHidden,
                    column.OutlineLevel,
                    column.OriginalExcelWidth,
                }),
                Cells = sheet.Cells.Select(cell => new
                {
                    cell.Row,
                    cell.Column,
                    cell.Address,
                    cell.DisplayText,
                    Placeholder = cell.Placeholder?.MarkerName,
                }),
                sheet.MergedRanges,
                sheet.Images,
                sheet.PageSetup,
                sheet.HeaderFooter,
                sheet.PrintArea,
                sheet.HorizontalPageBreaks,
                sheet.VerticalPageBreaks,
            }),
        };

    public static object CreatePdfPreparationPayload(ReportWorkbook workbook, object renderPlan) =>
        new
        {
            Workbook = CreateWorkbookPayload(workbook),
            RenderPlan = renderPlan,
            Environment = new
            {
                OperatingSystem = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
                Architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
                Culture = CultureInfo.CurrentCulture.Name,
                Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            },
        };
}
