namespace OysterReport.Generator.Models;

internal sealed record ReportCellValue
{
    public ReportCellValueKind Kind { get; init; } // Original data type

    public object? RawValue { get; init; } // Raw value retrieved from Excel
}
