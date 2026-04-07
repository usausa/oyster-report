namespace OysterReport.Generator.Models;

using ClosedXML.Excel;

internal sealed record ReportCellValue
{
    public XLDataType Kind { get; init; } = XLDataType.Blank;

    public object? RawValue { get; init; }
}
