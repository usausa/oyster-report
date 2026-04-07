namespace OysterReport.Generator.Models;

using ClosedXML.Excel;

internal sealed record ReportBorder
{
    public XLBorderStyleValues Style { get; init; } = XLBorderStyleValues.None;

    public string ColorHex { get; init; } = "#FF000000";

    public double Width { get; init; } = 0.5d;
}
