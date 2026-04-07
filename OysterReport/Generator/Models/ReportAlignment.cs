namespace OysterReport.Generator.Models;

using ClosedXML.Excel;

internal sealed record ReportAlignment
{
    public XLAlignmentHorizontalValues Horizontal { get; init; } = XLAlignmentHorizontalValues.General;

    public XLAlignmentVerticalValues Vertical { get; init; } = XLAlignmentVerticalValues.Top;
}
