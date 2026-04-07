namespace OysterReport.Generator.Models;

using OysterReport.Helpers;

internal sealed class ReportCell
{
    public ReportCell(
        int row,
        int column,
        ReportCellValue value,
        string displayText,
        ReportCellStyle style)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
        Value = value;
        DisplayText = displayText;
        Style = style;
    }

    public int Row { get; }

    public int Column { get; }

    public string Address { get; }

    public ReportCellValue Value { get; }

    public string DisplayText { get; }

    public ReportCellStyle Style { get; private set; }

    public ReportMergeInfo? Merge { get; private set; }

    internal void SetMerge(ReportMergeInfo? merge) => Merge = merge;

    internal void SetStyle(ReportCellStyle style) => Style = style;
}
