namespace OysterReport.Generator.Models;

using OysterReport.Helpers;

internal sealed class ReportCell
{
    public ReportCell(
        int row,
        int column,
        ReportCellValue value,
        string sourceText,
        string displayText,
        ReportCellStyle style,
        ReportPlaceholderText? placeholder = null)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
        Value = value;
        SourceText = sourceText;
        DisplayText = displayText;
        Placeholder = placeholder;
        Style = style;
    }

    public int Row { get; private set; } // Row number (1-based)

    public int Column { get; private set; } // Column number (1-based)

    public string Address { get; private set; } // Cell address in A1 notation

    public ReportCellValue Value { get; } // Original cell value

    public string SourceText { get; } // Original display text read from Excel

    public string DisplayText { get; private set; } // Current display text (may be modified by placeholder substitution)

    public ReportPlaceholderText? Placeholder { get; } // Placeholder information (null if not a placeholder cell)

    public ReportCellStyle Style { get; private set; } // Cell style

    public ReportRect Bounds { get; private set; } // Physical bounding rectangle of the cell

    public ReportMergeInfo? Merge { get; private set; } // Merge membership info (null if not part of a merge)

    internal void SetDisplayText(string displayText) => DisplayText = displayText;

    internal void SetBounds(ReportRect bounds) => Bounds = bounds;

    internal void SetMerge(ReportMergeInfo? merge) => Merge = merge;

    internal void SetStyle(ReportCellStyle style) => Style = style;

    internal void SetRowColumn(int row, int column)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
    }

    internal ReportCell CloneWithPosition(int row, int column)
    {
        var placeholder = Placeholder?.Clone();
        return new ReportCell(row, column, Value, SourceText, DisplayText, Style, placeholder)
        {
            Bounds = Bounds,
            Merge = Merge
        };
    }
}
