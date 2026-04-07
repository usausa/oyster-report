namespace OysterReport.Generator.Models;

using OysterReport.Helpers;

internal sealed class ReportMergedRange
{
    public ReportMergedRange(ReportRange range)
    {
        Range = range;
        OwnerCellAddress = AddressHelper.ToAddress(range.StartRow, range.StartColumn);
    }

    public ReportRange Range { get; private set; } // Merged cell range

    public string OwnerCellAddress { get; } // Owner cell address

    internal ReportMergedRange CloneShifted(int rowOffset) => new(Range.ShiftRows(rowOffset));

    internal void SetRange(ReportRange range) => Range = range;
}
