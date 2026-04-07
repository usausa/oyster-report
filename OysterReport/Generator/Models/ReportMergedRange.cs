namespace OysterReport.Generator.Models;

using OysterReport.Helpers;

internal sealed class ReportMergedRange
{
    public ReportMergedRange(ReportRange range)
    {
        Range = range;
        OwnerCellAddress = AddressHelper.ToAddress(range.StartRow, range.StartColumn);
    }

    public ReportRange Range { get; }

    public string OwnerCellAddress { get; }
}
