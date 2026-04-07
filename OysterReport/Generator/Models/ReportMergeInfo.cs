namespace OysterReport.Generator.Models;

internal sealed record ReportMergeInfo
{
    public string OwnerCellAddress { get; init; } = string.Empty; // Address of the owner cell in the merged range

    public bool IsOwner { get; init; } // Whether this cell is the owner of the merge

    public ReportRange Range { get; init; } // Merged cell range
}
