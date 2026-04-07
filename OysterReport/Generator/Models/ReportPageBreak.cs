namespace OysterReport.Generator.Models;

internal sealed record ReportPageBreak
{
    public int Index { get; init; } // Row or column index of the page break

    public bool IsHorizontal { get; init; } // Whether this is a horizontal page break
}
