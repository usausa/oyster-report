namespace OysterReport.Generator.Models;

internal sealed record PdfCellRenderInfo
{
    public string CellAddress { get; init; } = string.Empty; // Target cell address

    public ReportRect OuterBounds { get; init; } // Final outer bounds of the cell

    public ReportRect ContentBounds { get; init; } // Final content drawing bounds

    public ReportRect TextBounds { get; init; } // Final text overflow bounds

    public bool IsMergedOwner { get; init; } // Whether this is the owner cell of a merged range
}
