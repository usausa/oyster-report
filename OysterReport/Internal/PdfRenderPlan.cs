namespace OysterReport.Internal;

using OysterReport;

internal sealed record PdfRenderPlan
{
    public IReadOnlyList<PdfRenderSheetPlan> Sheets { get; init; } = Array.Empty<PdfRenderSheetPlan>(); // Resolved sheet render plans
}

internal sealed record PdfRenderSheetPlan
{
    public string SheetName { get; init; } = string.Empty; // Target sheet name

    public IReadOnlyList<PdfRenderPagePlan> Pages { get; init; } = Array.Empty<PdfRenderPagePlan>(); // Pages after layout resolution

    public IReadOnlyList<PdfImageRenderInfo> Images { get; init; } = Array.Empty<PdfImageRenderInfo>(); // Final image placements
}

internal sealed record PdfRenderPagePlan
{
    public int PageNumber { get; init; } // Page number (1-based)

    public ReportRect PageBounds { get; init; } // Rectangle of the entire page

    public ReportRect PrintableBounds { get; init; } // Printable area excluding margins

    public PdfHeaderFooterRenderInfo HeaderFooter { get; init; } = new(); // Header and footer render info for this page

    public IReadOnlyList<PdfCellRenderInfo> Cells { get; init; } = Array.Empty<PdfCellRenderInfo>(); // Cells to render on this page
}

internal sealed record PdfCellRenderInfo
{
    public string CellAddress { get; init; } = string.Empty; // Target cell address

    public ReportRect OuterBounds { get; init; } // Final outer bounds of the cell

    public ReportRect ContentBounds { get; init; } // Final content drawing bounds

    public ReportRect TextBounds { get; init; } // Final text overflow bounds

    public bool IsMergedOwner { get; init; } // Whether this is the owner cell of a merged range
}

internal sealed record PdfImageRenderInfo
{
    public string Name { get; init; } = string.Empty; // Image identifier name

    public ReportRect Bounds { get; init; } // Final bounding rectangle for drawing

    public ReadOnlyMemory<byte> ImageBytes { get; init; } // Image byte data
}

internal sealed record PdfHeaderFooterRenderInfo
{
    public string? HeaderText { get; init; } // Header text for this page (null if none)

    public string? FooterText { get; init; } // Footer text for this page (null if none)

    public ReportRect HeaderBounds { get; init; } // Header drawing area

    public ReportRect FooterBounds { get; init; } // Footer drawing area
}
