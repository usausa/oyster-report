namespace OysterReport.Generator.Models;

internal sealed record ReportHeaderFooter
{
    public bool AlignWithMargins { get; init; } = true; // Whether to align with page margins

    public bool DifferentFirst { get; init; } // Whether the first page uses a different header/footer

    public bool DifferentOddEven { get; init; } // Whether odd and even pages use different headers/footers

    public bool ScaleWithDocument { get; init; } = true; // Whether to scale with the document

    public string? OddHeader { get; init; } // Header text for odd pages

    public string? OddFooter { get; init; } // Footer text for odd pages

    public string? EvenHeader { get; init; } // Header text for even pages

    public string? EvenFooter { get; init; } // Footer text for even pages

    public string? FirstHeader { get; init; } // Header text for the first page

    public string? FirstFooter { get; init; } // Footer text for the first page
}
