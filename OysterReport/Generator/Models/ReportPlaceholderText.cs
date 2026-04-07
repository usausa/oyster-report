namespace OysterReport.Generator.Models;

internal sealed class ReportPlaceholderText
{
    public ReportPlaceholderText(string markerText, string markerName)
    {
        MarkerText = markerText;
        MarkerName = markerName;
    }

    public string MarkerText { get; } // Raw placeholder token as it appears in Excel

    public string MarkerName { get; } // Identifier used by the application for replacement

    public string? ResolvedText { get; private set; } // Display text after substitution

    internal ReportPlaceholderText Clone() =>
        new(MarkerText, MarkerName)
        {
            ResolvedText = ResolvedText
        };

    internal void SetResolvedText(string? text) => ResolvedText = text;
}
