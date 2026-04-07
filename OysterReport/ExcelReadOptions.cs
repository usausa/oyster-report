namespace OysterReport;

public sealed class ExcelReadOptions
{
    public IReadOnlyList<string>? TargetSheets { get; set; } // Sheet names to include (null or empty means all sheets)

    public bool IncludeImages { get; set; } = true; // Whether to include images
}
