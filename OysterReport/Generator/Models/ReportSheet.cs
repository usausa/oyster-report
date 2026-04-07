namespace OysterReport.Generator.Models;

internal sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public ReportSheet(string name)
    {
        Name = name;
        UsedRange = new ReportRange(1, 1, 1, 1);
    }

    public string Name { get; }

    public ReportRange UsedRange { get; private set; }

    public IReadOnlyList<ReportRow> Rows => rows;

    public IReadOnlyList<ReportColumn> Columns => columns;

    public IReadOnlyList<ReportCell> Cells => cells;

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges;

    public IReadOnlyList<ReportImage> Images => images;

    public ReportPageSetup PageSetup { get; private set; } = new();

    public ReportHeaderFooter HeaderFooter { get; private set; } = new();

    public ReportPrintArea? PrintArea { get; private set; }

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks;

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks;

    public bool ShowGridLines { get; private set; }

    internal void AddRowDefinition(ReportRow row) => rows.Add(row);

    internal void AddColumnDefinition(ReportColumn column) => columns.Add(column);

    internal void AddCell(ReportCell cell) => cells.Add(cell);

    internal void AddMergedRange(ReportMergedRange range) => mergedRanges.Add(range);

    internal void AddImage(ReportImage image) => images.Add(image);

    internal void AddHorizontalPageBreak(ReportPageBreak pageBreak) => horizontalPageBreaks.Add(pageBreak);

    internal void AddVerticalPageBreak(ReportPageBreak pageBreak) => verticalPageBreaks.Add(pageBreak);

    internal void SetPageSetup(ReportPageSetup pageSetup) => PageSetup = pageSetup;

    internal void SetHeaderFooter(ReportHeaderFooter headerFooter) => HeaderFooter = headerFooter;

    internal void SetPrintArea(ReportPrintArea? printArea) => PrintArea = printArea;

    internal void SetShowGridLines(bool showGridLines) => ShowGridLines = showGridLines;

    internal void SetUsedRange(ReportRange usedRange) => UsedRange = usedRange;

    internal void RecalculateLayout()
    {
        var top = 0d;
        foreach (var row in rows.OrderBy(static row => row.Index))
        {
            row.SetTop(top);
            top += row.HeightPoint;
        }

        var left = 0d;
        foreach (var column in columns.OrderBy(static column => column.Index))
        {
            column.SetLeft(left);
            left += column.WidthPoint;
        }
    }
}
