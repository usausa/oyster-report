namespace OysterReport;

using OysterReport.Helpers;

public sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];
    private readonly List<ReportDiagnostic> diagnostics = [];

    public ReportWorkbook(ReportMetadata? metadata = null, ReportMeasurementProfile? measurementProfile = null)
    {
        Metadata = metadata ?? new ReportMetadata();
        MeasurementProfile = measurementProfile ?? new ReportMeasurementProfile();
    }

    public IReadOnlyList<ReportSheet> Sheets => sheets; // List of sheets in the workbook

    public ReportMetadata Metadata { get; } // Metadata for the entire report workbook

    public ReportMeasurementProfile MeasurementProfile { get; } // Measurement settings and environment normalization profile

    public IReadOnlyList<ReportDiagnostic> Diagnostics => diagnostics; // Diagnostics collected during reading

    public ReportSheet AddSheet(string name)
    {
        var sheet = new ReportSheet(name);
        AddSheet(sheet);
        return sheet;
    }

    public void AddSheet(ReportSheet sheet)
    {
        sheets.Add(sheet);
    }

    internal void AddDiagnostic(ReportDiagnostic diagnostic)
    {
        diagnostics.Add(diagnostic);
    }
}

public sealed class ReportSheet
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

    public string Name { get; } // Sheet name

    public ReportRange UsedRange { get; private set; } // Used cell range

    public IReadOnlyList<ReportRow> Rows => rows; // Row definitions

    public IReadOnlyList<ReportColumn> Columns => columns; // Column definitions

    public IReadOnlyList<ReportCell> Cells => cells; // List of cells in the used range

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges; // List of merged cell ranges

    public IReadOnlyList<ReportImage> Images => images; // List of images on the sheet

    public ReportPageSetup PageSetup { get; private set; } = new(); // Page setup for printing

    public ReportHeaderFooter HeaderFooter { get; private set; } = new(); // Header and footer definition

    public ReportPrintArea? PrintArea { get; private set; } // Explicit print area (null if not set)

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks; // List of manual horizontal page breaks

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks; // List of manual vertical page breaks

    public bool ShowGridLines { get; private set; } // Whether to show grid lines

    public int ReplacePlaceholder(string markerName, string value)
    {
        var replaceCount = 0;
        foreach (var cell in cells.Where(static x => x.Placeholder is not null))
        {
            if (!string.Equals(cell.Placeholder!.MarkerName, markerName, StringComparison.Ordinal))
            {
                continue;
            }

            cell.SetDisplayText(value);
            cell.Placeholder.SetResolvedText(value);
            replaceCount++;
        }

        return replaceCount;
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var replaceCount = 0;
        foreach (var (key, value) in values)
        {
            replaceCount += ReplacePlaceholder(key, value ?? string.Empty);
        }

        return replaceCount;
    }

    public void AddRows(RowExpansionRequest request)
    {
        var repeatCount = request.GetRepeatCount();
        var templateRows = rows
            .Where(row => row.Index >= request.TemplateStartRowIndex && row.Index <= request.TemplateEndRowIndex)
            .OrderBy(row => row.Index)
            .ToList();

        if (templateRows.Count == 0)
        {
            throw new InvalidOperationException("Template rows were not found.");
        }

        var blockSize = request.TemplateEndRowIndex - request.TemplateStartRowIndex + 1;
        var additionalRows = repeatCount * blockSize;

        foreach (var row in rows.Where(row => row.Index > request.TemplateEndRowIndex))
        {
            row.SetIndex(row.Index + additionalRows);
        }

        foreach (var cell in cells.Where(cell => cell.Row > request.TemplateEndRowIndex))
        {
            cell.SetRowColumn(cell.Row + additionalRows, cell.Column);
        }

        foreach (var range in mergedRanges.Where(range => range.Range.StartRow > request.TemplateEndRowIndex))
        {
            range.SetRange(range.Range.ShiftRows(additionalRows));
        }

        foreach (var image in images.Where(image => image.FromRow > request.TemplateEndRowIndex))
        {
            image.ShiftRows(additionalRows);
        }

        var insertIndex = rows.FindLastIndex(row => row.Index <= request.TemplateEndRowIndex) + 1;
        var templateCells = cells
            .Where(cell => cell.Row >= request.TemplateStartRowIndex && cell.Row <= request.TemplateEndRowIndex)
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .ToList();

        for (var iteration = 0; iteration < repeatCount; iteration++)
        {
            var rowOffset = blockSize * (iteration + 1);
            foreach (var templateRow in templateRows)
            {
                rows.Insert(insertIndex++, templateRow.CloneWithIndex(templateRow.Index + rowOffset));
            }

            foreach (var templateCell in templateCells)
            {
                var clone = templateCell.CloneWithPosition(templateCell.Row + rowOffset, templateCell.Column);
                var placeholderValues = request.GetPlaceholderValues(iteration);
                if (clone.Placeholder is not null &&
                    placeholderValues.TryGetValue(clone.Placeholder.MarkerName, out var replacement))
                {
                    var resolvedText = replacement ?? string.Empty;
                    clone.SetDisplayText(resolvedText);
                    clone.Placeholder.SetResolvedText(resolvedText);
                }

                cells.Add(clone);
            }

            foreach (var templateRange in mergedRanges.Where(range => range.Range.StartRow >= request.TemplateStartRowIndex && range.Range.EndRow <= request.TemplateEndRowIndex).ToList())
            {
                mergedRanges.Add(templateRange.CloneShifted(rowOffset));
            }

            foreach (var templateImage in images.Where(image => image.FromRow >= request.TemplateStartRowIndex && image.FromRow <= request.TemplateEndRowIndex).ToList())
            {
                images.Add(templateImage.CloneShifted(rowOffset));
            }
        }

        rows.Sort(static (left, right) => left.Index.CompareTo(right.Index));
        cells.Sort(static (left, right) =>
        {
            var rowCompare = left.Row.CompareTo(right.Row);
            return rowCompare != 0 ? rowCompare : left.Column.CompareTo(right.Column);
        });
        mergedRanges.Sort(static (left, right) => left.Range.StartRow.CompareTo(right.Range.StartRow));

        UpdateUsedRange();
        RecalculateLayout();
    }

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

        foreach (var cell in cells)
        {
            var row = rows.FirstOrDefault(item => item.Index == cell.Row);
            var column = columns.FirstOrDefault(item => item.Index == cell.Column);
            if (row is null || column is null)
            {
                continue;
            }

            cell.SetBounds(new ReportRect
            {
                X = column.LeftPoint,
                Y = row.TopPoint,
                Width = column.WidthPoint,
                Height = row.HeightPoint
            });
        }
    }

    private void UpdateUsedRange()
    {
        if (cells.Count == 0)
        {
            UsedRange = new ReportRange(1, 1, 1, 1);
            return;
        }

        UsedRange = new ReportRange(
            cells.Min(static cell => cell.Row),
            cells.Min(static cell => cell.Column),
            cells.Max(static cell => cell.Row),
            cells.Max(static cell => cell.Column));
    }
}

public sealed class ReportRow
{
    public ReportRow(int index, double heightPoint, bool isHidden = false, int outlineLevel = 0)
    {
        Index = index;
        HeightPoint = heightPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
    }

    public int Index { get; private set; } // Row number (1-based)

    public double HeightPoint { get; } // Row height (points)

    public double TopPoint { get; private set; } // Top position from the sheet origin (points)

    public bool IsHidden { get; } // Whether the row is hidden

    public int OutlineLevel { get; } // Outline level

    internal ReportRow CloneWithIndex(int index) => new(index, HeightPoint, IsHidden, OutlineLevel);

    internal void SetIndex(int index) => Index = index;

    internal void SetTop(double topPoint) => TopPoint = topPoint;
}

public sealed class ReportColumn
{
    public ReportColumn(int index, double widthPoint, bool isHidden = false, int outlineLevel = 0, double originalExcelWidth = 0)
    {
        Index = index;
        WidthPoint = widthPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
        OriginalExcelWidth = originalExcelWidth;
    }

    public int Index { get; } // Column number (1-based)

    public double WidthPoint { get; } // Column width (points)

    public double LeftPoint { get; private set; } // Left position from the sheet origin (points)

    public bool IsHidden { get; } // Whether the column is hidden

    public int OutlineLevel { get; } // Outline level

    public double OriginalExcelWidth { get; } // Original column width value from Excel

    internal void SetLeft(double leftPoint) => LeftPoint = leftPoint;
}

public sealed class ReportCell
{
    public ReportCell(
        int row,
        int column,
        ReportCellValue value,
        string sourceText,
        string displayText,
        ReportCellStyle style,
        ReportPlaceholderText? placeholder = null)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
        Value = value;
        SourceText = sourceText;
        DisplayText = displayText;
        Placeholder = placeholder;
        Style = style;
    }

    public int Row { get; private set; } // Row number (1-based)

    public int Column { get; private set; } // Column number (1-based)

    public string Address { get; private set; } // Cell address in A1 notation

    public ReportCellValue Value { get; } // Original cell value

    public string SourceText { get; } // Original display text read from Excel

    public string DisplayText { get; private set; } // Current display text (may be modified by placeholder substitution)

    public ReportPlaceholderText? Placeholder { get; } // Placeholder information (null if not a placeholder cell)

    public ReportCellStyle Style { get; private set; } // Cell style

    public ReportRect Bounds { get; private set; } // Physical bounding rectangle of the cell

    public ReportMergeInfo? Merge { get; private set; } // Merge membership info (null if not part of a merge)

    internal void SetDisplayText(string displayText) => DisplayText = displayText;

    internal void SetBounds(ReportRect bounds) => Bounds = bounds;

    internal void SetMerge(ReportMergeInfo? merge) => Merge = merge;

    internal void SetStyle(ReportCellStyle style) => Style = style;

    internal void SetRowColumn(int row, int column)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
    }

    internal ReportCell CloneWithPosition(int row, int column)
    {
        var placeholder = Placeholder?.Clone();
        return new ReportCell(row, column, Value, SourceText, DisplayText, Style, placeholder)
        {
            Bounds = Bounds,
            Merge = Merge
        };
    }
}

public sealed class ReportPlaceholderText
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

public sealed class ReportMergedRange
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

public sealed class ReportImage
{
    public ReportImage(
        string name,
        ReportAnchorType anchorType,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        AnchorType = anchorType;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;

        var (row, _) = AddressHelper.ParseAddress(fromCellAddress);
        FromRow = row;
    }

    public string Name { get; } // Image identifier name

    public ReportAnchorType AnchorType { get; } // Anchor type

    public string FromCellAddress { get; private set; } // Starting cell address

    public string? ToCellAddress { get; private set; } // Ending cell address

    public ReportOffset Offset { get; } // Offset within the starting cell

    public double WidthPoint { get; } // Image width (points)

    public double HeightPoint { get; } // Image height (points)

    public ReadOnlyMemory<byte> ImageBytes { get; } // Raw image data

    internal int FromRow { get; private set; } // Starting row number (1-based)

    internal ReportImage CloneShifted(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        var shiftedFrom = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        string? shiftedTo = null;
        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            shiftedTo = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }

        return new ReportImage(Name, AnchorType, shiftedFrom, shiftedTo, Offset, WidthPoint, HeightPoint, ImageBytes);
    }

    internal void ShiftRows(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        FromCellAddress = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        FromRow += rowOffset;

        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            ToCellAddress = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }
    }
}

public sealed record RowExpansionRequest
{
    public int TemplateStartRowIndex { get; init; } // Start row index of the template block

    public int TemplateEndRowIndex { get; init; } // End row index of the template block

    public int RepeatCount { get; init; } // Number of additional repetitions to insert

    public IReadOnlyList<IReadOnlyDictionary<string, string?>> PlaceholderValuesByIteration { get; init; } =
        Array.Empty<IReadOnlyDictionary<string, string?>>(); // Placeholder values applied to each repeated row

    internal int GetRepeatCount()
    {
        if (RepeatCount > 0)
        {
            return RepeatCount;
        }

        if (PlaceholderValuesByIteration.Count > 0)
        {
            return PlaceholderValuesByIteration.Count;
        }

        throw new InvalidOperationException("RepeatCount or PlaceholderValuesByIteration must be specified.");
    }

    internal IReadOnlyDictionary<string, string?> GetPlaceholderValues(int iteration) =>
        iteration < PlaceholderValuesByIteration.Count
            ? PlaceholderValuesByIteration[iteration]
            : new Dictionary<string, string?>();
}
