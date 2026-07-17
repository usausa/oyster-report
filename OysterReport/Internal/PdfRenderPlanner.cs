namespace OysterReport.Internal;

//--------------------------------------------------------------------------------
// Pipeline plan
//--------------------------------------------------------------------------------

internal sealed record PdfImageRenderInfo
{
    public string Name { get; init; } = string.Empty;

    public ReportRect Bounds { get; init; }

    public ReadOnlyMemory<byte> ImageBytes { get; init; }
}

internal sealed record PdfHeaderFooterRenderInfo
{
    public string? HeaderText { get; init; }

    public string? FooterText { get; init; }

    public ReportRect HeaderBounds { get; init; }

    public ReportRect FooterBounds { get; init; }
}

internal sealed record PdfCellRenderInfo
{
    public string CellAddress { get; init; } = string.Empty;

    public string BackgroundColorHex { get; init; } = string.Empty;

    public ReportRect OuterBounds { get; init; }

    public ReportRect ContentBounds { get; init; }

    public ReportRect TextBounds { get; init; }

    public bool IsMergedOwner { get; init; }
}

internal sealed record PdfRenderPagePlan
{
    public int PageNumber { get; init; }

    public ReportRect PageBounds { get; init; }

    public ReportRect PrintableBounds { get; init; }

    public PdfHeaderFooterRenderInfo HeaderFooter { get; init; } = new();

    public IReadOnlyList<PdfCellRenderInfo> Cells { get; init; } = [];
}

internal sealed record PdfRenderSheetPlan
{
    public string SheetName { get; init; } = string.Empty;

    public IReadOnlyList<PdfRenderPagePlan> Pages { get; init; } = [];

    public IReadOnlyList<PdfImageRenderInfo> Images { get; init; } = [];
}

//--------------------------------------------------------------------------------
// Planner
//--------------------------------------------------------------------------------

internal static class PdfRenderPlanner
{
    //--------------------------------------------------------------------------------
    // Build
    //--------------------------------------------------------------------------------

    public static IReadOnlyList<PdfRenderSheetPlan> BuildPlan(ReportWorkbook workbook, ReportRenderOption? renderOption = null)
    {
        var effectiveOptions = renderOption ?? new ReportRenderOption();
        return workbook.Sheets.Select((x, i) => BuildSheetPlan(x, i + 1, effectiveOptions)).ToList();
    }

    private static PdfRenderSheetPlan BuildSheetPlan(ReportSheet sheet, int sheetNumber, ReportRenderOption renderOption)
    {
        // Resolve page bounding rectangle from paper size and orientation
        var pageBounds = ResolvePageBounds(sheet.PageSetup, renderOption);

        // Compute the printable area by subtracting margins from the page bounds
        var printableBounds = new ReportRect
        {
            X = sheet.PageSetup.Margins.Left,
            Y = sheet.PageSetup.Margins.Top,
            Width = pageBounds.Width - sheet.PageSetup.Margins.Left - sheet.PageSetup.Margins.Right,
            Height = pageBounds.Height - sheet.PageSetup.Margins.Top - sheet.PageSetup.Margins.Bottom
        };

        // Resolve the render range from print area or used cell range
        var renderRange = sheet.PrintArea?.Range ?? sheet.UsedRange;

        // Filter rows and columns within the render range, sorted by index
        var defaultRowHeight = sheet.DefaultRowHeightPoint;
        var rowsInRange = sheet.Rows
            .Where(x => (x.Index >= renderRange.StartRow) && (x.Index <= renderRange.EndRow))
            .OrderBy(static x => x.Index)
            .ToList();
        var visibleRows = rowsInRange.Where(static x => !x.IsHidden).ToList();
        var visibleColumns = sheet.Columns
            .Where(x => !x.IsHidden && (x.Index >= renderRange.StartColumn) && (x.Index <= renderRange.EndColumn))
            .OrderBy(static x => x.Index)
            .ToList();

        // Calculate content offsets for horizontal and vertical centering
        var nonMaterializedRowCount = (renderRange.EndRow - renderRange.StartRow + 1) - rowsInRange.Count;
        var contentWidth = visibleColumns.Sum(static x => x.WidthPoint);
        var contentHeight = visibleRows.Sum(static x => x.HeightPoint) + (nonMaterializedRowCount * defaultRowHeight);
        var contentOffsetX = sheet.PageSetup.CenterHorizontally
            ? Math.Max(0d, (printableBounds.Width - contentWidth) / 2d)
            : 0d;
        var contentOffsetY = sheet.PageSetup.CenterVertically
            ? (printableBounds.Height - contentHeight) / 2d
            : 0d;

        // Build the row-index to Y-offset (pt) dictionary, filling the gaps between materialized rows
        var rowOffsets = new Dictionary<int, double>();
        var currentTop = printableBounds.Y + contentOffsetY;
        var previousRowIndex = renderRange.StartRow - 1;
        foreach (var row in rowsInRange)
        {
            currentTop += (row.Index - previousRowIndex - 1) * defaultRowHeight;
            if (!row.IsHidden)
            {
                rowOffsets[row.Index] = currentTop;
                currentTop += row.HeightPoint;
            }

            previousRowIndex = row.Index;
        }

        // Build the column-index to X-offset (pt) dictionary
        var columnOffsets = new Dictionary<int, double>();
        var currentLeft = printableBounds.X + contentOffsetX;
        foreach (var column in visibleColumns)
        {
            columnOffsets[column.Index] = currentLeft;
            currentLeft += column.WidthPoint;
        }

        // Images can anchor on a blank (non-materialized) row that carries no cell
        AddImageAnchorRowOffsets(sheet, renderRange, rowsInRange, rowOffsets, printableBounds.Y + contentOffsetY, defaultRowHeight);

        // Build lookup dictionaries for merged ranges and columns; cells are resolved via sheet.FindCell
        var mergedRanges = sheet.MergedRanges.ToDictionary(range => range.OwnerCellAddress, range => range);
        var columnByIndex = visibleColumns.ToDictionary(c => c.Index);
        var rowByIndex = visibleRows.ToDictionary(r => r.Index);
        var mergedRangesByRow = BuildMergedRangesByRow(sheet.MergedRanges);

        // Compute bounding rectangles for each visible cell and build the PdfCellRenderInfo list
        var pageCells = new List<PdfCellRenderInfo>();
        foreach (var cell in sheet.Cells.Where(x => rowByIndex.ContainsKey(x.Row) && columnOffsets.ContainsKey(x.Column)))
        {
            var outerBounds = new ReportRect
            {
                X = columnOffsets[cell.Column],
                Y = rowOffsets[cell.Row],
                Width = columnByIndex[cell.Column].WidthPoint,
                Height = rowByIndex[cell.Row].HeightPoint
            };

            var isMergedOwner = mergedRanges.TryGetValue(cell.Address, out var mergedRange);
            if (isMergedOwner && mergedRange is not null)
            {
                outerBounds = BuildMergedBounds(mergedRange, visibleRows, visibleColumns, rowOffsets, columnOffsets);
            }

            var contentBounds = outerBounds.Deflate(new ReportThickness
            {
                Left = renderOption.HorizontalCellTextPadding,
                Right = renderOption.HorizontalCellTextPadding
            });
            pageCells.Add(new PdfCellRenderInfo
            {
                CellAddress = cell.Address,
                BackgroundColorHex = ResolveBackgroundColor(cell, sheet.Tables),
                OuterBounds = outerBounds,
                ContentBounds = contentBounds,
                TextBounds = ComputeTextOverflowBounds(
                    cell,
                    contentBounds,
                    outerBounds,
                    isMergedOwner,
                    isMergedOwner ? mergedRanges.GetValueOrDefault(cell.Address) : null,
                    sheet,
                    columnByIndex,
                    columnOffsets,
                    mergedRangesByRow),
                IsMergedOwner = isMergedOwner
            });
        }

        return new PdfRenderSheetPlan
        {
            SheetName = sheet.Name,
            Pages =
            [
                new PdfRenderPagePlan
                {
                    PageNumber = sheetNumber,
                    PageBounds = pageBounds,
                    PrintableBounds = printableBounds,
                    HeaderFooter = BuildHeaderFooter(sheet, pageBounds, printableBounds, sheetNumber),
                    Cells = pageCells
                }
            ],
            Images = BuildImageInfos(sheet, rowOffsets, columnOffsets)
        };
    }

    //--------------------------------------------------------------------------------
    // Page bounds
    //--------------------------------------------------------------------------------

    private static ReportRect ResolvePageBounds(ReportPageSetup pageSetup, ReportRenderOption renderOption)
    {
        // Determines the page bounding rectangle (pt) from paper size and orientation.
        var (width, height) = renderOption.PageSizeResolver(pageSetup.PaperSize);
        return pageSetup.Orientation == PageOrientation.Landscape
            ? new ReportRect { X = 0, Y = 0, Width = height, Height = width }
            : new ReportRect { X = 0, Y = 0, Width = width, Height = height };
    }

    //--------------------------------------------------------------------------------
    // Merged cell
    //--------------------------------------------------------------------------------

    private static ReportRect BuildMergedBounds(
        ReportMergedRange mergedRange,
        IEnumerable<ReportRow> visibleRows,
        IEnumerable<ReportColumn> visibleColumns,
        Dictionary<int, double> rowOffsets,
        Dictionary<int, double> columnOffsets)
    {
        // Computes the outer bounding rectangle of a merged cell from row and column offsets
        var targetRows = visibleRows.Where(x => (x.Index >= mergedRange.Range.StartRow) && (x.Index <= mergedRange.Range.EndRow)).ToList();
        var targetColumns = visibleColumns.Where(x => (x.Index >= mergedRange.Range.StartColumn) && (x.Index <= mergedRange.Range.EndColumn)).ToList();
        if ((targetRows.Count == 0) || (targetColumns.Count == 0))
        {
            return default;
        }

        return new ReportRect
        {
            X = columnOffsets[targetColumns[0].Index],
            Y = rowOffsets[targetRows[0].Index],
            Width = targetColumns.Sum(static x => x.WidthPoint),
            Height = targetRows.Sum(static x => x.HeightPoint)
        };
    }

    private static Dictionary<int, List<ReportMergedRange>> BuildMergedRangesByRow(
        IEnumerable<ReportMergedRange> mergedRanges)
    {
        // Indexes merged ranges by row so membership checks avoid expanding every covered cell
        var map = new Dictionary<int, List<ReportMergedRange>>();
        foreach (var mr in mergedRanges)
        {
            for (var r = mr.Range.StartRow; r <= mr.Range.EndRow; r++)
            {
                if (!map.TryGetValue(r, out var list))
                {
                    list = [];
                    map[r] = list;
                }

                list.Add(mr);
            }
        }

        return map;
    }

    private static bool IsInsideMergedRange(Dictionary<int, List<ReportMergedRange>> mergedRangesByRow, int row, int column)
    {
        if (!mergedRangesByRow.TryGetValue(row, out var list))
        {
            return false;
        }

        foreach (var mr in list)
        {
            if ((column >= mr.Range.StartColumn) && (column <= mr.Range.EndColumn))
            {
                return true;
            }
        }

        return false;
    }

    //--------------------------------------------------------------------------------
    // Table stripe
    //--------------------------------------------------------------------------------

    private static string ResolveBackgroundColor(ReportCell cell, IReadOnlyList<ReportTable> tables)
    {
        var hex = cell.Style.Fill.BackgroundColorHex;
        if (!hex.StartsWith("#00", StringComparison.Ordinal))
        {
            return hex;
        }

        for (var i = tables.Count - 1; i >= 0; i--)
        {
            var table = tables[i];
            if (!table.ShowRowStripes || String.IsNullOrEmpty(table.StripeColorHex))
            {
                continue;
            }

            var firstDataRow = table.Range.StartRow + (table.ShowHeader ? 1 : 0);
            var lastDataRow = table.Range.EndRow - (table.ShowTotals ? 1 : 0);
            if ((cell.Row < firstDataRow) || (cell.Row > lastDataRow) ||
                (((cell.Row - firstDataRow) % 2) != 0) ||
                (cell.Column < table.Range.StartColumn) || (cell.Column > table.Range.EndColumn))
            {
                continue;
            }

            return table.StripeColorHex;
        }

        return hex;
    }

    //--------------------------------------------------------------------------------
    // Text layout
    //--------------------------------------------------------------------------------

    private static ReportRect ComputeTextOverflowBounds(
        ReportCell cell,
        ReportRect contentBounds,
        ReportRect outerBounds,
        bool isMergedOwner,
        ReportMergedRange? mergedRange,
        ReportSheet sheet,
        Dictionary<int, ReportColumn> columnByIndex,
        Dictionary<int, double> columnOffsets,
        Dictionary<int, List<ReportMergedRange>> mergedRangesByRow)
    {
        // Computes text drawing bounds
        if (cell.Style.WrapText || cell.DisplayText.Contains('\n', StringComparison.Ordinal))
        {
            return contentBounds;
        }

        if ((cell.Style.Alignment.Horizontal != HorizontalAlignment.General) &&
            (cell.Style.Alignment.Horizontal != HorizontalAlignment.Left))
        {
            return contentBounds;
        }

        var rightmostCol = cell.Column;
        if (isMergedOwner && mergedRange != null)
        {
            rightmostCol = mergedRange.Range.EndColumn;
        }

        var overflowRight = outerBounds.Right;
        var nextCol = rightmostCol + 1;

        while (columnOffsets.TryGetValue(nextCol, out var nextColLeft) &&
               columnByIndex.TryGetValue(nextCol, out var nextColInfo))
        {
            // Merged cells block text overflow (consistent with Excel's behavior)
            if (IsInsideMergedRange(mergedRangesByRow, cell.Row, nextCol))
            {
                break;
            }

            if (sheet.FindCell(cell.Row, nextCol) is { } adjacentCell)
            {
                if (!String.IsNullOrEmpty(adjacentCell.DisplayText))
                {
                    break;
                }

                if (adjacentCell.Style.Borders.Left.Style != BorderLineStyle.None)
                {
                    break;
                }
            }

            overflowRight = nextColLeft + nextColInfo.WidthPoint;
            nextCol++;
        }

        return contentBounds with
        {
            Width = Math.Max(contentBounds.Width, overflowRight - contentBounds.X)
        };
    }

    //--------------------------------------------------------------------------------
    // Header / Footer
    //--------------------------------------------------------------------------------

    private static PdfHeaderFooterRenderInfo BuildHeaderFooter(ReportSheet sheet, ReportRect pageBounds, ReportRect printableBounds, int pageNumber)
    {
        var headerText = sheet.HeaderFooter.DifferentFirst && (pageNumber == 1)
            ? sheet.HeaderFooter.FirstHeader
            : sheet.HeaderFooter.DifferentOddEven && (pageNumber % 2 == 0)
                ? sheet.HeaderFooter.EvenHeader
                : sheet.HeaderFooter.OddHeader;
        var footerText = sheet.HeaderFooter.DifferentFirst && (pageNumber == 1)
            ? sheet.HeaderFooter.FirstFooter
            : sheet.HeaderFooter.DifferentOddEven && (pageNumber % 2 == 0)
                ? sheet.HeaderFooter.EvenFooter
                : sheet.HeaderFooter.OddFooter;

        return new PdfHeaderFooterRenderInfo
        {
            HeaderText = headerText,
            FooterText = footerText,
            HeaderBounds = new ReportRect
            {
                X = printableBounds.X,
                Y = sheet.PageSetup.HeaderMarginPoint,
                Width = printableBounds.Width,
                Height = Math.Max(0d, sheet.PageSetup.Margins.Top - sheet.PageSetup.HeaderMarginPoint)
            },
            FooterBounds = new ReportRect
            {
                X = printableBounds.X,
                Y = pageBounds.Height - sheet.PageSetup.Margins.Bottom,
                Width = printableBounds.Width,
                Height = Math.Max(0d, sheet.PageSetup.Margins.Bottom - sheet.PageSetup.FooterMarginPoint)
            }
        };
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    private static void AddImageAnchorRowOffsets(
        ReportSheet sheet,
        ReportRange renderRange,
        List<ReportRow> rowsInRange,
        Dictionary<int, double> rowOffsets,
        double baseTop,
        double defaultRowHeight)
    {
        if (sheet.Images.Count == 0)
        {
            return;
        }

        var materialized = new HashSet<int>(rowsInRange.Count);
        foreach (var row in rowsInRange)
        {
            materialized.Add(row.Index);
        }

        foreach (var image in sheet.Images)
        {
            AddressHelper.ParseAddress(image.FromCellAddress, out var anchorRow, out _);

            if ((anchorRow < renderRange.StartRow) || (anchorRow > renderRange.EndRow) ||
                materialized.Contains(anchorRow) || rowOffsets.ContainsKey(anchorRow))
            {
                continue;
            }

            rowOffsets[anchorRow] = ResolveRowTop(rowsInRange, renderRange.StartRow, baseTop, defaultRowHeight, anchorRow);
        }
    }

    private static double ResolveRowTop(
        List<ReportRow> rowsInRange,
        int rangeStartRow,
        double baseTop,
        double defaultRowHeight,
        int targetRow)
    {
        var top = baseTop;
        var previousIndex = rangeStartRow - 1;
        foreach (var row in rowsInRange)
        {
            if (row.Index >= targetRow)
            {
                break;
            }

            top += (row.Index - previousIndex - 1) * defaultRowHeight;
            if (!row.IsHidden)
            {
                top += row.HeightPoint;
            }

            previousIndex = row.Index;
        }

        top += (targetRow - previousIndex - 1) * defaultRowHeight;
        return top;
    }

    private static List<PdfImageRenderInfo> BuildImageInfos(
        ReportSheet sheet,
        Dictionary<int, double> rowOffsets,
        Dictionary<int, double> columnOffsets)
    {
        // Converts images in the sheet from cell coordinates to point coordinates and build list
        var results = new List<PdfImageRenderInfo>();
        foreach (var image in sheet.Images)
        {
            AddressHelper.ParseAddress(image.FromCellAddress, out var row, out var column);
            if (!rowOffsets.TryGetValue(row, out var top) || !columnOffsets.TryGetValue(column, out var left))
            {
                continue;
            }

            results.Add(new PdfImageRenderInfo
            {
                Name = image.Name,
                Bounds = new ReportRect
                {
                    X = left + image.Offset.X,
                    Y = top + image.Offset.Y,
                    Width = image.WidthPoint,
                    Height = image.HeightPoint
                },
                ImageBytes = image.ImageBytes
            });
        }

        return results;
    }
}
