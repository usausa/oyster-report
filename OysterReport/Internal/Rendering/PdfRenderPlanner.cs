namespace OysterReport.Internal.Rendering;

using System.Globalization;
using OysterReport.Common;
using OysterReport.Common.Geometry;
using OysterReport.Helpers;
using OysterReport.Model;
using OysterReport.Writing.Pdf;

internal static class PdfRenderPlanner
{
    public static PdfRenderPlan BuildPlan(ReportWorkbook workbook, PdfGenerateOptions options)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(options);

        var sheets = workbook.Sheets.Select((sheet, index) => BuildSheetPlan(sheet, workbook.MeasurementProfile, index + 1)).ToList();
        return new PdfRenderPlan
        {
            Sheets = sheets,
        };
    }

    private static PdfRenderSheetPlan BuildSheetPlan(ReportSheet sheet, ReportMeasurementProfile measurementProfile, int sheetNumber)
    {
        var pageBounds = ResolvePageBounds(sheet.PageSetup);
        var printableBounds = new ReportRect
        {
            X = sheet.PageSetup.Margins.Left,
            Y = sheet.PageSetup.Margins.Top,
            Width = pageBounds.Width - sheet.PageSetup.Margins.Left - sheet.PageSetup.Margins.Right,
            Height = pageBounds.Height - sheet.PageSetup.Margins.Top - sheet.PageSetup.Margins.Bottom,
        };

        var renderRange = sheet.PrintArea?.Range ?? sheet.UsedRange;

        var visibleRows = sheet.Rows
            .Where(row => !row.IsHidden && row.Index >= renderRange.StartRow && row.Index <= renderRange.EndRow)
            .OrderBy(row => row.Index)
            .ToList();
        var visibleColumns = sheet.Columns
            .Where(column => !column.IsHidden && column.Index >= renderRange.StartColumn && column.Index <= renderRange.EndColumn)
            .OrderBy(column => column.Index)
            .ToList();

        var contentHeight = visibleRows.Sum(row => row.HeightPoint);
        var contentWidth = visibleColumns.Sum(column => column.WidthPoint);
        var contentOffsetX = sheet.PageSetup.CenterHorizontally
            ? Math.Max(0d, (printableBounds.Width - contentWidth) / 2d)
            : 0d;
        var contentOffsetY = sheet.PageSetup.CenterVertically
            ? Math.Max(0d, (printableBounds.Height - contentHeight) / 2d)
            : 0d;

        var rowOffsets = new Dictionary<int, double>();
        var currentTop = printableBounds.Y + contentOffsetY;
        foreach (var row in visibleRows)
        {
            rowOffsets[row.Index] = currentTop;
            currentTop += row.HeightPoint;
        }

        var columnOffsets = new Dictionary<int, double>();
        var currentLeft = printableBounds.X + contentOffsetX;
        foreach (var column in visibleColumns)
        {
            columnOffsets[column.Index] = currentLeft;
            currentLeft += column.WidthPoint;
        }

        var mergedRanges = sheet.MergedRanges.ToDictionary(range => range.OwnerCellAddress, range => range);
        var pageCells = new List<PdfCellRenderInfo>();
        foreach (var cell in sheet.Cells.Where(cell => rowOffsets.ContainsKey(cell.Row) && columnOffsets.ContainsKey(cell.Column)))
        {
            var outerBounds = new ReportRect
            {
                X = columnOffsets[cell.Column],
                Y = rowOffsets[cell.Row],
                Width = visibleColumns.First(column => column.Index == cell.Column).WidthPoint,
                Height = visibleRows.First(row => row.Index == cell.Row).HeightPoint,
            };

            var isMergedOwner = mergedRanges.TryGetValue(cell.Address, out var mergedRange);
            if (isMergedOwner && mergedRange is not null)
            {
                outerBounds = BuildMergedBounds(mergedRange, visibleRows, visibleColumns, rowOffsets, columnOffsets);
            }

            var contentBounds = outerBounds.Deflate(ReportThickness.Uniform(2d));
            pageCells.Add(new PdfCellRenderInfo
            {
                CellAddress = cell.Address,
                OuterBounds = outerBounds,
                ContentBounds = contentBounds,
                TextBounds = EstimateTextBounds(contentBounds, cell.DisplayText, measurementProfile, cell.Style.Font.Size),
                IsMergedOwner = isMergedOwner,
                IsClipped = false,
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
                    Cells = pageCells,
                },
            ],
            Borders = BuildBorderInfos(sheet, pageCells),
            Images = BuildImageInfos(sheet, rowOffsets, columnOffsets),
        };
    }

    private static ReportRect ResolvePageBounds(ReportPageSetup pageSetup)
    {
        var (width, height) = PageSizeResolver.GetPageSize(pageSetup.PaperSize);
        return pageSetup.Orientation == ReportPageOrientation.Landscape
            ? new ReportRect { X = 0, Y = 0, Width = height, Height = width }
            : new ReportRect { X = 0, Y = 0, Width = width, Height = height };
    }

    private static ReportRect BuildMergedBounds(
        ReportMergedRange mergedRange,
        IReadOnlyList<ReportRow> visibleRows,
        IReadOnlyList<ReportColumn> visibleColumns,
        Dictionary<int, double> rowOffsets,
        Dictionary<int, double> columnOffsets)
    {
        var targetRows = visibleRows.Where(row => row.Index >= mergedRange.Range.StartRow && row.Index <= mergedRange.Range.EndRow).ToList();
        var targetColumns = visibleColumns.Where(column => column.Index >= mergedRange.Range.StartColumn && column.Index <= mergedRange.Range.EndColumn).ToList();
        if (targetRows.Count == 0 || targetColumns.Count == 0)
        {
            return default;
        }

        return new ReportRect
        {
            X = columnOffsets[targetColumns[0].Index],
            Y = rowOffsets[targetRows[0].Index],
            Width = targetColumns.Sum(column => column.WidthPoint),
            Height = targetRows.Sum(row => row.HeightPoint),
        };
    }

    private static ReportRect EstimateTextBounds(ReportRect contentBounds, string text, ReportMeasurementProfile measurementProfile, double fontSize)
    {
        var effectiveFontSize = fontSize <= 0 ? measurementProfile.DefaultFontSize : fontSize;
        var width = Math.Min(contentBounds.Width, Math.Max(0, text.Length * effectiveFontSize * 0.55d));
        var height = Math.Min(contentBounds.Height, Math.Max(effectiveFontSize + 2d, 0));
        return new ReportRect
        {
            X = contentBounds.X,
            Y = contentBounds.Y,
            Width = width,
            Height = height,
        };
    }

    private static PdfHeaderFooterRenderInfo BuildHeaderFooter(ReportSheet sheet, ReportRect pageBounds, ReportRect printableBounds, int pageNumber)
    {
        var headerText = sheet.HeaderFooter.DifferentFirst && pageNumber == 1
            ? sheet.HeaderFooter.FirstHeader
            : sheet.HeaderFooter.DifferentOddEven && (pageNumber % 2 == 0)
                ? sheet.HeaderFooter.EvenHeader
                : sheet.HeaderFooter.OddHeader;
        var footerText = sheet.HeaderFooter.DifferentFirst && pageNumber == 1
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
                Height = Math.Max(0d, sheet.PageSetup.Margins.Top - sheet.PageSetup.HeaderMarginPoint),
            },
            FooterBounds = new ReportRect
            {
                X = printableBounds.X,
                Y = pageBounds.Height - sheet.PageSetup.Margins.Bottom,
                Width = printableBounds.Width,
                Height = Math.Max(0d, sheet.PageSetup.Margins.Bottom - sheet.PageSetup.FooterMarginPoint),
            },
        };
    }

    private static List<PdfBorderRenderInfo> BuildBorderInfos(ReportSheet sheet, IReadOnlyList<PdfCellRenderInfo> cellInfos)
    {
        var borderInfos = new Dictionary<string, PdfBorderRenderInfo>(StringComparer.Ordinal);
        foreach (var cellInfo in cellInfos)
        {
            var cell = sheet.Cells.First(sourceCell => sourceCell.Address == cellInfo.CellAddress);
            AddBorder(
                borderInfos,
                BuildLineKey("L", cellInfo.OuterBounds.X, cellInfo.OuterBounds.Y, cellInfo.OuterBounds.X, cellInfo.OuterBounds.Bottom),
                new ReportLine
                {
                    X1 = cellInfo.OuterBounds.X,
                    Y1 = cellInfo.OuterBounds.Y,
                    X2 = cellInfo.OuterBounds.X,
                    Y2 = cellInfo.OuterBounds.Bottom,
                },
                cell.Style.Borders.Left,
                cell.Address);
            AddBorder(
                borderInfos,
                BuildLineKey("T", cellInfo.OuterBounds.X, cellInfo.OuterBounds.Y, cellInfo.OuterBounds.Right, cellInfo.OuterBounds.Y),
                new ReportLine
                {
                    X1 = cellInfo.OuterBounds.X,
                    Y1 = cellInfo.OuterBounds.Y,
                    X2 = cellInfo.OuterBounds.Right,
                    Y2 = cellInfo.OuterBounds.Y,
                },
                cell.Style.Borders.Top,
                cell.Address);
            AddBorder(
                borderInfos,
                BuildLineKey("R", cellInfo.OuterBounds.Right, cellInfo.OuterBounds.Y, cellInfo.OuterBounds.Right, cellInfo.OuterBounds.Bottom),
                new ReportLine
                {
                    X1 = cellInfo.OuterBounds.Right,
                    Y1 = cellInfo.OuterBounds.Y,
                    X2 = cellInfo.OuterBounds.Right,
                    Y2 = cellInfo.OuterBounds.Bottom,
                },
                cell.Style.Borders.Right,
                cell.Address);
            AddBorder(
                borderInfos,
                BuildLineKey("B", cellInfo.OuterBounds.X, cellInfo.OuterBounds.Bottom, cellInfo.OuterBounds.Right, cellInfo.OuterBounds.Bottom),
                new ReportLine
                {
                    X1 = cellInfo.OuterBounds.X,
                    Y1 = cellInfo.OuterBounds.Bottom,
                    X2 = cellInfo.OuterBounds.Right,
                    Y2 = cellInfo.OuterBounds.Bottom,
                },
                cell.Style.Borders.Bottom,
                cell.Address);
        }

        return borderInfos.Values.ToList();
    }

    private static void AddBorder(
        Dictionary<string, PdfBorderRenderInfo> borders,
        string key,
        ReportLine line,
        ReportBorder border,
        string ownerCellAddress)
    {
        if (border.Style == ReportBorderStyle.None)
        {
            return;
        }

        var candidate = new PdfBorderRenderInfo
        {
            Line = line,
            Style = border.Style,
            Width = border.Width,
            ColorHex = border.ColorHex,
            OwnerCellAddress = ownerCellAddress,
        };

        if (!borders.TryGetValue(key, out var existingBorder) || GetBorderPriority(candidate.Style) > GetBorderPriority(existingBorder.Style))
        {
            borders[key] = candidate;
        }
    }

    private static List<PdfImageRenderInfo> BuildImageInfos(
        ReportSheet sheet,
        Dictionary<int, double> rowOffsets,
        Dictionary<int, double> columnOffsets)
    {
        var results = new List<PdfImageRenderInfo>();
        foreach (var image in sheet.Images)
        {
            var (row, column) = AddressHelper.ParseAddress(image.FromCellAddress);
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
                    Height = image.HeightPoint,
                },
                ImageBytes = image.ImageBytes,
            });
        }

        return results;
    }

    private static string BuildLineKey(string prefix, double x1, double y1, double x2, double y2) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{prefix}:{Math.Round(x1, 4)}:{Math.Round(y1, 4)}:{Math.Round(x2, 4)}:{Math.Round(y2, 4)}");

    private static int GetBorderPriority(ReportBorderStyle style) =>
        style switch
        {
            ReportBorderStyle.DoubleLine => 7,
            ReportBorderStyle.Thick => 6,
            ReportBorderStyle.Medium => 5,
            ReportBorderStyle.Thin => 4,
            ReportBorderStyle.Dashed => 3,
            ReportBorderStyle.DashDot => 2,
            ReportBorderStyle.Dotted => 1,
            ReportBorderStyle.Hair => 0,
            _ => -1,
        };
}
