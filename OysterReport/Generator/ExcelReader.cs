namespace OysterReport.Generator;

using System.Globalization;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using OysterReport.Helpers;

internal sealed class ExcelReader
{
    public ReportWorkbook Read(IXLWorkbook workbook)
    {
        var measurementProfile = CreateMeasurementProfile(workbook);
        var metadata = new ReportMetadata { TemplateName = workbook.Properties.Title ?? "Workbook" };
        return ReadInternal(workbook, measurementProfile, metadata);
    }

    public ReportWorkbook Read(Stream stream)
    {
        using var workbook = new XLWorkbook(stream);
        return Read(workbook);
    }

    public ReportWorkbook Read(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        var measurementProfile = CreateMeasurementProfile(workbook);
        var metadata = new ReportMetadata
        {
            TemplateName = Path.GetFileNameWithoutExtension(filePath),
            SourceFilePath = filePath,
            SourceLastWriteTime = File.Exists(filePath) ? File.GetLastWriteTimeUtc(filePath) : null
        };
        return ReadInternal(workbook, measurementProfile, metadata);
    }

    private static ReportWorkbook ReadInternal(IXLWorkbook workbook, ReportMeasurementProfile measurementProfile, ReportMetadata metadata)
    {
        var reportWorkbook = new ReportWorkbook
        {
            Metadata = metadata,
            MeasurementProfile = measurementProfile
        };

        foreach (var worksheet in workbook.Worksheets)
        {
            reportWorkbook.AddSheet(ReadSheet(worksheet, measurementProfile));
        }

        return reportWorkbook;
    }

    private static ReportMeasurementProfile CreateMeasurementProfile(IXLWorkbook workbook) =>
        new()
        {
            DefaultFontName = workbook.Style.Font.FontName,
            DefaultFontSize = workbook.Style.Font.FontSize,
            MaxDigitWidth = FontMeasurementHelper.ResolveMaxDigitWidth(workbook.Style.Font.FontName, workbook.Style.Font.FontSize)
        };

    private static ReportSheet ReadSheet(IXLWorksheet worksheet, ReportMeasurementProfile measurementProfile)
    {
        var reportSheet = new ReportSheet { Name = worksheet.Name };
        var printArea = ReadPrintArea(worksheet);
        if (!TryResolveSheetRange(worksheet, printArea, out var range))
        {
            return reportSheet;
        }

        reportSheet.UsedRange = range;
        reportSheet.PageSetup = ReadPageSetup(worksheet);
        reportSheet.HeaderFooter = ReadHeaderFooter(worksheet);
        reportSheet.PrintArea = printArea;
        reportSheet.ShowGridLines = worksheet.PageSetup.ShowGridlines;

        for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
        {
            var row = worksheet.Row(rowIndex);
            reportSheet.AddRowDefinition(new ReportRow
            {
                Index = rowIndex,
                HeightPoint = row.Height,
                IsHidden = row.IsHidden,
                OutlineLevel = row.OutlineLevel
            });
        }

        for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
        {
            var column = worksheet.Column(columnIndex);
            var widthPoint = ColumnWidthConverter.ToPoint(column.Width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment);
            reportSheet.AddColumnDefinition(new ReportColumn
            {
                Index = columnIndex,
                WidthPoint = widthPoint,
                IsHidden = column.IsHidden,
                OutlineLevel = column.OutlineLevel,
                OriginalExcelWidth = column.Width
            });
        }

        foreach (var mergedRange in worksheet.MergedRanges)
        {
            reportSheet.AddMergedRange(new ReportMergedRange
            {
                Range = new ReportRange(
                    mergedRange.RangeAddress.FirstAddress.RowNumber,
                    mergedRange.RangeAddress.FirstAddress.ColumnNumber,
                    mergedRange.RangeAddress.LastAddress.RowNumber,
                    mergedRange.RangeAddress.LastAddress.ColumnNumber)
            });
        }

        for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
        {
            for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
            {
                var cell = worksheet.Cell(rowIndex, columnIndex);
                reportSheet.AddCell(new ReportCell
                {
                    Row = rowIndex,
                    Column = columnIndex,
                    Value = ReadCellValue(cell),
                    DisplayText = cell.GetFormattedString(),
                    Style = ReadCellStyle(cell)
                });
            }
        }

        foreach (var pageBreak in worksheet.PageSetup.RowBreaks)
        {
            reportSheet.AddHorizontalPageBreak(new ReportPageBreak { Index = pageBreak });
        }

        foreach (var pageBreak in worksheet.PageSetup.ColumnBreaks)
        {
            reportSheet.AddVerticalPageBreak(new ReportPageBreak { Index = pageBreak });
        }

        foreach (var picture in worksheet.Pictures)
        {
            reportSheet.AddImage(ReadImage(picture));
        }

        reportSheet.RecalculateLayout();
        ApplyMergedRanges(reportSheet);
        ApplyTableStyles(reportSheet, worksheet);
        return reportSheet;
    }

    private static ReportCellValue ReadCellValue(IXLCell cell) =>
        new()
        {
            Kind = cell.DataType,
            RawValue = cell.DataType switch
            {
                XLDataType.Boolean => cell.Value.GetBoolean(),
                XLDataType.Number => cell.Value.GetNumber(),
                XLDataType.DateTime => cell.Value.GetDateTime(),
                XLDataType.Text => cell.Value.GetText(),
                _ => cell.Value.ToString(CultureInfo.InvariantCulture)
            }
        };

    private static ReportCellStyle ReadCellStyle(IXLCell cell)
    {
        var style = cell.Style;
        return new ReportCellStyle
        {
            Font = new ReportFont
            {
                Name = style.Font.FontName,
                Size = style.Font.FontSize,
                Bold = style.Font.Bold,
                Italic = style.Font.Italic,
                Underline = style.Font.Underline != XLFontUnderlineValues.None,
                Strikeout = style.Font.Strikethrough,
                ColorHex = ColorHelper.ResolveHex(style.Font.FontColor, cell.Worksheet.Workbook, "#FF000000")
            },
            Fill = new ReportFill
            {
                BackgroundColorHex = ResolveFillColorHex(style.Fill, cell.Worksheet.Workbook)
            },
            Borders = new ReportBorders
            {
                Left = ReadBorder(style.Border.LeftBorder, ColorHelper.ResolveHex(style.Border.LeftBorderColor, cell.Worksheet.Workbook, "#FF000000")),
                Top = ReadBorder(style.Border.TopBorder, ColorHelper.ResolveHex(style.Border.TopBorderColor, cell.Worksheet.Workbook, "#FF000000")),
                Right = ReadBorder(style.Border.RightBorder, ColorHelper.ResolveHex(style.Border.RightBorderColor, cell.Worksheet.Workbook, "#FF000000")),
                Bottom = ReadBorder(style.Border.BottomBorder, ColorHelper.ResolveHex(style.Border.BottomBorderColor, cell.Worksheet.Workbook, "#FF000000"))
            },
            Alignment = new ReportAlignment
            {
                Horizontal = style.Alignment.Horizontal,
                Vertical = style.Alignment.Vertical
            },
            WrapText = style.Alignment.WrapText
        };
    }

    private static ReportBorder ReadBorder(XLBorderStyleValues styleValue, string colorHex)
    {
        var resolvedColorHex = ColorHelper.NormalizeHex(colorHex);
        if (styleValue != XLBorderStyleValues.None && resolvedColorHex.StartsWith("#00", StringComparison.Ordinal))
        {
            resolvedColorHex = "#FF000000";
        }

        return new ReportBorder
        {
            Style = styleValue,
            ColorHex = resolvedColorHex,
            Width = PdfRenderingConstants.ResolveBorderWidth(styleValue)
        };
    }

    private static ReportPageSetup ReadPageSetup(IXLWorksheet worksheet) =>
        new()
        {
            PaperSize = worksheet.PageSetup.PaperSize,
            Orientation = worksheet.PageSetup.PageOrientation,
            Margins = new ReportThickness
            {
                Left = ConvertInchToPoint(worksheet.PageSetup.Margins.Left),
                Top = ConvertInchToPoint(worksheet.PageSetup.Margins.Top),
                Right = ConvertInchToPoint(worksheet.PageSetup.Margins.Right),
                Bottom = ConvertInchToPoint(worksheet.PageSetup.Margins.Bottom)
            },
            HeaderMarginPoint = ConvertInchToPoint(worksheet.PageSetup.Margins.Header),
            FooterMarginPoint = ConvertInchToPoint(worksheet.PageSetup.Margins.Footer),
            ScalePercent = worksheet.PageSetup.Scale,
            FitToPagesWide = worksheet.PageSetup.PagesWide == 0 ? null : worksheet.PageSetup.PagesWide,
            FitToPagesTall = worksheet.PageSetup.PagesTall == 0 ? null : worksheet.PageSetup.PagesTall,
            CenterHorizontally = worksheet.PageSetup.CenterHorizontally,
            CenterVertically = worksheet.PageSetup.CenterVertically
        };

    private static ReportHeaderFooter ReadHeaderFooter(IXLWorksheet worksheet) =>
        new()
        {
            AlignWithMargins = worksheet.PageSetup.AlignHFWithMargins,
            DifferentFirst = !string.IsNullOrWhiteSpace(worksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage)) ||
                             !string.IsNullOrWhiteSpace(worksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage)),
            DifferentOddEven = !string.IsNullOrWhiteSpace(worksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages)) ||
                               !string.IsNullOrWhiteSpace(worksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages)),
            ScaleWithDocument = worksheet.PageSetup.ScaleHFWithDocument,
            OddHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.OddPages),
            OddFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.OddPages),
            EvenHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages),
            EvenFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages),
            FirstHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage),
            FirstFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage)
        };

    private static ReportPrintArea? ReadPrintArea(IXLWorksheet worksheet)
    {
        var printArea = worksheet.PageSetup.PrintAreas.FirstOrDefault();
        if (printArea is null)
        {
            return null;
        }

        return new ReportPrintArea
        {
            Range = new ReportRange(
                printArea.RangeAddress.FirstAddress.RowNumber,
                printArea.RangeAddress.FirstAddress.ColumnNumber,
                printArea.RangeAddress.LastAddress.RowNumber,
                printArea.RangeAddress.LastAddress.ColumnNumber)
        };
    }

    private static bool TryResolveSheetRange(IXLWorksheet worksheet, ReportPrintArea? printArea, out ReportRange range)
    {
        var contentRange = worksheet.RangeUsed();
        var formattedRange = worksheet.RangeUsed(XLCellsUsedOptions.All);
        if (contentRange is null && formattedRange is null && worksheet.MergedRanges.Count == 0 && printArea is null)
        {
            range = default;
            return false;
        }

        var startRow = int.MaxValue;
        var startColumn = int.MaxValue;
        var endRow = int.MinValue;
        var endColumn = int.MinValue;

        IncludeRange(contentRange);
        IncludeRange(formattedRange);
        if (printArea is not null) IncludeReportRange(printArea.Range);

        foreach (var mergedRange in worksheet.MergedRanges)
        {
            startRow = Math.Min(startRow, mergedRange.RangeAddress.FirstAddress.RowNumber);
            startColumn = Math.Min(startColumn, mergedRange.RangeAddress.FirstAddress.ColumnNumber);
            endRow = Math.Max(endRow, mergedRange.RangeAddress.LastAddress.RowNumber);
            endColumn = Math.Max(endColumn, mergedRange.RangeAddress.LastAddress.ColumnNumber);
        }

        if (startRow == int.MaxValue || endRow == int.MinValue)
        {
            range = default;
            return false;
        }

        range = new ReportRange(startRow, startColumn, endRow, endColumn);
        return true;

        void IncludeRange(IXLRange? r)
        {
            if (r is null) return;
            startRow = Math.Min(startRow, r.RangeAddress.FirstAddress.RowNumber);
            startColumn = Math.Min(startColumn, r.RangeAddress.FirstAddress.ColumnNumber);
            endRow = Math.Max(endRow, r.RangeAddress.LastAddress.RowNumber);
            endColumn = Math.Max(endColumn, r.RangeAddress.LastAddress.ColumnNumber);
        }

        void IncludeReportRange(ReportRange r)
        {
            startRow = Math.Min(startRow, r.StartRow);
            startColumn = Math.Min(startColumn, r.StartColumn);
            endRow = Math.Max(endRow, r.EndRow);
            endColumn = Math.Max(endColumn, r.EndColumn);
        }
    }

    private static ReportImage ReadImage(IXLPicture picture)
    {
        using var memoryStream = new MemoryStream();
        picture.ImageStream.Position = 0;
        picture.ImageStream.CopyTo(memoryStream);
        return new ReportImage
        {
            Name = picture.Name,
            FromCellAddress = picture.TopLeftCell.Address.ToStringRelative(false),
            ToCellAddress = TryGetBottomRightCellAddress(picture),
            Offset = new ReportOffset
            {
                X = picture.Left * 72d / 96d,
                Y = picture.Top * 72d / 96d
            },
            WidthPoint = picture.Width * 72d / 96d,
            HeightPoint = picture.Height * 72d / 96d,
            ImageBytes = memoryStream.ToArray()
        };
    }

    private static string? TryGetBottomRightCellAddress(IXLPicture picture)
    {
        if (picture.Placement != XLPicturePlacement.MoveAndSize) return null;

        try
        {
            return picture.BottomRightCell?.Address.ToStringRelative(false);
        }
        catch (Exception ex) when (ex is NullReferenceException or InvalidOperationException)
        {
            return null;
        }
    }

    private static string ResolveFillColorHex(IXLFill fill, IXLWorkbook workbook)
    {
        if (fill.PatternType == XLFillPatternValues.None) return "#00000000";

        var background = ColorHelper.ResolveHex(fill.BackgroundColor, workbook, "#00000000");
        if (!background.StartsWith("#00", StringComparison.Ordinal)) return background;

        return ColorHelper.ResolveHex(fill.PatternColor, workbook, "#00000000");
    }

    private static void ApplyTableStyles(ReportSheet reportSheet, IXLWorksheet worksheet)
    {
        foreach (var table in worksheet.Tables)
        {
            if (!table.ShowRowStripes)
            {
                continue;
            }

            var themeName = table.Theme.ToString();
            if (!string.Equals(themeName, "TableStyleLight4", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            const string stripeFillHex = "#FFDEEBF7";

            var tableRange = new ReportRange(
                table.RangeAddress.FirstAddress.RowNumber,
                table.RangeAddress.FirstAddress.ColumnNumber,
                table.RangeAddress.LastAddress.RowNumber,
                table.RangeAddress.LastAddress.ColumnNumber);

            var firstDataRow = tableRange.StartRow + (table.ShowHeaderRow ? 1 : 0);
            var lastDataRow = tableRange.EndRow - (table.ShowTotalsRow ? 1 : 0);

            for (var rowIndex = firstDataRow; rowIndex <= lastDataRow; rowIndex++)
            {
                if (((rowIndex - firstDataRow) % 2) != 0)
                {
                    continue;
                }

                foreach (var cell in reportSheet.Cells.Where(cell =>
                             cell.Row == rowIndex &&
                             cell.Column >= tableRange.StartColumn &&
                             cell.Column <= tableRange.EndColumn))
                {
                    if (!IsTransparentFill(cell.Style.Fill.BackgroundColorHex))
                    {
                        continue;
                    }

                    cell.Style = cell.Style with
                    {
                        Fill = new ReportFill { BackgroundColorHex = stripeFillHex }
                    };
                }
            }
        }
    }

    private static void ApplyMergedRanges(ReportSheet reportSheet)
    {
        foreach (var mergedRange in reportSheet.MergedRanges)
        {
            foreach (var cell in reportSheet.Cells.Where(cell => mergedRange.Range.Contains(cell.Row, cell.Column)))
            {
                cell.Merge = new ReportMergeInfo
                {
                    OwnerCellAddress = mergedRange.OwnerCellAddress,
                    Range = mergedRange.Range
                };
            }
        }
    }

    private static double ConvertInchToPoint(double inch) => inch * 72d;

    private static bool IsTransparentFill(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);
}
