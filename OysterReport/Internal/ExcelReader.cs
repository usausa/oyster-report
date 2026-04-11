namespace OysterReport.Internal;

using System.Globalization;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

internal static class ExcelReader
{
    // Excel column width calculation (pixel conversion parameters per Excel spec)
    private const double DefaultMaxDigitWidth = 7d;
    private const double ExcelColumnPaddingMultiplier = 2d;
    private const double ExcelColumnPaddingDivisor = 4d;
    private const double ExcelColumnPaddingOffsetPixels = 1d;
    private const double ExcelColumnWidthGranularity = 256d;
    private const double ExcelColumnWidthRoundingOffset = 128d;
    private const double PointsPerInch = 72d;
    private const double ScreenDpi = 96d;

    //--------------------------------------------------------------------------------
    // Read
    //--------------------------------------------------------------------------------

    public static ReportWorkbook Read(IXLWorkbook workbook, ReportRenderOption? renderOption = null)
    {
        //Generate ReportWorkbook from a ClosedXML workbook
        var effectiveOptions = renderOption ?? new ReportRenderOption();
        var measurementProfile = CreateMeasurementProfile(workbook, effectiveOptions);
        var metadata = new ReportMetadata { TemplateName = workbook.Properties.Title ?? "Workbook" };
        return ReadInternal(workbook, measurementProfile, metadata);
    }

    public static ReportWorkbook Read(Stream stream, ReportRenderOption? renderOption = null)
    {
        using var workbook = new XLWorkbook(stream);
        return Read(workbook, renderOption);
    }

    public static ReportWorkbook Read(IXLWorksheet worksheet, ReportRenderOption? renderOption = null)
    {
        var effectiveOptions = renderOption ?? new ReportRenderOption();
        var measurementProfile = CreateMeasurementProfile(worksheet.Workbook, effectiveOptions);
        var metadata = new ReportMetadata { TemplateName = worksheet.Name };
        var reportWorkbook = new ReportWorkbook
        {
            Metadata = metadata,
            MeasurementProfile = measurementProfile
        };
        reportWorkbook.AddSheet(ReadSheet(worksheet, measurementProfile));
        return reportWorkbook;
    }

    //--------------------------------------------------------------------------------
    // Read internal
    //--------------------------------------------------------------------------------

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

    private static ReportMeasurementProfile CreateMeasurementProfile(IXLWorkbook workbook, ReportRenderOption renderOption)
    {
        // Creates a column width measurement profile from workbook default font settings.
        return new()
        {
            MaxDigitWidth = ResolveMaxDigitWidth(workbook.Style.Font.FontName, workbook.Style.Font.FontSize, renderOption),
            ColumnWidthAdjustment = renderOption.ColumnWidthAdjustment
        };
    }

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

        // Row
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

        // Column
        for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
        {
            var column = worksheet.Column(columnIndex);
            reportSheet.AddColumnDefinition(new ReportColumn
            {
                Index = columnIndex,
                WidthPoint = ConvertExcelColumnWidthToPoint(column.Width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment),
                IsHidden = column.IsHidden,
                OutlineLevel = column.OutlineLevel,
                OriginalExcelWidth = column.Width
            });
        }

        // Merged
        foreach (var mergedRange in worksheet.MergedRanges)
        {
            reportSheet.AddMergedRange(new ReportMergedRange
            {
                Range = new ReportRange
                {
                    StartRow = mergedRange.RangeAddress.FirstAddress.RowNumber,
                    StartColumn = mergedRange.RangeAddress.FirstAddress.ColumnNumber,
                    EndRow = mergedRange.RangeAddress.LastAddress.RowNumber,
                    EndColumn = mergedRange.RangeAddress.LastAddress.ColumnNumber
                }
            });
        }

        // Cell
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

        // Page break
        foreach (var pageBreak in worksheet.PageSetup.RowBreaks)
        {
            reportSheet.AddHorizontalPageBreak(new ReportPageBreak { Index = pageBreak });
        }

        foreach (var pageBreak in worksheet.PageSetup.ColumnBreaks)
        {
            reportSheet.AddVerticalPageBreak(new ReportPageBreak { Index = pageBreak });
        }

        // Picture
        foreach (var picture in worksheet.Pictures)
        {
            reportSheet.AddImage(ReadImage(picture));
        }

        reportSheet.RecalculateLayout();
        ApplyMergedRanges(reportSheet);
        ApplyTableStyles(reportSheet, worksheet);

        return reportSheet;
    }

    //--------------------------------------------------------------------------------
    // Print area
    //--------------------------------------------------------------------------------

    private static ReportPrintArea? ReadPrintArea(IXLWorksheet worksheet)
    {
        // Converts the print area to a ReportPrintArea if set
        var printArea = worksheet.PageSetup.PrintAreas.FirstOrDefault();
        if (printArea is null)
        {
            return null;
        }

        return new ReportPrintArea
        {
            Range = new ReportRange
            {
                StartRow = printArea.RangeAddress.FirstAddress.RowNumber,
                StartColumn = printArea.RangeAddress.FirstAddress.ColumnNumber,
                EndRow = printArea.RangeAddress.LastAddress.RowNumber,
                EndColumn = printArea.RangeAddress.LastAddress.ColumnNumber
            }
        };
    }

    //--------------------------------------------------------------------------------
    // Sheet range
    //--------------------------------------------------------------------------------

    private static bool TryResolveSheetRange(IXLWorksheet worksheet, ReportPrintArea? printArea, out ReportRange range)
    {
        // Merges the used cell range, print area, and merged ranges to determine the rendering range
        var contentRange = worksheet.RangeUsed();
        var formattedRange = worksheet.RangeUsed(XLCellsUsedOptions.All);
        if ((contentRange is null) && (formattedRange is null) && (worksheet.MergedRanges.Count == 0) && (printArea is null))
        {
            range = default;
            return false;
        }

        var startRow = Int32.MaxValue;
        var startColumn = Int32.MaxValue;
        var endRow = Int32.MinValue;
        var endColumn = Int32.MinValue;

        IncludeRange(contentRange);
        IncludeRange(formattedRange);
        if (printArea is not null)
        {
            IncludeReportRange(printArea.Range);
        }

        foreach (var mergedRange in worksheet.MergedRanges)
        {
            startRow = Math.Min(startRow, mergedRange.RangeAddress.FirstAddress.RowNumber);
            startColumn = Math.Min(startColumn, mergedRange.RangeAddress.FirstAddress.ColumnNumber);
            endRow = Math.Max(endRow, mergedRange.RangeAddress.LastAddress.RowNumber);
            endColumn = Math.Max(endColumn, mergedRange.RangeAddress.LastAddress.ColumnNumber);
        }

        if ((startRow == Int32.MaxValue) || (endRow == Int32.MinValue))
        {
            range = default;
            return false;
        }

        range = new ReportRange { StartRow = startRow, StartColumn = startColumn, EndRow = endRow, EndColumn = endColumn };
        return true;

        void IncludeRange(IXLRange? r)
        {
            if (r is null)
            {
                return;
            }

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

    //--------------------------------------------------------------------------------
    // Page setup
    //--------------------------------------------------------------------------------

    private static ReportPageSetup ReadPageSetup(IXLWorksheet worksheet)
    {
        // Converts page setup (paper size, margins, centering, etc.) to a ReportPageSetup
        return new()
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
    }

    private static ReportHeaderFooter ReadHeaderFooter(IXLWorksheet worksheet)
    {
        // Converts header/footer text and display conditions to a ReportHeaderFooter
        return new()
        {
            AlignWithMargins = worksheet.PageSetup.AlignHFWithMargins,
            DifferentFirst = !String.IsNullOrWhiteSpace(worksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage)) ||
                             !String.IsNullOrWhiteSpace(worksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage)),
            DifferentOddEven = !String.IsNullOrWhiteSpace(worksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages)) ||
                               !String.IsNullOrWhiteSpace(worksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages)),
            ScaleWithDocument = worksheet.PageSetup.ScaleHFWithDocument,
            OddHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.OddPages),
            OddFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.OddPages),
            EvenHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages),
            EvenFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages),
            FirstHeader = worksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage),
            FirstFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage)
        };
    }

    //--------------------------------------------------------------------------------
    // Cell
    //--------------------------------------------------------------------------------

    private static ReportCellValue ReadCellValue(IXLCell cell)
    {
        // Reads the cell value by type and converts it to a ReportCellValue
        return new()
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
    }

    private static ReportCellStyle ReadCellStyle(IXLCell cell)
    {
        // Converts the cell style (font, fill, borders, alignment) to a ReportCellStyle
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
        // Converts a border style and color to a ReportBorder. Transparent colors are corrected to black
        var resolvedColorHex = ColorHelper.NormalizeHex(colorHex);
        if ((styleValue != XLBorderStyleValues.None) && resolvedColorHex.StartsWith("#00", StringComparison.Ordinal))
        {
            resolvedColorHex = "#FF000000";
        }

        return new ReportBorder
        {
            Style = styleValue,
            ColorHex = resolvedColorHex,
            Width = ResolveBorderWidth(styleValue, new ReportRenderOption())
        };
    }

    private static string ResolveFillColorHex(IXLFill fill, IXLWorkbook workbook)
    {
        // Converts the cell fill background color to an ARGB hex string
        // Pattern fills prioritize the pattern color.
        if (fill.PatternType == XLFillPatternValues.None)
        {
            return "#00000000";
        }

        var background = ColorHelper.ResolveHex(fill.BackgroundColor, workbook, "#00000000");
        if (!background.StartsWith("#00", StringComparison.Ordinal))
        {
            return background;
        }

        return ColorHelper.ResolveHex(fill.PatternColor, workbook, "#00000000");
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    private static ReportImage ReadImage(IXLPicture picture)
    {
        // Converts ClosedXML image data to point-based coordinates and produces a ReportImage
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
        // Returns the bottom-right cell address for MoveAndSize-placed images only
        if (picture.Placement != XLPicturePlacement.MoveAndSize)
        {
            return null;
        }

        try
        {
            return picture.BottomRightCell?.Address.ToStringRelative(false);
        }
        catch (Exception ex) when (ex is NullReferenceException or InvalidOperationException)
        {
            return null;
        }
    }

    //--------------------------------------------------------------------------------
    // Post processing
    //--------------------------------------------------------------------------------

    private static void ApplyMergedRanges(ReportSheet reportSheet)
    {
        // Assigns merged-cell info to the Merge property of each affected cell
        foreach (var mergedRange in reportSheet.MergedRanges)
        {
            foreach (var cell in reportSheet.Cells.Where(x => mergedRange.Range.Contains(x.Row, x.Column)))
            {
                cell.Merge = new ReportMergeInfo
                {
                    OwnerCellAddress = mergedRange.OwnerCellAddress,
                    Range = mergedRange.Range
                };
            }
        }
    }

    private static void ApplyTableStyles(ReportSheet reportSheet, IXLWorksheet worksheet)
    {
        // Applies table styles (striped rows, etc.) to cell styles
        foreach (var table in worksheet.Tables)
        {
            if (!table.ShowRowStripes)
            {
                continue;
            }

            var themeName = table.Theme.ToString();
            if (!TableStyleCatalog.TryResolveBand1RowFillHex(themeName, worksheet.Workbook, out var stripeFillHex))
            {
                continue;
            }

            var tableRange = new ReportRange
            {
                StartRow = table.RangeAddress.FirstAddress.RowNumber,
                StartColumn = table.RangeAddress.FirstAddress.ColumnNumber,
                EndRow = table.RangeAddress.LastAddress.RowNumber,
                EndColumn = table.RangeAddress.LastAddress.ColumnNumber
            };

            var firstDataRow = tableRange.StartRow + (table.ShowHeaderRow ? 1 : 0);
            var lastDataRow = tableRange.EndRow - (table.ShowTotalsRow ? 1 : 0);

            for (var rowIndex = firstDataRow; rowIndex <= lastDataRow; rowIndex++)
            {
                if (((rowIndex - firstDataRow) % 2) != 0)
                {
                    continue;
                }

                foreach (var cell in reportSheet.Cells.Where(x => x.Row == rowIndex &&
                                                                  x.Column >= tableRange.StartColumn &&
                                                                  x.Column <= tableRange.EndColumn))
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

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private static double ConvertInchToPoint(double inch) => inch * 72d;

    private static bool IsTransparentFill(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);

    private static double ConvertExcelColumnWidthToPoint(double excelWidth, double maxDigitWidth, double adjustment)
    {
        // Converts Excel column width (in character units) to points
        // Calculated using max digit width and screen DPI per the Excel column width spec
        var normalizedWidth = Math.Max(0, excelWidth);
        var effectiveMaxDigitWidth = maxDigitWidth <= 0d ? DefaultMaxDigitWidth : maxDigitWidth;
        var pixelPadding = (ExcelColumnPaddingMultiplier * Math.Ceiling(effectiveMaxDigitWidth / ExcelColumnPaddingDivisor)) + ExcelColumnPaddingOffsetPixels;
        double pixelWidth;
        if (normalizedWidth < 1d)
        {
            pixelWidth = normalizedWidth * (effectiveMaxDigitWidth + pixelPadding);
        }
        else
        {
            var normalizedCharacters = ((ExcelColumnWidthGranularity * normalizedWidth) + Math.Round(ExcelColumnWidthRoundingOffset / effectiveMaxDigitWidth)) / ExcelColumnWidthGranularity;
            pixelWidth = (normalizedCharacters * effectiveMaxDigitWidth) + pixelPadding;
        }

        return pixelWidth * PointsPerInch / ScreenDpi * adjustment;
    }

    private static double ResolveMaxDigitWidth(string? fontName, double fontSize, ReportRenderOption renderOption)
    {
        // Resolves the maximum digit width for the workbook's default font
        if (String.IsNullOrWhiteSpace(fontName) || fontSize <= 0d)
        {
            return renderOption.FallbackMaxDigitWidth;
        }

        var directMeasured = FontMetricsHelper.MeasureMaxDigitWidth(fontName, fontSize);
        if (directMeasured is > 0d)
        {
            return Math.Max(renderOption.FallbackMaxDigitWidth, directMeasured.Value);
        }

        return renderOption.FallbackMaxDigitWidth;
    }

    private static double ResolveBorderWidth(XLBorderStyleValues style, ReportRenderOption renderOption) =>
        style switch
        {
            XLBorderStyleValues.Thick => renderOption.ThickBorderWidth,
            XLBorderStyleValues.Medium => renderOption.MediumBorderWidth,
            XLBorderStyleValues.Double => renderOption.NormalBorderWidth,
            XLBorderStyleValues.Hair => renderOption.HairBorderWidth,
            _ => renderOption.NormalBorderWidth
        };
}
