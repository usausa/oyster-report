namespace OysterReport.Reading;

using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using OysterReport.Common;
using OysterReport.Helpers;
using OysterReport.Model;

public sealed class ExcelReader
{
    private readonly StringComparer sheetNameComparer = StringComparer.OrdinalIgnoreCase;

    public ReportWorkbook Read(string filePath, ExcelReadOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        using var stream = File.OpenRead(filePath);
        var metadata = new ReportMetadata
        {
            TemplateName = Path.GetFileNameWithoutExtension(filePath),
            SourceFilePath = filePath,
            SourceLastWriteTime = File.Exists(filePath) ? File.GetLastWriteTimeUtc(filePath) : null
        };
        return ReadInternal(stream, options, metadata);
    }

    public ReportWorkbook Read(Stream stream, ExcelReadOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return ReadInternal(stream, options, null);
    }

    private ReportWorkbook ReadInternal(Stream stream, ExcelReadOptions? options, ReportMetadata? metadata)
    {
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        using var workbookStream = new MemoryStream();
        stream.CopyTo(workbookStream);
        var workbookBytes = workbookStream.ToArray();
        workbookStream.Position = 0;

        var workbookTableStyles = WorkbookTableStyleMap.Load(workbookBytes);
        var rawColumnWidths = WorkbookRawColumnWidths.Load(workbookBytes);
        using var workbook = new XLWorkbook(workbookStream);
        var measurementProfile = new ReportMeasurementProfile
        {
            DefaultFontName = workbook.Style.Font.FontName,
            DefaultFontSize = workbook.Style.Font.FontSize,
            MaxDigitWidth = FontMeasurementHelper.ResolveMaxDigitWidth(workbook.Style.Font.FontName, workbook.Style.Font.FontSize)
        };
        var reportWorkbook = new ReportWorkbook(
            metadata ?? new ReportMetadata { TemplateName = workbook.Properties.Title ?? "Workbook" },
            measurementProfile);
        var targetSheets = options?.TargetSheets is { Count: > 0 }
            ? new HashSet<string>(options.TargetSheets, sheetNameComparer)
            : null;

        foreach (var (worksheet, worksheetIndex) in workbook.Worksheets.Select((worksheet, index) => (worksheet, index)))
        {
            if (targetSheets is not null && !targetSheets.Contains(worksheet.Name))
            {
                continue;
            }

            var sheetTableStyles = workbookTableStyles.GetTableStyles(worksheetIndex);
            reportWorkbook.AddSheet(ReadSheet(worksheet, measurementProfile, options?.IncludeImages ?? true, sheetTableStyles, worksheetIndex, rawColumnWidths));
        }

        return reportWorkbook;
    }

    private static ReportSheet ReadSheet(
        IXLWorksheet worksheet,
        ReportMeasurementProfile measurementProfile,
        bool includeImages,
        IReadOnlyList<TableStyleInfo> tableStyles,
        int sheetIndex,
        WorkbookRawColumnWidths rawColumnWidths)
    {
        var reportSheet = new ReportSheet(worksheet.Name);
        var printArea = ReadPrintArea(worksheet);
        if (!TryResolveSheetRange(worksheet, printArea, out var range))
        {
            return reportSheet;
        }

        reportSheet.SetUsedRange(range);
        reportSheet.SetPageSetup(ReadPageSetup(worksheet));
        reportSheet.SetHeaderFooter(ReadHeaderFooter(worksheet));
        reportSheet.SetPrintArea(printArea);
        reportSheet.SetShowGridLines(worksheet.PageSetup.ShowGridlines);

        for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
        {
            var row = worksheet.Row(rowIndex);
            reportSheet.AddRowDefinition(new ReportRow(rowIndex, row.Height, row.IsHidden, row.OutlineLevel));
        }

        for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
        {
            var column = worksheet.Column(columnIndex);
            var rawWidth = rawColumnWidths.TryGetRawWidth(sheetIndex, columnIndex);
            var widthPoint = rawWidth.HasValue
                ? Math.Ceiling(rawWidth.Value * measurementProfile.MaxDigitWidth) * (72d / 96d) * measurementProfile.ColumnWidthAdjustment
                : ColumnWidthConverter.ToPoint(column.Width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment);
            reportSheet.AddColumnDefinition(new ReportColumn(columnIndex, widthPoint, column.IsHidden, column.OutlineLevel, column.Width));
        }

        foreach (var mergedRange in worksheet.MergedRanges)
        {
            reportSheet.AddMergedRange(new ReportMergedRange(new ReportRange(
                mergedRange.RangeAddress.FirstAddress.RowNumber,
                mergedRange.RangeAddress.FirstAddress.ColumnNumber,
                mergedRange.RangeAddress.LastAddress.RowNumber,
                mergedRange.RangeAddress.LastAddress.ColumnNumber)));
        }

        for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
        {
            for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
            {
                var cell = worksheet.Cell(rowIndex, columnIndex);
                var displayText = cell.GetFormattedString();
                ReportPlaceholderText? placeholder = null;
                if (PlaceholderParser.TryParse(displayText, out var markerName))
                {
                    placeholder = new ReportPlaceholderText(displayText, markerName);
                }

                reportSheet.AddCell(new ReportCell(
                    rowIndex,
                    columnIndex,
                    ReadCellValue(cell),
                    displayText,
                    displayText,
                    ReadCellStyle(cell),
                    placeholder));
            }
        }

        foreach (var pageBreak in worksheet.PageSetup.RowBreaks)
        {
            reportSheet.AddHorizontalPageBreak(new ReportPageBreak { Index = pageBreak, IsHorizontal = true });
        }

        foreach (var pageBreak in worksheet.PageSetup.ColumnBreaks)
        {
            reportSheet.AddVerticalPageBreak(new ReportPageBreak { Index = pageBreak, IsHorizontal = false });
        }

        if (includeImages)
        {
            foreach (var picture in worksheet.Pictures)
            {
                reportSheet.AddImage(ReadImage(picture));
            }
        }

        reportSheet.RecalculateLayout();
        ApplyMergedRanges(reportSheet);
        ApplyTableStyles(reportSheet, worksheet.Workbook, tableStyles);
        return reportSheet;
    }

    private static ReportCellValue ReadCellValue(IXLCell cell)
    {
        return cell.DataType switch
        {
            XLDataType.Boolean => new ReportCellValue { Kind = ReportCellValueKind.Boolean, RawValue = cell.Value.GetBoolean() },
            XLDataType.Number => new ReportCellValue { Kind = ReportCellValueKind.Number, RawValue = cell.Value.GetNumber() },
            XLDataType.DateTime => new ReportCellValue { Kind = ReportCellValueKind.DateTime, RawValue = cell.Value.GetDateTime() },
            XLDataType.Text => new ReportCellValue { Kind = ReportCellValueKind.Text, RawValue = cell.Value.GetText() },
            XLDataType.Error => new ReportCellValue { Kind = ReportCellValueKind.Error, RawValue = cell.Value.ToString(CultureInfo.InvariantCulture) },
            _ => new ReportCellValue { Kind = ReportCellValueKind.Blank, RawValue = cell.Value.ToString(CultureInfo.InvariantCulture) }
        };
    }

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
                Horizontal = style.Alignment.Horizontal switch
                {
                    XLAlignmentHorizontalValues.General => ReportHorizontalAlignment.General,
                    XLAlignmentHorizontalValues.CenterContinuous => ReportHorizontalAlignment.Center,
                    XLAlignmentHorizontalValues.Center => ReportHorizontalAlignment.Center,
                    XLAlignmentHorizontalValues.Right => ReportHorizontalAlignment.Right,
                    XLAlignmentHorizontalValues.Justify => ReportHorizontalAlignment.Justify,
                    _ => ReportHorizontalAlignment.Left
                },
                Vertical = style.Alignment.Vertical switch
                {
                    XLAlignmentVerticalValues.Center => ReportVerticalAlignment.Center,
                    XLAlignmentVerticalValues.Bottom => ReportVerticalAlignment.Bottom,
                    XLAlignmentVerticalValues.Justify => ReportVerticalAlignment.Justify,
                    _ => ReportVerticalAlignment.Top
                }
            },
            NumberFormat = style.NumberFormat.Format,
            WrapText = style.Alignment.WrapText,
            Rotation = style.Alignment.TextRotation,
            ShrinkToFit = style.Alignment.ShrinkToFit
        };
    }

    private static ReportBorder ReadBorder(XLBorderStyleValues styleValue, string colorHex)
    {
        var style = styleValue switch
        {
            XLBorderStyleValues.Thick => ReportBorderStyle.Thick,
            XLBorderStyleValues.Medium => ReportBorderStyle.Medium,
            XLBorderStyleValues.Double => ReportBorderStyle.DoubleLine,
            XLBorderStyleValues.Dashed => ReportBorderStyle.Dashed,
            XLBorderStyleValues.Dotted => ReportBorderStyle.Dotted,
            XLBorderStyleValues.Hair => ReportBorderStyle.Hair,
            XLBorderStyleValues.DashDot => ReportBorderStyle.DashDot,
            XLBorderStyleValues.None => ReportBorderStyle.None,
            _ => ReportBorderStyle.Thin
        };

        var resolvedColorHex = ColorHelper.NormalizeHex(colorHex);
        if (style != ReportBorderStyle.None && resolvedColorHex.StartsWith("#00", StringComparison.Ordinal))
        {
            resolvedColorHex = "#FF000000";
        }

        return new ReportBorder
        {
            Style = style,
            ColorHex = resolvedColorHex,
            Width = style switch
            {
                ReportBorderStyle.Thick => 2.25d,
                ReportBorderStyle.Medium => 1.5d,
                ReportBorderStyle.DoubleLine => 0.75d,
                ReportBorderStyle.Hair => 0.25d,
                _ => 0.75d
            }
        };
    }

    private static ReportPageSetup ReadPageSetup(IXLWorksheet worksheet) =>
        new()
        {
            PaperSize = worksheet.PageSetup.PaperSize switch
            {
                XLPaperSize.LetterPaper => ReportPaperSize.Letter,
                XLPaperSize.LegalPaper => ReportPaperSize.Legal,
                _ => ReportPaperSize.A4
            },
            Orientation = worksheet.PageSetup.PageOrientation == XLPageOrientation.Landscape
                ? ReportPageOrientation.Landscape
                : ReportPageOrientation.Portrait,
            Margins = new()
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
        ArgumentNullException.ThrowIfNull(worksheet);

        var contentRange = worksheet.RangeUsed();
        var formattedRange = worksheet.RangeUsed(XLCellsUsedOptions.All);
        if (contentRange is null && formattedRange is null && worksheet.MergedRanges.Count == 0 && printArea is null)
        {
            range = default!;
            return false;
        }

        var startRow = int.MaxValue;
        var startColumn = int.MaxValue;
        var endRow = int.MinValue;
        var endColumn = int.MinValue;

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

        if (startRow == int.MaxValue || startColumn == int.MaxValue || endRow == int.MinValue || endColumn == int.MinValue)
        {
            range = default!;
            return false;
        }

        range = new ReportRange(startRow, startColumn, endRow, endColumn);
        return true;

        void IncludeRange(IXLRange? range)
        {
            if (range is null)
            {
                return;
            }

            startRow = Math.Min(startRow, range.RangeAddress.FirstAddress.RowNumber);
            startColumn = Math.Min(startColumn, range.RangeAddress.FirstAddress.ColumnNumber);
            endRow = Math.Max(endRow, range.RangeAddress.LastAddress.RowNumber);
            endColumn = Math.Max(endColumn, range.RangeAddress.LastAddress.ColumnNumber);
        }

        void IncludeReportRange(ReportRange range)
        {
            startRow = Math.Min(startRow, range.StartRow);
            startColumn = Math.Min(startColumn, range.StartColumn);
            endRow = Math.Max(endRow, range.EndRow);
            endColumn = Math.Max(endColumn, range.EndColumn);
        }
    }

    private static ReportImage ReadImage(IXLPicture picture)
    {
        using var memoryStream = new MemoryStream();
        picture.ImageStream.Position = 0;
        picture.ImageStream.CopyTo(memoryStream);
        var imageBytes = memoryStream.ToArray();
        var placement = picture.Placement switch
        {
            XLPicturePlacement.FreeFloating => ReportAnchorType.Absolute,
            XLPicturePlacement.Move => ReportAnchorType.MoveWithCells,
            _ => ReportAnchorType.MoveAndSizeWithCells
        };
        return new ReportImage(
            picture.Name,
            placement,
            picture.TopLeftCell.Address.ToStringRelative(false),
            TryGetBottomRightCellAddress(picture),
            new ReportOffset
            {
                X = picture.Left * 72d / 96d,
                Y = picture.Top * 72d / 96d
            },
            picture.Width * 72d / 96d,
            picture.Height * 72d / 96d,
            imageBytes);
    }

    private static string? TryGetBottomRightCellAddress(IXLPicture picture)
    {
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

    private static string ResolveFillColorHex(IXLFill fill, IXLWorkbook workbook)
    {
        ArgumentNullException.ThrowIfNull(fill);
        ArgumentNullException.ThrowIfNull(workbook);

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

    private static void ApplyTableStyles(
        ReportSheet reportSheet,
        IXLWorkbook workbook,
        IReadOnlyList<TableStyleInfo> tableStyles)
    {
        ArgumentNullException.ThrowIfNull(reportSheet);
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(tableStyles);

        foreach (var tableStyle in tableStyles)
        {
            if (!tableStyle.ShowRowStripes ||
                !string.Equals(tableStyle.ThemeName, "TableStyleLight4", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            const string stripeFillHex = "#FFDEEBF7";

            var firstDataRow = tableStyle.Range.StartRow + (tableStyle.HasHeaderRow ? 1 : 0);
            var lastDataRow = tableStyle.Range.EndRow - (tableStyle.HasTotalsRow ? 1 : 0);
            for (var rowIndex = firstDataRow; rowIndex <= lastDataRow; rowIndex++)
            {
                if (((rowIndex - firstDataRow) % 2) != 0)
                {
                    continue;
                }

                foreach (var cell in reportSheet.Cells.Where(cell =>
                             cell.Row == rowIndex &&
                             cell.Column >= tableStyle.Range.StartColumn &&
                             cell.Column <= tableStyle.Range.EndColumn))
                {
                    if (!IsTransparentFill(cell.Style.Fill.BackgroundColorHex))
                    {
                        continue;
                    }

                    cell.SetStyle(cell.Style with
                    {
                        Fill = cell.Style.Fill with
                        {
                            BackgroundColorHex = stripeFillHex
                        }
                    });
                }
            }
        }
    }

    private static double ConvertInchToPoint(double inch) => inch * 72d;

    private static bool IsTransparentFill(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);

    private static void ApplyMergedRanges(ReportSheet reportSheet)
    {
        foreach (var mergedRange in reportSheet.MergedRanges)
        {
            foreach (var cell in reportSheet.Cells.Where(cell => mergedRange.Range.Contains(cell.Row, cell.Column)))
            {
                cell.SetMerge(new ReportMergeInfo
                {
                    OwnerCellAddress = mergedRange.OwnerCellAddress,
                    IsOwner = string.Equals(cell.Address, mergedRange.OwnerCellAddress, StringComparison.Ordinal),
                    Range = mergedRange.Range
                });
            }
        }
    }

    private sealed record TableStyleInfo(
        ReportRange Range,
        string ThemeName,
        bool ShowRowStripes,
        bool HasHeaderRow,
        bool HasTotalsRow);

    private sealed class WorkbookRawColumnWidths
    {
        private readonly IReadOnlyList<IReadOnlyList<(int Min, int Max, double Width)>> sheetColumnRanges;

        private WorkbookRawColumnWidths(IReadOnlyList<IReadOnlyList<(int Min, int Max, double Width)>> sheetColumnRanges)
        {
            this.sheetColumnRanges = sheetColumnRanges;
        }

        public static WorkbookRawColumnWidths Load(byte[] workbookBytes)
        {
            ArgumentNullException.ThrowIfNull(workbookBytes);

            using var stream = new MemoryStream(workbookBytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            var mainNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var sheetRanges = archive.Entries
                .Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(entry => GetSheetOrder(entry.FullName))
                .Select(entry => LoadSheetColumnRanges(archive, entry.FullName, mainNamespace))
                .ToList();

            return new WorkbookRawColumnWidths(sheetRanges);
        }

        public double? TryGetRawWidth(int sheetIndex, int columnIndex)
        {
            if (sheetIndex < 0 || sheetIndex >= sheetColumnRanges.Count)
            {
                return null;
            }

            foreach (var (min, max, width) in sheetColumnRanges[sheetIndex])
            {
                if (columnIndex >= min && columnIndex <= max)
                {
                    return width;
                }
            }

            return null;
        }

        private static IReadOnlyList<(int Min, int Max, double Width)> LoadSheetColumnRanges(
            ZipArchive archive, string sheetPath, XNamespace mainNamespace)
        {
            var doc = LoadXml(archive, sheetPath);
            if (doc.Root is null)
            {
                return Array.Empty<(int, int, double)>();
            }

            var results = new List<(int Min, int Max, double Width)>();
            var colsElement = doc.Root.Element(mainNamespace + "cols");
            if (colsElement is null)
            {
                return results;
            }

            foreach (var col in colsElement.Elements(mainNamespace + "col"))
            {
                if (!double.TryParse(
                        col.Attribute("width")?.Value,
                        NumberStyles.Number,
                        CultureInfo.InvariantCulture,
                        out var width) || width <= 0)
                {
                    continue;
                }

                if (!int.TryParse(col.Attribute("min")?.Value, out var min) ||
                    !int.TryParse(col.Attribute("max")?.Value, out var max) ||
                    min <= 0 || max < min)
                {
                    continue;
                }

                results.Add((min, max, width));
            }

            return results;
        }

        private static int GetSheetOrder(string entryPath)
        {
            var fileName = Path.GetFileNameWithoutExtension(entryPath.Replace('/', Path.DirectorySeparatorChar));
            return fileName is not null && fileName.StartsWith("sheet", StringComparison.OrdinalIgnoreCase) &&
                   int.TryParse(fileName["sheet".Length..], out var order)
                ? order
                : int.MaxValue;
        }

        private static XDocument LoadXml(ZipArchive archive, string path)
        {
            var normalizedPath = path.Replace('\\', '/');
            var entry = archive.GetEntry(normalizedPath);
            if (entry is null)
            {
                return new XDocument();
            }

            using var entryStream = entry.Open();
            return XDocument.Load(entryStream);
        }
    }

    private sealed class WorkbookTableStyleMap
    {
        private readonly IReadOnlyList<IReadOnlyList<TableStyleInfo>> tableStylesBySheetIndex;

        private WorkbookTableStyleMap(IReadOnlyList<IReadOnlyList<TableStyleInfo>> tableStylesBySheetIndex)
        {
            this.tableStylesBySheetIndex = tableStylesBySheetIndex;
        }

        public static WorkbookTableStyleMap Load(byte[] workbookBytes)
        {
            ArgumentNullException.ThrowIfNull(workbookBytes);

            using var stream = new MemoryStream(workbookBytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            var mainNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var relationshipNamespace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            var tableStylesBySheetIndex = archive.Entries
                .Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(entry => GetSheetOrder(entry.FullName))
                .Select(entry => LoadSheetTableStyles(archive, entry.FullName, mainNamespace, relationshipNamespace))
                .ToList();

            return new WorkbookTableStyleMap(tableStylesBySheetIndex);
        }

        public IReadOnlyList<TableStyleInfo> GetTableStyles(int sheetIndex)
        {
            return sheetIndex >= 0 && sheetIndex < tableStylesBySheetIndex.Count
                ? tableStylesBySheetIndex[sheetIndex]
                : Array.Empty<TableStyleInfo>();
        }

        private static IReadOnlyList<TableStyleInfo> LoadSheetTableStyles(
            ZipArchive archive,
            string sheetPath,
            XNamespace mainNamespace,
            XNamespace relationshipNamespace)
        {
            var sheetDocument = LoadXml(archive, sheetPath);
            if (sheetDocument.Root is null)
            {
                return Array.Empty<TableStyleInfo>();
            }

            var tablePartElements = sheetDocument.Root
                .Elements(mainNamespace + "tableParts")
                .Elements(mainNamespace + "tablePart")
                .ToList();
            if (tablePartElements.Count == 0)
            {
                return Array.Empty<TableStyleInfo>();
            }

            var sheetRelationshipsPath = BuildRelationshipPath(sheetPath);
            var sheetRelationships = LoadRelationships(archive, sheetRelationshipsPath);
            var results = new List<TableStyleInfo>();
            foreach (var tablePartElement in tablePartElements)
            {
                var relationshipId = tablePartElement.Attribute(relationshipNamespace + "id")?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId) ||
                    !sheetRelationships.TryGetValue(relationshipId, out var tablePath))
                {
                    continue;
                }

                var tableDocument = LoadXml(archive, tablePath);
                var tableRoot = tableDocument.Root;
                if (tableRoot is null)
                {
                    continue;
                }

                var styleInfoElement = tableRoot.Element(mainNamespace + "tableStyleInfo");
                var rangeReference = tableRoot.Attribute("ref")?.Value;
                if (styleInfoElement is null || string.IsNullOrWhiteSpace(rangeReference))
                {
                    continue;
                }

                results.Add(new TableStyleInfo(
                    ParseRange(rangeReference),
                    styleInfoElement.Attribute("name")?.Value ?? string.Empty,
                    ReadBoolAttribute(styleInfoElement.Attribute("showRowStripes")),
                    !string.Equals(tableRoot.Attribute("headerRowCount")?.Value, "0", StringComparison.Ordinal),
                    ReadBoolAttribute(tableRoot.Attribute("totalsRowShown"))));
            }

            return results;
        }

        private static Dictionary<string, string> LoadRelationships(ZipArchive archive, string path)
        {
            var document = LoadXml(archive, path);
            var root = document.Root;
            if (root is null)
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var sourcePath = GetRelationshipSourcePath(path);
            return root.Elements()
                .Where(element => element.Attribute("Id") is not null && element.Attribute("Target") is not null)
                .ToDictionary(
                    element => element.Attribute("Id")!.Value,
                    element => ResolveZipPath(sourcePath, element.Attribute("Target")!.Value),
                    StringComparer.OrdinalIgnoreCase);
        }

        private static XDocument LoadXml(ZipArchive archive, string path)
        {
            var normalizedPath = path.Replace('\\', '/');
            var entry = archive.GetEntry(normalizedPath);
            if (entry is null)
            {
                return new XDocument();
            }

            using var entryStream = entry.Open();
            return XDocument.Load(entryStream);
        }

        private static string BuildRelationshipPath(string path)
        {
            var normalizedPath = path.Replace('\\', '/');
            var lastSlashIndex = normalizedPath.LastIndexOf('/');
            return lastSlashIndex < 0
                ? $"_rels/{normalizedPath}.rels"
                : $"{normalizedPath[..lastSlashIndex]}/_rels/{normalizedPath[(lastSlashIndex + 1)..]}.rels";
        }

        private static string GetRelationshipSourcePath(string relationshipPath)
        {
            var normalizedPath = relationshipPath.Replace('\\', '/');
            var marker = "/_rels/";
            var markerIndex = normalizedPath.IndexOf(marker, StringComparison.Ordinal);
            if (markerIndex < 0 || !normalizedPath.EndsWith(".rels", StringComparison.Ordinal))
            {
                return normalizedPath;
            }

            var prefix = normalizedPath[..markerIndex];
            var fileName = normalizedPath[(markerIndex + marker.Length)..^".rels".Length];
            return string.IsNullOrEmpty(prefix) ? fileName : $"{prefix}/{fileName}";
        }

        private static string ResolveZipPath(string sourcePath, string target)
        {
            if (target.StartsWith('/'))
            {
                return target.TrimStart('/');
            }

            var normalizedTarget = target.Replace('\\', '/');
            var normalizedSource = sourcePath.Replace('\\', '/');
            var lastSlash = normalizedSource.LastIndexOf('/');
            var baseDir = lastSlash >= 0 ? normalizedSource[..lastSlash] : string.Empty;
            var combined = string.IsNullOrEmpty(baseDir) ? normalizedTarget : $"{baseDir}/{normalizedTarget}";
            var parts = combined.Split('/');
            var resultParts = new List<string>();
            foreach (var part in parts)
            {
                if (part == "..")
                {
                    if (resultParts.Count > 0)
                    {
                        resultParts.RemoveAt(resultParts.Count - 1);
                    }
                }
                else if (part.Length > 0 && part != ".")
                {
                    resultParts.Add(part);
                }
            }

            return string.Join("/", resultParts);
        }

        private static int GetSheetOrder(string entryPath)
        {
            var fileName = Path.GetFileNameWithoutExtension(entryPath.Replace('/', Path.DirectorySeparatorChar));
            return fileName is not null && fileName.StartsWith("sheet", StringComparison.OrdinalIgnoreCase) &&
                   int.TryParse(fileName["sheet".Length..], out var order)
                ? order
                : int.MaxValue;
        }

        private static bool ReadBoolAttribute(XAttribute? attribute) =>
            string.Equals(attribute?.Value, "1", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(attribute?.Value, "true", StringComparison.OrdinalIgnoreCase);

        private static ReportRange ParseRange(string rangeReference)
        {
            var segments = rangeReference.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            var start = segments[0];
            var end = segments.Length > 1 ? segments[1] : segments[0];
            var (startRow, startColumn) = AddressHelper.ParseAddress(start);
            var (endRow, endColumn) = AddressHelper.ParseAddress(end);
            return new ReportRange(startRow, startColumn, endRow, endColumn);
        }
    }
}
