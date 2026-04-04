namespace OysterReport.Reading;

using System.Globalization;
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
            SourceLastWriteTime = File.Exists(filePath) ? File.GetLastWriteTimeUtc(filePath) : null,
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

        using var workbook = new XLWorkbook(stream);
        var measurementProfile = new ReportMeasurementProfile
        {
            DefaultFontName = workbook.Style.Font.FontName,
            DefaultFontSize = workbook.Style.Font.FontSize,
        };
        var reportWorkbook = new ReportWorkbook(
            metadata ?? new ReportMetadata { TemplateName = workbook.Properties.Title ?? "Workbook" },
            measurementProfile);
        var targetSheets = options?.TargetSheets is { Count: > 0 }
            ? new HashSet<string>(options.TargetSheets, sheetNameComparer)
            : null;

        foreach (var worksheet in workbook.Worksheets)
        {
            if (targetSheets is not null && !targetSheets.Contains(worksheet.Name))
            {
                continue;
            }

            reportWorkbook.AddSheet(ReadSheet(worksheet, measurementProfile, options?.IncludeImages ?? true));
        }

        return reportWorkbook;
    }

    private static ReportSheet ReadSheet(IXLWorksheet worksheet, ReportMeasurementProfile measurementProfile, bool includeImages)
    {
        var reportSheet = new ReportSheet(worksheet.Name);
        var usedRange = worksheet.RangeUsed();
        if (usedRange is null)
        {
            return reportSheet;
        }

        var range = new ReportRange(
            usedRange.RangeAddress.FirstAddress.RowNumber,
            usedRange.RangeAddress.FirstAddress.ColumnNumber,
            usedRange.RangeAddress.LastAddress.RowNumber,
            usedRange.RangeAddress.LastAddress.ColumnNumber);

        reportSheet.SetUsedRange(range);
        reportSheet.SetPageSetup(ReadPageSetup(worksheet));
        reportSheet.SetHeaderFooter(ReadHeaderFooter(worksheet));
        reportSheet.SetPrintArea(ReadPrintArea(worksheet));
        reportSheet.SetShowGridLines(worksheet.PageSetup.ShowGridlines);

        for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
        {
            var row = worksheet.Row(rowIndex);
            reportSheet.AddRowDefinition(new ReportRow(rowIndex, row.Height, row.IsHidden, row.OutlineLevel));
        }

        for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
        {
            var column = worksheet.Column(columnIndex);
            var widthPoint = ColumnWidthConverter.ToPoint(column.Width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment);
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
            _ => new ReportCellValue { Kind = ReportCellValueKind.Blank, RawValue = cell.Value.ToString(CultureInfo.InvariantCulture) },
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
                ColorHex = ColorHelper.NormalizeHex(style.Font.FontColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
            },
            Fill = new ReportFill
            {
                BackgroundColorHex = ColorHelper.NormalizeHex(style.Fill.BackgroundColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
            },
            Borders = new ReportBorders
            {
                Left = ReadBorder(style.Border.LeftBorder, style.Border.LeftBorderColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
                Top = ReadBorder(style.Border.TopBorder, style.Border.TopBorderColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
                Right = ReadBorder(style.Border.RightBorder, style.Border.RightBorderColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
                Bottom = ReadBorder(style.Border.BottomBorder, style.Border.BottomBorderColor.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)),
            },
            Alignment = new ReportAlignment
            {
                Horizontal = style.Alignment.Horizontal switch
                {
                    XLAlignmentHorizontalValues.Center => ReportHorizontalAlignment.Center,
                    XLAlignmentHorizontalValues.Right => ReportHorizontalAlignment.Right,
                    XLAlignmentHorizontalValues.Justify => ReportHorizontalAlignment.Justify,
                    _ => ReportHorizontalAlignment.Left,
                },
                Vertical = style.Alignment.Vertical switch
                {
                    XLAlignmentVerticalValues.Center => ReportVerticalAlignment.Center,
                    XLAlignmentVerticalValues.Bottom => ReportVerticalAlignment.Bottom,
                    XLAlignmentVerticalValues.Justify => ReportVerticalAlignment.Justify,
                    _ => ReportVerticalAlignment.Top,
                },
            },
            NumberFormat = style.NumberFormat.Format,
            WrapText = style.Alignment.WrapText,
            Rotation = style.Alignment.TextRotation,
            ShrinkToFit = style.Alignment.ShrinkToFit,
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
            _ => ReportBorderStyle.Thin,
        };

        return new ReportBorder
        {
            Style = style,
            ColorHex = ColorHelper.NormalizeHex(colorHex),
            Width = style switch
            {
                ReportBorderStyle.Thick => 2d,
                ReportBorderStyle.Medium => 1d,
                ReportBorderStyle.DoubleLine => 1.5d,
                _ => 0.5d,
            },
        };
    }

    private static ReportPageSetup ReadPageSetup(IXLWorksheet worksheet) =>
        new()
        {
            PaperSize = worksheet.PageSetup.PaperSize switch
            {
                XLPaperSize.LetterPaper => ReportPaperSize.Letter,
                XLPaperSize.LegalPaper => ReportPaperSize.Legal,
                _ => ReportPaperSize.A4,
            },
            Orientation = worksheet.PageSetup.PageOrientation == XLPageOrientation.Landscape
                ? ReportPageOrientation.Landscape
                : ReportPageOrientation.Portrait,
            Margins = new()
            {
                Left = worksheet.PageSetup.Margins.Left,
                Top = worksheet.PageSetup.Margins.Top,
                Right = worksheet.PageSetup.Margins.Right,
                Bottom = worksheet.PageSetup.Margins.Bottom,
            },
            HeaderMarginPoint = worksheet.PageSetup.Margins.Header,
            FooterMarginPoint = worksheet.PageSetup.Margins.Footer,
            ScalePercent = worksheet.PageSetup.Scale,
            FitToPagesWide = worksheet.PageSetup.PagesWide == 0 ? null : worksheet.PageSetup.PagesWide,
            FitToPagesTall = worksheet.PageSetup.PagesTall == 0 ? null : worksheet.PageSetup.PagesTall,
            CenterHorizontally = worksheet.PageSetup.CenterHorizontally,
            CenterVertically = worksheet.PageSetup.CenterVertically,
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
            FirstFooter = worksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage),
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
                printArea.RangeAddress.LastAddress.ColumnNumber),
        };
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
            _ => ReportAnchorType.MoveAndSizeWithCells,
        };
        return new ReportImage(
            picture.Name,
            placement,
            picture.TopLeftCell.Address.ToStringRelative(false),
            picture.BottomRightCell?.Address.ToStringRelative(false),
            new ReportOffset
            {
                X = picture.Left * 72d / 96d,
                Y = picture.Top * 72d / 96d,
            },
            picture.Width * 72d / 96d,
            picture.Height * 72d / 96d,
            imageBytes);
    }

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
                    Range = mergedRange.Range,
                });
            }
        }
    }
}
