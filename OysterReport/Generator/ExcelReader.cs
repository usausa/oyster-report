namespace OysterReport.Generator;

using System.Globalization;
using System.Text;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using OysterReport.Helpers;

internal static class ExcelReader
{
    // Excel の列幅計算に使用する定数 (Excel 仕様に基づくピクセル変換パラメータ)
    private const double DefaultMaxDigitWidth = 7d;
    private const double ExcelColumnPaddingMultiplier = 2d;
    private const double ExcelColumnPaddingDivisor = 4d;
    private const double ExcelColumnPaddingOffsetPixels = 1d;
    private const double ExcelColumnWidthGranularity = 256d;
    private const double ExcelColumnWidthRoundingOffset = 128d;
    private const double PointsPerInch = 72d;
    private const double ScreenDpi = 96d;

    // フォント計測に使用する定数 (10pt 基準の1文字最大幅)
    private const double DefaultFallbackDigitWidth = 7d;
    private const double ReferenceFontSizePoints = 10d;
    private const double MeiryoDigitWidthAt10Pt = 8.28125d;
    private const double YuGothicDigitWidthAt10Pt = 8.5d;
    private const double MsPgothicDigitWidthAt10Pt = 6.66667d;
    private const double ArialDigitWidthAt10Pt = 7.41536d;

    /// <summary>ClosedXML ワークブックオブジェクトから ReportWorkbook を生成する。</summary>
    public static ReportWorkbook Read(IXLWorkbook workbook)
    {
        var measurementProfile = CreateMeasurementProfile(workbook);
        var metadata = new ReportMetadata { TemplateName = workbook.Properties.Title ?? "Workbook" };
        return ReadInternal(workbook, measurementProfile, metadata);
    }

    /// <summary>ストリームから Excel を読み込み ReportWorkbook を生成する。</summary>
    public static ReportWorkbook Read(Stream stream)
    {
        using var workbook = new XLWorkbook(stream);
        return Read(workbook);
    }

    /// <summary>ワークブック全体を読み込み、シートを列挙して ReportWorkbook を構築する。</summary>
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

    /// <summary>ワークブックの既定フォント情報から列幅計算用プロファイルを作成する。</summary>
    private static ReportMeasurementProfile CreateMeasurementProfile(IXLWorkbook workbook) =>
        new()
        {
            DefaultFontName = workbook.Style.Font.FontName,
            DefaultFontSize = workbook.Style.Font.FontSize,
            MaxDigitWidth = ResolveMaxDigitWidth(workbook.Style.Font.FontName, workbook.Style.Font.FontSize)
        };

    /// <summary>ワークシートを読み込み、行・列・セル・画像等を ReportSheet に変換する。</summary>
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
            reportSheet.AddColumnDefinition(new ReportColumn
            {
                Index = columnIndex,
                WidthPoint = ConvertExcelColumnWidthToPoint(column.Width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment),
                IsHidden = column.IsHidden,
                OutlineLevel = column.OutlineLevel,
                OriginalExcelWidth = column.Width
            });
        }

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

    /// <summary>セルの値を種別ごとに取得し ReportCellValue に変換する。</summary>
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

    /// <summary>セルのスタイル（フォント・塗り・罫線・配置）を ReportCellStyle に変換する。</summary>
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

    /// <summary>罫線スタイルと色を ReportBorder に変換する。透明色は黒に補正する。</summary>
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

    /// <summary>ページ設定（用紙・余白・中央揃え等）を ReportPageSetup に変換する。</summary>
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

    /// <summary>ヘッダー・フッターのテキストと表示条件を ReportHeaderFooter に変換する。</summary>
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

    /// <summary>印刷範囲が設定されていれば ReportPrintArea に変換する。未設定の場合は null を返す。</summary>
    private static ReportPrintArea? ReadPrintArea(IXLWorksheet worksheet)
    {
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

    /// <summary>シートの使用セル範囲・印刷範囲・マージ範囲を統合し、描画対象範囲を決定する。</summary>
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

        if (startRow == int.MaxValue || endRow == int.MinValue)
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

    /// <summary>ClosedXML の画像情報をポイント単位の座標に変換し ReportImage を生成する。</summary>
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

    /// <summary>MoveAndSize 配置の画像のみ右下セルアドレスを取得する。取得できない場合は null を返す。</summary>
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

    /// <summary>セル塗りつぶしの背景色を ARGB16 進文字列に変換する。パターン塗りはパターン色を優先する。</summary>
    private static string ResolveFillColorHex(IXLFill fill, IXLWorkbook workbook)
    {
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

    /// <summary>テーブルスタイル（縞模様等）をセルスタイルに適用する。現在は TableStyleLight4 の奇数行縞に対応。</summary>
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

    /// <summary>マージセル情報を各セルの Merge プロパティに設定する。</summary>
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

    /// <summary>
    /// Excel 列幅（文字数単位）をポイント値に変換する。
    /// Excel の列幅ピクセル計算仕様に従い、最大桁幅と画面 DPI を用いて算出する。
    /// </summary>
    private static double ConvertExcelColumnWidthToPoint(double excelWidth, double maxDigitWidth, double adjustment)
    {
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

    /// <summary>
    /// フォント名とサイズから最大桁幅（ポイント）を推定する。
    /// フォントごとの実測値を 10pt 基準で保持し、指定サイズに比例補正して返す。
    /// </summary>
    private static double ResolveMaxDigitWidth(string? fontName, double fontSize)
    {
        const double fallback = DefaultFallbackDigitWidth;
        if (string.IsNullOrWhiteSpace(fontName) || fontSize <= 0d)
        {
            return fallback;
        }

        var normalized = NormalizeFontName(fontName);
        var baseWidth = normalized switch
        {
            var v when v.Contains("meiryoui", StringComparison.Ordinal) => MeiryoDigitWidthAt10Pt,
            var v when v.Contains("meiryo", StringComparison.Ordinal) => MeiryoDigitWidthAt10Pt,
            var v when v.Contains("yugothicui", StringComparison.Ordinal) => YuGothicDigitWidthAt10Pt,
            var v when v.Contains("yugothic", StringComparison.Ordinal) => YuGothicDigitWidthAt10Pt,
            var v when v.Contains("mspgothic", StringComparison.Ordinal) => MsPgothicDigitWidthAt10Pt,
            var v when v.Contains("msgothic", StringComparison.Ordinal) => MsPgothicDigitWidthAt10Pt,
            var v when v.Contains("arial", StringComparison.Ordinal) => ArialDigitWidthAt10Pt,
            _ => fallback
        };

        return Math.Max(fallback, baseWidth * (fontSize / ReferenceFontSizePoints));
    }

    /// <summary>フォント名を正規化形式に変換し、英数字のみ小文字で返す。</summary>
    private static string NormalizeFontName(string fontName)
    {
        var normalized = fontName.Normalize(NormalizationForm.FormKC);
        var builder = new StringBuilder(normalized.Length);
        foreach (var character in normalized)
        {
            if (char.IsLetterOrDigit(character))
            {
                builder.Append(char.ToLowerInvariant(character));
            }
        }

        return builder.ToString();
    }
}
