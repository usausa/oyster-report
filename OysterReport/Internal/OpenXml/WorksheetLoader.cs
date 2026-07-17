namespace OysterReport.Internal.OpenXml;

using System.Globalization;
using System.Runtime.CompilerServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal sealed class WorksheetLoader
{
    private const double PointsPerInch = 72d;

    // Difference in days between the 1900 and 1904 Excel date systems
    private const double Date1904SerialOffset = 1462d;

    private static readonly ReportCellValue BlankCellValue = new() { Kind = CellValueKind.Blank, RawValue = string.Empty };

    private readonly StyleCatalog styles;
    private readonly string[] sharedStrings;
    private readonly ReportMeasurementProfile measurementProfile;
    private readonly bool date1904;

    private readonly Dictionary<int, ReportCellStyle> styleCache = [];

    public WorksheetLoader(StyleCatalog styles, string[] sharedStrings, ReportMeasurementProfile measurementProfile, bool date1904)
    {
        this.styles = styles;
        this.sharedStrings = sharedStrings;
        this.measurementProfile = measurementProfile;
        this.date1904 = date1904;
    }

    public ReportSheet Load(WorksheetPart part, string name, ReportPrintArea? printArea, IEnumerable<ReportTable> tables)
    {
        var sheet = new ReportSheet { Name = name };
        var rawCells = new List<RawCell>();
        var rowInfos = new Dictionary<int, RowInfo>();
        var columnInfos = new List<ColumnInfo>();
        var merges = new List<ReportRange>();
        var rowBreaks = new List<int>();
        var colBreaks = new List<int>();
        ReportPageSetup pageSetup = new();
        ReportHeaderFooter headerFooter = new();
        var showGridLines = true;
        var defaultRowHeight = 15d;

        using var reader = OpenXmlReader.Create(part);

        while (reader.Read())
        {
            if (!reader.IsStartElement)
            {
                continue;
            }

            var t = reader.ElementType;
            if (t == typeof(SheetFormatProperties))
            {
                var fmt = (SheetFormatProperties)reader.LoadCurrentElement()!;
                if (fmt.DefaultRowHeight is not null)
                {
                    defaultRowHeight = fmt.DefaultRowHeight.Value;
                }
            }
            else if (t == typeof(Columns))
            {
                var cols = (Columns)reader.LoadCurrentElement()!;
                foreach (var col in cols.Elements<Column>())
                {
                    var effectiveWidth = col.Width?.Value is { } w ? w - ColumnWidthOffset : 0d;
                    columnInfos.Add(new ColumnInfo(
                        (int)(col.Min?.Value ?? 1u),
                        (int)(col.Max?.Value ?? 1u),
                        effectiveWidth,
                        col.Hidden?.Value ?? false,
                        col.OutlineLevel?.Value ?? 0,
                        col.CustomWidth?.Value ?? false,
                        (int)(col.Style?.Value ?? 0u)));
                }
            }
            else if (t == typeof(Row))
            {
                var row = (Row)reader.LoadCurrentElement()!;
                ProcessRow(row, rawCells, rowInfos, defaultRowHeight, columnInfos);
            }
            else if (t == typeof(MergeCells))
            {
                var m = (MergeCells)reader.LoadCurrentElement()!;
                foreach (var merge in m.Elements<MergeCell>())
                {
                    if (merge.Reference?.Value is { } refStr)
                    {
                        merges.Add(ParseRangeRef(refStr));
                    }
                }
            }
            else if (t == typeof(PageSetup))
            {
                var ps = (PageSetup)reader.LoadCurrentElement()!;
                var wide = ps.FitToWidth?.Value ?? 0u;
                var tall = ps.FitToHeight?.Value ?? 0u;
                pageSetup = pageSetup with
                {
                    PaperSize = EnumMaps.ToPaperSize(ps.PaperSize?.Value),
                    Orientation = EnumMaps.ToPageOrientation(ps.Orientation?.Value),
                    ScalePercent = (int)(ps.Scale?.Value ?? 100u),
                    FitToPagesWide = wide == 0u ? null : (int)wide,
                    FitToPagesTall = tall == 0u ? null : (int)tall
                };
            }
            else if (t == typeof(PageMargins))
            {
                var pm = (PageMargins)reader.LoadCurrentElement()!;
                pageSetup = pageSetup with
                {
                    Margins = new ReportThickness
                    {
                        Left = InchToPoint(pm.Left?.Value ?? 0d),
                        Top = InchToPoint(pm.Top?.Value ?? 0d),
                        Right = InchToPoint(pm.Right?.Value ?? 0d),
                        Bottom = InchToPoint(pm.Bottom?.Value ?? 0d)
                    },
                    HeaderMarginPoint = InchToPoint(pm.Header?.Value ?? 0d),
                    FooterMarginPoint = InchToPoint(pm.Footer?.Value ?? 0d)
                };
            }
            else if (t == typeof(PrintOptions))
            {
                var po = (PrintOptions)reader.LoadCurrentElement()!;
                pageSetup = pageSetup with
                {
                    CenterHorizontally = po.HorizontalCentered?.Value ?? false,
                    CenterVertically = po.VerticalCentered?.Value ?? false
                };
            }
            else if (t == typeof(HeaderFooter))
            {
                var hf = (HeaderFooter)reader.LoadCurrentElement()!;
                headerFooter = new ReportHeaderFooter
                {
                    AlignWithMargins = hf.AlignWithMargins?.Value ?? true,
                    DifferentFirst = hf.DifferentFirst?.Value ?? false,
                    DifferentOddEven = hf.DifferentOddEven?.Value ?? false,
                    ScaleWithDocument = hf.ScaleWithDoc?.Value ?? true,
                    OddHeader = hf.OddHeader?.Text,
                    OddFooter = hf.OddFooter?.Text,
                    EvenHeader = hf.EvenHeader?.Text,
                    EvenFooter = hf.EvenFooter?.Text,
                    FirstHeader = hf.FirstHeader?.Text,
                    FirstFooter = hf.FirstFooter?.Text
                };
            }
            else if (t == typeof(SheetView))
            {
                var view = (SheetView)reader.LoadCurrentElement()!;
                showGridLines = view.ShowGridLines?.Value ?? true;
            }
            else if (t == typeof(RowBreaks))
            {
                var rb = (RowBreaks)reader.LoadCurrentElement()!;
                foreach (var br in rb.Elements<Break>())
                {
                    if (br.Id?.Value is { } idx)
                    {
                        rowBreaks.Add((int)idx);
                    }
                }
            }
            else if (t == typeof(ColumnBreaks))
            {
                var cb = (ColumnBreaks)reader.LoadCurrentElement()!;
                foreach (var br in cb.Elements<Break>())
                {
                    if (br.Id?.Value is { } idx)
                    {
                        colBreaks.Add((int)idx);
                    }
                }
            }
        }

        sheet.PageSetup = pageSetup;
        sheet.HeaderFooter = headerFooter;
        sheet.PrintArea = printArea;
        sheet.ShowGridLines = showGridLines;

        foreach (var table in tables)
        {
            sheet.AddTable(table);
        }

        AssembleSheet(sheet, rawCells, rowInfos, columnInfos, merges, rowBreaks, colBreaks, defaultRowHeight);

        return sheet;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    private void ProcessRow(
        Row row,
        List<RawCell> rawCells,
        Dictionary<int, RowInfo> rowInfos,
        double defaultRowHeight,
        List<ColumnInfo> columnInfos)
    {
        if (row.RowIndex?.Value is not { } rowIndex)
        {
            return;
        }

        var height = row.Height?.Value ?? defaultRowHeight;
        rowInfos[(int)rowIndex] = new RowInfo(
            (int)rowIndex,
            height,
            row.Hidden?.Value ?? false,
            row.OutlineLevel?.Value ?? 0);

        foreach (var c in row.Elements<Cell>())
        {
            var raw = ParseCell(c, columnInfos);
            if (raw is not null)
            {
                rawCells.Add(raw);
            }
        }
    }

    private RawCell? ParseCell(Cell cell, List<ColumnInfo> columnInfos)
    {
        if (cell.CellReference?.Value is not { } addr)
        {
            return null;
        }

        AddressHelper.ParseAddress(addr, out var row, out var col);
        var styleIndex = cell.StyleIndex?.Value is { } s ? (int)s : ResolveColumnStyleIndex(columnInfos, col);
        var type = cell.DataType?.Value;
        string? rawValue = null;
        object? typedValue = null;
        var kind = CellValueKind.Blank;

        if (type == CellValues.SharedString)
        {
            if (cell.CellValue?.Text is { } idxText && Int32.TryParse(idxText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx) &&
                idx >= 0 && idx < sharedStrings.Length)
            {
                rawValue = sharedStrings[idx];
                typedValue = rawValue;
                kind = CellValueKind.Text;
            }
        }
        else if (type == CellValues.InlineString)
        {
            rawValue = cell.InlineString is { } inlineString ? OpenXmlText.Extract(inlineString) : string.Empty;
            typedValue = rawValue;
            kind = CellValueKind.Text;
        }
        else if (type == CellValues.String)
        {
            rawValue = cell.CellValue?.Text ?? string.Empty;
            typedValue = rawValue;
            kind = CellValueKind.Text;
        }
        else if (type == CellValues.Boolean)
        {
            var v = cell.CellValue?.Text;
            typedValue = v == "1";
            rawValue = typedValue.ToString();
            kind = CellValueKind.Boolean;
        }
        else if (type == CellValues.Error)
        {
            rawValue = cell.CellValue?.Text ?? string.Empty;
            typedValue = rawValue;
            kind = CellValueKind.Error;
        }
        else
        {
            if (cell.CellValue?.Text is { Length: > 0 } num &&
                Double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            {
                typedValue = d;
                rawValue = num;
                var xf = GetXf(styleIndex);
                var fmtCode = styles.ResolveNumberFormat(xf.NumFmtId);
                if (ExcelFormatCode.IsDateTime(fmtCode))
                {
                    kind = CellValueKind.DateTime;
                    try
                    {
                        typedValue = DateTime.FromOADate(date1904 ? d + Date1904SerialOffset : d);
                    }
                    catch (ArgumentException)
                    {
                        typedValue = d;
                        kind = CellValueKind.Number;
                    }
                }
                else
                {
                    kind = CellValueKind.Number;
                }
            }
        }

        if (rawValue is null && cell.CellValue is null)
        {
            return new RawCell(row, col, styleIndex, CellValueKind.Blank, string.Empty, string.Empty);
        }

        var displayText = ComputeDisplayText(typedValue, kind, styleIndex);
        return new RawCell(row, col, styleIndex, kind, typedValue, displayText);
    }

    private string ComputeDisplayText(object? value, CellValueKind kind, int styleIndex)
    {
        if (value is null)
        {
            return string.Empty;
        }

        var xf = GetXf(styleIndex);
        var fmtCode = styles.ResolveNumberFormat(xf.NumFmtId);

        return kind switch
        {
            CellValueKind.Number when value is double d => ExcelFormatCode.Format(d, fmtCode),
            CellValueKind.DateTime when value is DateTime dt => ExcelFormatCode.Format(dt.ToOADate(), fmtCode),
            CellValueKind.Boolean when value is bool b => b ? "TRUE" : "FALSE",
            _ => value.ToString() ?? string.Empty
        };
    }

    private CellXfEntry GetXf(int index)
    {
        if ((uint)index >= (uint)styles.CellXfs.Length)
        {
            return new CellXfEntry(0, 0, 0, 0, false, false, false, HorizontalAlignment.General, VerticalAlignment.Bottom, false);
        }
        return styles.CellXfs[index];
    }

    private void AssembleSheet(
        ReportSheet sheet,
        List<RawCell> rawCells,
        Dictionary<int, RowInfo> rowInfos,
        List<ColumnInfo> columnInfos,
        List<ReportRange> merges,
        List<int> rowBreaks,
        List<int> colBreaks,
        double defaultRowHeight)
    {
        var range = ResolveRange(rawCells, merges, sheet.PrintArea);
        if (range is null)
        {
            return;
        }
        sheet.UsedRange = range.Value;
        sheet.DefaultRowHeightPoint = defaultRowHeight;

        var colWidthByIndex = new Dictionary<int, (double Width, bool Hidden, int Outline)>();
        foreach (var info in columnInfos)
        {
            for (var c = info.Min; c <= info.Max; c++)
            {
                colWidthByIndex[c] = (info.Width, info.Hidden, info.OutlineLevel);
            }
        }

        for (var c = range.Value.StartColumn; c <= range.Value.EndColumn; c++)
        {
            var width = colWidthByIndex.TryGetValue(c, out var ci) ? ci.Width : 8.43d;
            var hidden = colWidthByIndex.TryGetValue(c, out var ci2) && ci2.Hidden;
            var outline = colWidthByIndex.TryGetValue(c, out var ci3) ? ci3.Outline : 0;
            sheet.AddColumnDefinition(new ReportColumn
            {
                Index = c,
                WidthPoint = ConvertExcelColumnWidthToPoint(width, measurementProfile.MaxDigitWidth, measurementProfile.ColumnWidthAdjustment),
                IsHidden = hidden,
                OutlineLevel = outline,
                OriginalExcelWidth = width
            });
        }

        foreach (var merge in merges)
        {
            sheet.AddMergedRange(new ReportMergedRange { Range = merge });
        }

        var startColumn = range.Value.StartColumn;
        var blankStylesByColumn = new ReportCellStyle[range.Value.EndColumn - startColumn + 1];
        for (var i = 0; i < blankStylesByColumn.Length; i++)
        {
            blankStylesByColumn[i] = GetOrBuildStyle(ResolveColumnStyleIndex(columnInfos, startColumn + i));
        }

        var cellLookup = rawCells.ToDictionary(rc => (rc.Row, rc.Column));
        for (var r = range.Value.StartRow; r <= range.Value.EndRow; r++)
        {
            for (var c = range.Value.StartColumn; c <= range.Value.EndColumn; c++)
            {
                if (cellLookup.TryGetValue((r, c), out var raw))
                {
                    sheet.AddCell(BuildCell(raw));
                }
                else
                {
                    var blankStyle = blankStylesByColumn[c - startColumn];
                    if (!ShouldMaterializeBlankCell(blankStyle, r, c, merges, sheet.Tables))
                    {
                        continue;
                    }

                    sheet.AddCell(new ReportCell
                    {
                        Row = r,
                        Column = c,
                        Value = BlankCellValue,
                        DisplayText = string.Empty,
                        Style = blankStyle
                    });
                }
            }
        }

        var rowsWithCells = new HashSet<int>();
        foreach (var cell in sheet.Cells)
        {
            rowsWithCells.Add(cell.Row);
        }

        for (var r = range.Value.StartRow; r <= range.Value.EndRow; r++)
        {
            if (rowInfos.TryGetValue(r, out var info))
            {
                sheet.AddRowDefinition(new ReportRow
                {
                    Index = r,
                    HeightPoint = info.Height,
                    IsHidden = info.Hidden,
                    OutlineLevel = info.OutlineLevel
                });
            }
            else if (rowsWithCells.Contains(r))
            {
                sheet.AddRowDefinition(new ReportRow
                {
                    Index = r,
                    HeightPoint = defaultRowHeight,
                    IsHidden = false,
                    OutlineLevel = 0
                });
            }
        }

        foreach (var br in rowBreaks)
        {
            sheet.AddHorizontalPageBreak(new ReportPageBreak { Index = br });
        }
        foreach (var br in colBreaks)
        {
            sheet.AddVerticalPageBreak(new ReportPageBreak { Index = br });
        }

        sheet.RecalculateLayout();
        ApplyMerges(sheet);
    }

    private static bool ShouldMaterializeBlankCell(ReportCellStyle style, int row, int column, List<ReportRange> merges, IEnumerable<ReportTable> tables)
    {
        if (HasVisibleStyle(style))
        {
            return true;
        }

        foreach (var merge in merges)
        {
            if (merge.Contains(row, column))
            {
                return true;
            }
        }

        foreach (var table in tables)
        {
            if (table.Range.Contains(row, column))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleStyle(ReportCellStyle style) =>
        !style.Fill.BackgroundColorHex.StartsWith("#00", StringComparison.Ordinal) ||
        (style.Borders.Left.Style != BorderLineStyle.None) ||
        (style.Borders.Top.Style != BorderLineStyle.None) ||
        (style.Borders.Right.Style != BorderLineStyle.None) ||
        (style.Borders.Bottom.Style != BorderLineStyle.None);

    private ReportCell BuildCell(RawCell raw)
    {
        return new ReportCell
        {
            Row = raw.Row,
            Column = raw.Column,
            Value = raw.Kind == CellValueKind.Blank
                ? BlankCellValue
                : new ReportCellValue { Kind = raw.Kind, RawValue = raw.TypedValue },
            DisplayText = raw.DisplayText,
            Style = GetOrBuildStyle(raw.StyleIndex)
        };
    }

    private ReportCellStyle GetOrBuildStyle(int styleIndex)
    {
        if (!styleCache.TryGetValue(styleIndex, out var style))
        {
            style = BuildStyle(styleIndex);
            styleCache[styleIndex] = style;
        }

        return style;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    private ReportCellStyle BuildStyle(int styleIndex)
    {
        var xf = GetXf(styleIndex);
        var font = ((uint)xf.FontId < (uint)styles.Fonts.Length)
            ? styles.Fonts[xf.FontId]
            : new FontEntry("Calibri", 11d, false, false, false, false, "#FF000000");
        var fill = ((uint)xf.FillId < (uint)styles.Fills.Length)
            ? styles.Fills[xf.FillId]
            : new FillEntry(FillPattern.None, "#00000000", "#00000000", null, null, 0);
        var border = ((uint)xf.BorderId < (uint)styles.Borders.Length)
            ? styles.Borders[xf.BorderId]
            : new BorderEntry(
                BorderLineStyle.None,
                null,
                BorderLineStyle.None,
                null,
                BorderLineStyle.None,
                null,
                BorderLineStyle.None,
                null);

        return new ReportCellStyle
        {
            Font = new ReportFont
            {
                Name = font.Name,
                Size = font.Size,
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = font.Underline,
                Strikeout = font.Strike,
                ColorHex = font.ColorHex
            },
            Fill = new ReportFill
            {
                BackgroundColorHex = ResolveFillColor(fill)
            },
            Borders = new ReportBorders
            {
                Left = BuildBorder(border.LeftStyle, border.LeftColor),
                Top = BuildBorder(border.TopStyle, border.TopColor),
                Right = BuildBorder(border.RightStyle, border.RightColor),
                Bottom = BuildBorder(border.BottomStyle, border.BottomColor)
            },
            Alignment = new ReportAlignment
            {
                Horizontal = xf.Horizontal,
                Vertical = xf.Vertical
            },
            WrapText = xf.WrapText
        };
    }

    private string ResolveFillColor(FillEntry fill)
    {
        if (fill.Pattern == FillPattern.None)
        {
            return "#00000000";
        }

        var fg = styles.ColorResolver.Resolve(fill.RawFg, "#00000000");
        if (!fg.StartsWith("#00", StringComparison.Ordinal))
        {
            return fg;
        }
        return styles.ColorResolver.Resolve(fill.RawBg, "#00000000");
    }

    private ReportBorder BuildBorder(BorderLineStyle style, DocumentFormat.OpenXml.Spreadsheet.ColorType? color)
    {
        var colorHex = styles.ColorResolver.Resolve(color, "#FF000000");
        if ((style != BorderLineStyle.None) && colorHex.StartsWith("#00", StringComparison.Ordinal))
        {
            colorHex = "#FF000000";
        }

        return new ReportBorder
        {
            Style = style,
            ColorHex = colorHex,
            Width = ResolveBorderWidth(style)
        };
    }

    private static double ResolveBorderWidth(BorderLineStyle style) => style switch
    {
        BorderLineStyle.Thick => 2.25d,
        BorderLineStyle.Medium => 1.5d,
        BorderLineStyle.Hair => 0.25d,
        _ => 0.75d
    };

    private static ReportRange? ResolveRange(List<RawCell> cells, List<ReportRange> merges, ReportPrintArea? printArea)
    {
        if (cells.Count == 0 && merges.Count == 0 && printArea is null)
        {
            return null;
        }

        var startRow = Int32.MaxValue;
        var startCol = Int32.MaxValue;
        var endRow = Int32.MinValue;
        var endCol = Int32.MinValue;

        foreach (var c in cells)
        {
            startRow = Math.Min(startRow, c.Row);
            startCol = Math.Min(startCol, c.Column);
            endRow = Math.Max(endRow, c.Row);
            endCol = Math.Max(endCol, c.Column);
        }
        foreach (var m in merges)
        {
            startRow = Math.Min(startRow, m.StartRow);
            startCol = Math.Min(startCol, m.StartColumn);
            endRow = Math.Max(endRow, m.EndRow);
            endCol = Math.Max(endCol, m.EndColumn);
        }
        if (printArea is not null)
        {
            var r = printArea.Range;
            startRow = Math.Min(startRow, r.StartRow);
            startCol = Math.Min(startCol, r.StartColumn);
            endRow = Math.Max(endRow, r.EndRow);
            endCol = Math.Max(endCol, r.EndColumn);
        }

        if (startRow == Int32.MaxValue || endRow == Int32.MinValue)
        {
            return null;
        }

        return new ReportRange { StartRow = startRow, StartColumn = startCol, EndRow = endRow, EndColumn = endCol };
    }

    private static ReportRange ParseRangeRef(string reference)
    {
        var colonIdx = reference.IndexOf(':', StringComparison.Ordinal);
        if (colonIdx < 0)
        {
            AddressHelper.ParseAddress(reference, out var row, out var col);
            return new ReportRange { StartRow = row, StartColumn = col, EndRow = row, EndColumn = col };
        }

        AddressHelper.ParseAddress(reference[..colonIdx], out var r1, out var c1);
        AddressHelper.ParseAddress(reference[(colonIdx + 1)..], out var r2, out var c2);
        return new ReportRange
        {
            StartRow = Math.Min(r1, r2),
            StartColumn = Math.Min(c1, c2),
            EndRow = Math.Max(r1, r2),
            EndColumn = Math.Max(c1, c2)
        };
    }

    private static void ApplyMerges(ReportSheet sheet)
    {
        foreach (var merge in sheet.MergedRanges)
        {
            var info = new ReportMergeInfo
            {
                OwnerCellAddress = merge.OwnerCellAddress,
                Range = merge.Range
            };

            for (var r = merge.Range.StartRow; r <= merge.Range.EndRow; r++)
            {
                for (var c = merge.Range.StartColumn; c <= merge.Range.EndColumn; c++)
                {
                    var cell = sheet.FindCell(r, c);
                    cell?.Merge = info;
                }
            }
        }
    }

    private static double InchToPoint(double inch) => inch * PointsPerInch;

    private const double ColumnWidthOffset = 0.710625d;
    private const double DefaultMaxDigitWidth = 7d;
    private const double ExcelColumnPaddingMultiplier = 2d;
    private const double ExcelColumnPaddingDivisor = 4d;
    private const double ExcelColumnPaddingOffsetPixels = 1d;
    private const double ExcelColumnWidthGranularity = 256d;
    private const double ExcelColumnWidthRoundingOffset = 128d;
    private const double ScreenDpi = 96d;

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
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

    private static int ResolveColumnStyleIndex(List<ColumnInfo> columnInfos, int column)
    {
        foreach (var info in columnInfos)
        {
            if (column >= info.Min && column <= info.Max)
            {
                return info.StyleIndex;
            }
        }
        return 0;
    }

    // ReSharper disable NotAccessedPositionalProperty.Local
    private sealed record RawCell(int Row, int Column, int StyleIndex, CellValueKind Kind, object? TypedValue, string DisplayText);

    private sealed record RowInfo(int Index, double Height, bool Hidden, int OutlineLevel);

    private sealed record ColumnInfo(int Min, int Max, double Width, bool Hidden, int OutlineLevel, bool CustomWidth, int StyleIndex);
    // ReSharper restore NotAccessedPositionalProperty.Local
}
