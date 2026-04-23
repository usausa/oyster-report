// Canonical deep-diff between two ReportWorkbook instances.
// Used to verify that the prototype OpenXmlLoader produces output indistinguishable from the ClosedXML path,
// since PdfGenerator is deterministic w.r.t. its ReportWorkbook input.

namespace OysterReport.Prototype;

using System.Globalization;
using System.Text;

using OysterReport.Internal;

internal static class WorkbookDiff
{
    public static IReadOnlyList<string> Compare(ReportWorkbook expected, ReportWorkbook actual)
    {
        var diffs = new List<string>();

        if (expected.Metadata.TemplateName != actual.Metadata.TemplateName)
        {
            diffs.Add($"Metadata.TemplateName: expected='{expected.Metadata.TemplateName}' actual='{actual.Metadata.TemplateName}'");
        }

        if (Math.Abs(expected.MeasurementProfile.MaxDigitWidth - actual.MeasurementProfile.MaxDigitWidth) > 1e-6)
        {
            diffs.Add($"MeasurementProfile.MaxDigitWidth: expected={expected.MeasurementProfile.MaxDigitWidth} actual={actual.MeasurementProfile.MaxDigitWidth}");
        }
        if (Math.Abs(expected.MeasurementProfile.ColumnWidthAdjustment - actual.MeasurementProfile.ColumnWidthAdjustment) > 1e-6)
        {
            diffs.Add($"MeasurementProfile.ColumnWidthAdjustment: expected={expected.MeasurementProfile.ColumnWidthAdjustment} actual={actual.MeasurementProfile.ColumnWidthAdjustment}");
        }

        if (expected.Sheets.Count != actual.Sheets.Count)
        {
            diffs.Add($"Sheets.Count: expected={expected.Sheets.Count} actual={actual.Sheets.Count}");
            return diffs;
        }

        for (var i = 0; i < expected.Sheets.Count; i++)
        {
            CompareSheet($"Sheets[{i}]", expected.Sheets[i], actual.Sheets[i], diffs);
        }

        return diffs;
    }

    private static void CompareSheet(string path, ReportSheet a, ReportSheet b, List<string> diffs)
    {
        if (a.Name != b.Name)
        {
            diffs.Add($"{path}.Name: expected='{a.Name}' actual='{b.Name}'");
        }
        if (!a.UsedRange.Equals(b.UsedRange))
        {
            diffs.Add($"{path}.UsedRange: expected={a.UsedRange} actual={b.UsedRange}");
        }
        if (a.ShowGridLines != b.ShowGridLines)
        {
            diffs.Add($"{path}.ShowGridLines: expected={a.ShowGridLines} actual={b.ShowGridLines}");
        }

        ComparePageSetup($"{path}.PageSetup", a.PageSetup, b.PageSetup, diffs);
        CompareHeaderFooter($"{path}.HeaderFooter", a.HeaderFooter, b.HeaderFooter, diffs);
        ComparePrintArea($"{path}.PrintArea", a.PrintArea, b.PrintArea, diffs);

        CompareRows($"{path}.Rows", a.Rows, b.Rows, diffs);
        CompareColumns($"{path}.Columns", a.Columns, b.Columns, diffs);
        CompareMerges($"{path}.MergedRanges", a.MergedRanges, b.MergedRanges, diffs);
        CompareImages($"{path}.Images", a.Images, b.Images, diffs);
        CompareBreaks($"{path}.HorizontalPageBreaks", a.HorizontalPageBreaks, b.HorizontalPageBreaks, diffs);
        CompareBreaks($"{path}.VerticalPageBreaks", a.VerticalPageBreaks, b.VerticalPageBreaks, diffs);
        CompareCells($"{path}.Cells", a.Cells, b.Cells, diffs);
    }

    private static void ComparePageSetup(string path, ReportPageSetup a, ReportPageSetup b, List<string> diffs)
    {
        if (a.PaperSize != b.PaperSize)
        {
            diffs.Add($"{path}.PaperSize: expected={a.PaperSize} actual={b.PaperSize}");
        }
        if (a.Orientation != b.Orientation)
        {
            diffs.Add($"{path}.Orientation: expected={a.Orientation} actual={b.Orientation}");
        }
        if (!ApproxEqual(a.Margins.Left, b.Margins.Left) ||
            !ApproxEqual(a.Margins.Top, b.Margins.Top) ||
            !ApproxEqual(a.Margins.Right, b.Margins.Right) ||
            !ApproxEqual(a.Margins.Bottom, b.Margins.Bottom))
        {
            diffs.Add($"{path}.Margins: expected={FormatThickness(a.Margins)} actual={FormatThickness(b.Margins)}");
        }
        if (!ApproxEqual(a.HeaderMarginPoint, b.HeaderMarginPoint))
        {
            diffs.Add($"{path}.HeaderMarginPoint: expected={a.HeaderMarginPoint} actual={b.HeaderMarginPoint}");
        }
        if (!ApproxEqual(a.FooterMarginPoint, b.FooterMarginPoint))
        {
            diffs.Add($"{path}.FooterMarginPoint: expected={a.FooterMarginPoint} actual={b.FooterMarginPoint}");
        }
        if (a.ScalePercent != b.ScalePercent)
        {
            diffs.Add($"{path}.ScalePercent: expected={a.ScalePercent} actual={b.ScalePercent}");
        }
        if (a.FitToPagesWide != b.FitToPagesWide)
        {
            diffs.Add($"{path}.FitToPagesWide: expected={a.FitToPagesWide} actual={b.FitToPagesWide}");
        }
        if (a.FitToPagesTall != b.FitToPagesTall)
        {
            diffs.Add($"{path}.FitToPagesTall: expected={a.FitToPagesTall} actual={b.FitToPagesTall}");
        }
        if (a.CenterHorizontally != b.CenterHorizontally)
        {
            diffs.Add($"{path}.CenterHorizontally: expected={a.CenterHorizontally} actual={b.CenterHorizontally}");
        }
        if (a.CenterVertically != b.CenterVertically)
        {
            diffs.Add($"{path}.CenterVertically: expected={a.CenterVertically} actual={b.CenterVertically}");
        }
    }

    private static void CompareHeaderFooter(string path, ReportHeaderFooter a, ReportHeaderFooter b, List<string> diffs)
    {
        if (a.AlignWithMargins != b.AlignWithMargins)
        {
            diffs.Add($"{path}.AlignWithMargins: expected={a.AlignWithMargins} actual={b.AlignWithMargins}");
        }
        if (a.DifferentFirst != b.DifferentFirst)
        {
            diffs.Add($"{path}.DifferentFirst: expected={a.DifferentFirst} actual={b.DifferentFirst}");
        }
        if (a.DifferentOddEven != b.DifferentOddEven)
        {
            diffs.Add($"{path}.DifferentOddEven: expected={a.DifferentOddEven} actual={b.DifferentOddEven}");
        }
        if (a.ScaleWithDocument != b.ScaleWithDocument)
        {
            diffs.Add($"{path}.ScaleWithDocument: expected={a.ScaleWithDocument} actual={b.ScaleWithDocument}");
        }
        if (!NullableEqual(a.OddHeader, b.OddHeader))
        {
            diffs.Add($"{path}.OddHeader: expected='{a.OddHeader}' actual='{b.OddHeader}'");
        }
        if (!NullableEqual(a.OddFooter, b.OddFooter))
        {
            diffs.Add($"{path}.OddFooter: expected='{a.OddFooter}' actual='{b.OddFooter}'");
        }
        if (!NullableEqual(a.EvenHeader, b.EvenHeader))
        {
            diffs.Add($"{path}.EvenHeader: expected='{a.EvenHeader}' actual='{b.EvenHeader}'");
        }
        if (!NullableEqual(a.EvenFooter, b.EvenFooter))
        {
            diffs.Add($"{path}.EvenFooter: expected='{a.EvenFooter}' actual='{b.EvenFooter}'");
        }
        if (!NullableEqual(a.FirstHeader, b.FirstHeader))
        {
            diffs.Add($"{path}.FirstHeader: expected='{a.FirstHeader}' actual='{b.FirstHeader}'");
        }
        if (!NullableEqual(a.FirstFooter, b.FirstFooter))
        {
            diffs.Add($"{path}.FirstFooter: expected='{a.FirstFooter}' actual='{b.FirstFooter}'");
        }
    }

    private static void ComparePrintArea(string path, ReportPrintArea? a, ReportPrintArea? b, List<string> diffs)
    {
        if (a is null && b is null)
        {
            return;
        }
        if (a is null || b is null)
        {
            diffs.Add($"{path}: expected={(a is null ? "null" : a.Range.ToString())} actual={(b is null ? "null" : b.Range.ToString())}");
            return;
        }
        if (!a.Range.Equals(b.Range))
        {
            diffs.Add($"{path}: expected={a.Range} actual={b.Range}");
        }
    }

    private static void CompareRows(string path, IReadOnlyList<ReportRow> a, IReadOnlyList<ReportRow> b, List<string> diffs)
    {
        if (a.Count != b.Count)
        {
            diffs.Add($"{path}.Count: expected={a.Count} actual={b.Count}");
            return;
        }
        for (var i = 0; i < a.Count; i++)
        {
            if (a[i].Index != b[i].Index)
            {
                diffs.Add($"{path}[{i}].Index: expected={a[i].Index} actual={b[i].Index}");
            }
            if (!ApproxEqual(a[i].HeightPoint, b[i].HeightPoint))
            {
                diffs.Add($"{path}[{i}] (row {a[i].Index}).HeightPoint: expected={a[i].HeightPoint} actual={b[i].HeightPoint}");
            }
            if (a[i].IsHidden != b[i].IsHidden)
            {
                diffs.Add($"{path}[{i}] (row {a[i].Index}).IsHidden: expected={a[i].IsHidden} actual={b[i].IsHidden}");
            }
            if (a[i].OutlineLevel != b[i].OutlineLevel)
            {
                diffs.Add($"{path}[{i}] (row {a[i].Index}).OutlineLevel: expected={a[i].OutlineLevel} actual={b[i].OutlineLevel}");
            }
        }
    }

    private static void CompareColumns(string path, IReadOnlyList<ReportColumn> a, IReadOnlyList<ReportColumn> b, List<string> diffs)
    {
        if (a.Count != b.Count)
        {
            diffs.Add($"{path}.Count: expected={a.Count} actual={b.Count}");
            return;
        }
        for (var i = 0; i < a.Count; i++)
        {
            if (a[i].Index != b[i].Index)
            {
                diffs.Add($"{path}[{i}].Index: expected={a[i].Index} actual={b[i].Index}");
            }
            if (!ApproxEqual(a[i].WidthPoint, b[i].WidthPoint, 0.05))
            {
                diffs.Add($"{path}[{i}] (col {a[i].Index}).WidthPoint: expected={a[i].WidthPoint:F3} actual={b[i].WidthPoint:F3}");
            }
            if (a[i].IsHidden != b[i].IsHidden)
            {
                diffs.Add($"{path}[{i}] (col {a[i].Index}).IsHidden: expected={a[i].IsHidden} actual={b[i].IsHidden}");
            }
        }
    }

    private static void CompareMerges(string path, IReadOnlyList<ReportMergedRange> a, IReadOnlyList<ReportMergedRange> b, List<string> diffs)
    {
        var aSet = a.Select(x => x.Range.ToString()).OrderBy(x => x, StringComparer.Ordinal).ToArray();
        var bSet = b.Select(x => x.Range.ToString()).OrderBy(x => x, StringComparer.Ordinal).ToArray();

        if (aSet.Length != bSet.Length)
        {
            diffs.Add($"{path}.Count: expected={aSet.Length} actual={bSet.Length}");
        }

        var onlyA = aSet.Except(bSet).ToArray();
        var onlyB = bSet.Except(aSet).ToArray();
        foreach (var r in onlyA)
        {
            diffs.Add($"{path}: expected contains {r} but actual does not");
        }
        foreach (var r in onlyB)
        {
            diffs.Add($"{path}: actual contains {r} but expected does not");
        }
    }

    private static void CompareImages(string path, IReadOnlyList<ReportImage> a, IReadOnlyList<ReportImage> b, List<string> diffs)
    {
        if (a.Count != b.Count)
        {
            diffs.Add($"{path}.Count: expected={a.Count} actual={b.Count}");
            return;
        }
        for (var i = 0; i < a.Count; i++)
        {
            var ea = a[i];
            var eb = b[i];
            if (ea.FromCellAddress != eb.FromCellAddress)
            {
                diffs.Add($"{path}[{i}].FromCellAddress: expected='{ea.FromCellAddress}' actual='{eb.FromCellAddress}'");
            }
            if (ea.ToCellAddress != eb.ToCellAddress)
            {
                diffs.Add($"{path}[{i}].ToCellAddress: expected='{ea.ToCellAddress}' actual='{eb.ToCellAddress}'");
            }
            if (!ApproxEqual(ea.Offset.X, eb.Offset.X, 0.5) || !ApproxEqual(ea.Offset.Y, eb.Offset.Y, 0.5))
            {
                diffs.Add($"{path}[{i}].Offset: expected=({ea.Offset.X:F2},{ea.Offset.Y:F2}) actual=({eb.Offset.X:F2},{eb.Offset.Y:F2})");
            }
            if (!ApproxEqual(ea.WidthPoint, eb.WidthPoint, 0.5))
            {
                diffs.Add($"{path}[{i}].WidthPoint: expected={ea.WidthPoint:F2} actual={eb.WidthPoint:F2}");
            }
            if (!ApproxEqual(ea.HeightPoint, eb.HeightPoint, 0.5))
            {
                diffs.Add($"{path}[{i}].HeightPoint: expected={ea.HeightPoint:F2} actual={eb.HeightPoint:F2}");
            }
            if (ea.ImageBytes.Length != eb.ImageBytes.Length)
            {
                diffs.Add($"{path}[{i}].ImageBytes.Length: expected={ea.ImageBytes.Length} actual={eb.ImageBytes.Length}");
            }
        }
    }

    private static void CompareBreaks(string path, IReadOnlyList<ReportPageBreak> a, IReadOnlyList<ReportPageBreak> b, List<string> diffs)
    {
        if (a.Count != b.Count)
        {
            diffs.Add($"{path}.Count: expected={a.Count} actual={b.Count}");
            return;
        }
        for (var i = 0; i < a.Count; i++)
        {
            if (a[i].Index != b[i].Index)
            {
                diffs.Add($"{path}[{i}].Index: expected={a[i].Index} actual={b[i].Index}");
            }
        }
    }

    private static void CompareCells(string path, IReadOnlyList<ReportCell> a, IReadOnlyList<ReportCell> b, List<string> diffs)
    {
        var aMap = a.ToDictionary(c => (c.Row, c.Column));
        var bMap = b.ToDictionary(c => (c.Row, c.Column));

        var onlyA = aMap.Keys.Except(bMap.Keys).Take(5).ToArray();
        var onlyB = bMap.Keys.Except(aMap.Keys).Take(5).ToArray();
        foreach (var (row, column) in onlyA)
        {
            diffs.Add($"{path}: expected has cell {AddressHelper.ToAddress(row, column)} but actual does not");
        }
        foreach (var (row, column) in onlyB)
        {
            diffs.Add($"{path}: actual has cell {AddressHelper.ToAddress(row, column)} but expected does not");
        }

        var capture = 0;
        const int maxCellDiffs = 40;
        foreach (var key in aMap.Keys.Intersect(bMap.Keys).OrderBy(k => k.Row).ThenBy(k => k.Column))
        {
            var ea = aMap[key];
            var eb = bMap[key];
            var addr = AddressHelper.ToAddress(key.Row, key.Column);
            var sub = new List<string>();
            CompareCell(ea, eb, sub);
            if (sub.Count > 0)
            {
                foreach (var s in sub)
                {
                    diffs.Add($"{path}[{addr}].{s}");
                    capture++;
                    if (capture >= maxCellDiffs)
                    {
                        diffs.Add($"{path}: truncated (more cells differ)");
                        return;
                    }
                }
            }
        }
    }

    private static void CompareCell(ReportCell a, ReportCell b, List<string> diffs)
    {
        if (a.Value.Kind != b.Value.Kind)
        {
            diffs.Add($"Value.Kind: expected={a.Value.Kind} actual={b.Value.Kind}");
        }
        if (!ObjectEqual(a.Value.RawValue, b.Value.RawValue))
        {
            diffs.Add($"Value.RawValue: expected={FormatValue(a.Value.RawValue)} actual={FormatValue(b.Value.RawValue)}");
        }
        if (a.DisplayText != b.DisplayText)
        {
            diffs.Add($"DisplayText: expected='{a.DisplayText}' actual='{b.DisplayText}'");
        }
        if (a.Style.WrapText != b.Style.WrapText)
        {
            diffs.Add($"Style.WrapText: expected={a.Style.WrapText} actual={b.Style.WrapText}");
        }

        CompareFont(a.Style.Font, b.Style.Font, diffs);
        CompareFill(a.Style.Fill, b.Style.Fill, diffs);
        CompareBorders(a.Style.Borders, b.Style.Borders, diffs);
        CompareAlignment(a.Style.Alignment, b.Style.Alignment, diffs);

        if ((a.Merge is null) != (b.Merge is null))
        {
            diffs.Add($"Merge: expected={(a.Merge is null ? "null" : a.Merge.Range.ToString())} actual={(b.Merge is null ? "null" : b.Merge.Range.ToString())}");
        }
        else if (a.Merge is not null && b.Merge is not null && !a.Merge.Range.Equals(b.Merge.Range))
        {
            diffs.Add($"Merge.Range: expected={a.Merge.Range} actual={b.Merge.Range}");
        }
    }

    private static void CompareFont(ReportFont a, ReportFont b, List<string> diffs)
    {
        if (a.Name != b.Name)
        {
            diffs.Add($"Font.Name: expected='{a.Name}' actual='{b.Name}'");
        }
        if (!ApproxEqual(a.Size, b.Size))
        {
            diffs.Add($"Font.Size: expected={a.Size} actual={b.Size}");
        }
        if (a.Bold != b.Bold)
        {
            diffs.Add($"Font.Bold: expected={a.Bold} actual={b.Bold}");
        }
        if (a.Italic != b.Italic)
        {
            diffs.Add($"Font.Italic: expected={a.Italic} actual={b.Italic}");
        }
        if (a.Underline != b.Underline)
        {
            diffs.Add($"Font.Underline: expected={a.Underline} actual={b.Underline}");
        }
        if (a.Strikeout != b.Strikeout)
        {
            diffs.Add($"Font.Strikeout: expected={a.Strikeout} actual={b.Strikeout}");
        }
        if (!HexEqual(a.ColorHex, b.ColorHex))
        {
            diffs.Add($"Font.ColorHex: expected={a.ColorHex} actual={b.ColorHex}");
        }
    }

    private static void CompareFill(ReportFill a, ReportFill b, List<string> diffs)
    {
        if (!HexEqual(a.BackgroundColorHex, b.BackgroundColorHex))
        {
            diffs.Add($"Fill.BackgroundColorHex: expected={a.BackgroundColorHex} actual={b.BackgroundColorHex}");
        }
    }

    private static void CompareBorders(ReportBorders a, ReportBorders b, List<string> diffs)
    {
        CompareBorder("Borders.Left", a.Left, b.Left, diffs);
        CompareBorder("Borders.Top", a.Top, b.Top, diffs);
        CompareBorder("Borders.Right", a.Right, b.Right, diffs);
        CompareBorder("Borders.Bottom", a.Bottom, b.Bottom, diffs);
    }

    private static void CompareBorder(string label, ReportBorder a, ReportBorder b, List<string> diffs)
    {
        if (a.Style != b.Style)
        {
            diffs.Add($"{label}.Style: expected={a.Style} actual={b.Style}");
        }
        if (!HexEqual(a.ColorHex, b.ColorHex))
        {
            diffs.Add($"{label}.ColorHex: expected={a.ColorHex} actual={b.ColorHex}");
        }
        if (!ApproxEqual(a.Width, b.Width))
        {
            diffs.Add($"{label}.Width: expected={a.Width} actual={b.Width}");
        }
    }

    private static void CompareAlignment(ReportAlignment a, ReportAlignment b, List<string> diffs)
    {
        if (a.Horizontal != b.Horizontal)
        {
            diffs.Add($"Alignment.Horizontal: expected={a.Horizontal} actual={b.Horizontal}");
        }
        if (a.Vertical != b.Vertical)
        {
            diffs.Add($"Alignment.Vertical: expected={a.Vertical} actual={b.Vertical}");
        }
    }

    private static bool ApproxEqual(double a, double b, double eps = 1e-3) => Math.Abs(a - b) <= eps;

    private static bool NullableEqual(string? a, string? b) => String.Equals(a ?? string.Empty, b ?? string.Empty, StringComparison.Ordinal);

    private static bool HexEqual(string a, string b) =>
        String.Equals(a?.ToUpperInvariant(), b?.ToUpperInvariant(), StringComparison.Ordinal);

    private static bool ObjectEqual(object? a, object? b)
    {
        if (a is null && b is null)
        {
            return true;
        }
        if (a is null || b is null)
        {
            return false;
        }
        if (a is double da && b is double db)
        {
            return ApproxEqual(da, db, 1e-9);
        }
        if (a is DateTime dta && b is DateTime dtb)
        {
            return dta.Ticks == dtb.Ticks;
        }
        return a.Equals(b);
    }

    private static string FormatValue(object? value)
    {
        if (value is null)
        {
            return "null";
        }
        if (value is double d)
        {
            return d.ToString("G17", CultureInfo.InvariantCulture);
        }
        if (value is DateTime dt)
        {
            return dt.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture);
        }
        return "'" + value.ToString() + "'";
    }

    private static string FormatThickness(ReportThickness t) =>
        String.Create(CultureInfo.InvariantCulture, $"({t.Left:F2},{t.Top:F2},{t.Right:F2},{t.Bottom:F2})");

    public static string Summarize(IReadOnlyList<string> diffs)
    {
        if (diffs.Count == 0)
        {
            return "No differences.";
        }

        var sb = new StringBuilder();
        sb.Append(CultureInfo.InvariantCulture, $"{diffs.Count} difference(s):\n");
        foreach (var d in diffs)
        {
            sb.Append("  - ").Append(d).Append('\n');
        }
        return sb.ToString();
    }
}
