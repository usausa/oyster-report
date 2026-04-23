namespace OysterReport.Internal.OpenXml;

using System.Globalization;
using System.Xml.Linq;

using ClosedXML.Excel;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using ExcelNumberFormat;

using Color = System.Drawing.Color;

internal sealed class StyleCatalog
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public FontEntry[] Fonts { get; }

    public FillEntry[] Fills { get; }

    public BorderEntry[] Borders { get; }

    public CellXfEntry[] CellXfs { get; }

    public Dictionary<int, string> CustomNumberFormats { get; }

    public ColorResolver ColorResolver { get; }

    public string DefaultFontName { get; }

    public double DefaultFontSize { get; }

    private StyleCatalog(
        FontEntry[] fonts,
        FillEntry[] fills,
        BorderEntry[] borders,
        CellXfEntry[] cellXfs,
        Dictionary<int, string> customFormats,
        ColorResolver resolver,
        string defaultFontName,
        double defaultFontSize)
    {
        Fonts = fonts;
        Fills = fills;
        Borders = borders;
        CellXfs = cellXfs;
        CustomNumberFormats = customFormats;
        ColorResolver = resolver;
        DefaultFontName = defaultFontName;
        DefaultFontSize = defaultFontSize;
    }

    public static StyleCatalog Load(WorkbookPart workbookPart)
    {
        var themeColors = LoadThemeColors(workbookPart.ThemePart);
        var resolver = new ColorResolver(themeColors);

        var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (styles is null)
        {
            return new StyleCatalog(
                [new FontEntry("Calibri", 11, false, false, false, false, "#FF000000")],
                [new FillEntry(XLFillPatternValues.None, "#00000000", "#00000000", null, null, 0)],
                [EmptyBorder()],
                [new CellXfEntry(0, 0, 0, 0, false, false, false, null, null, false)],
                [],
                resolver,
                "Calibri",
                11d);
        }

        var fonts = ReadFonts(styles, resolver).ToArray();
        var fills = ReadFills(styles).ToArray();
        var borders = ReadBorders(styles).ToArray();
        var cellXfs = ReadCellXfs(styles).ToArray();
        var customFormats = ReadCustomNumberFormats(styles);

        var defaultFont = fonts.Length > 0 ? fonts[0] : new FontEntry("Calibri", 11, false, false, false, false, "#FF000000");

        return new StyleCatalog(fonts, fills, borders, cellXfs, customFormats, resolver, defaultFont.Name, defaultFont.Size);
    }

    public string ResolveNumberFormat(int numFmtId)
    {
        if (CustomNumberFormats.TryGetValue(numFmtId, out var code))
        {
            return code;
        }
        return BuiltInNumberFormat.GetCode(numFmtId);
    }

    public bool IsDateTimeFormat(int numFmtId) => NumberFormatCategorizer.IsDateTime(ResolveNumberFormat(numFmtId));

    private static IEnumerable<FontEntry> ReadFonts(Stylesheet styles, ColorResolver resolver)
    {
        if (styles.Fonts is null)
        {
            yield break;
        }

        foreach (var f in styles.Fonts.Elements<Font>())
        {
            var name = f.FontName?.Val?.Value ?? "Calibri";
            var size = f.FontSize?.Val?.Value ?? 11d;
            var bold = f.Bold is not null && (f.Bold.Val?.Value ?? true);
            var italic = f.Italic is not null && (f.Italic.Val?.Value ?? true);
            var underline = f.Underline is not null;
            var strike = f.Strike is not null && (f.Strike.Val?.Value ?? true);
            var colorHex = resolver.Resolve(f.Color, "#FF000000");
            yield return new FontEntry(name, size, bold, italic, underline, strike, colorHex);
        }
    }

    private static IEnumerable<FillEntry> ReadFills(Stylesheet styles)
    {
        if (styles.Fills is null)
        {
            yield break;
        }

        foreach (var fill in styles.Fills.Elements<Fill>())
        {
            var pattern = fill.PatternFill;
            if (pattern is null)
            {
                yield return new FillEntry(XLFillPatternValues.None, "#00000000", "#00000000", null, null, 0);
                continue;
            }

            var patType = EnumMaps.ToFillPattern(pattern.PatternType?.Value);
            yield return new FillEntry(
                patType,
                "#00000000",
                "#00000000",
                pattern.ForegroundColor,
                pattern.BackgroundColor,
                0);
        }
    }

    private static IEnumerable<BorderEntry> ReadBorders(Stylesheet styles)
    {
        if (styles.Borders is null)
        {
            yield break;
        }

        foreach (var b in styles.Borders.Elements<Border>())
        {
            yield return new BorderEntry(
                EnumMaps.ToBorderStyle(b.LeftBorder?.Style?.Value),
                b.LeftBorder?.Color,
                EnumMaps.ToBorderStyle(b.TopBorder?.Style?.Value),
                b.TopBorder?.Color,
                EnumMaps.ToBorderStyle(b.RightBorder?.Style?.Value),
                b.RightBorder?.Color,
                EnumMaps.ToBorderStyle(b.BottomBorder?.Style?.Value),
                b.BottomBorder?.Color);
        }
    }

    private static IEnumerable<CellXfEntry> ReadCellXfs(Stylesheet styles)
    {
        if (styles.CellFormats is null)
        {
            yield break;
        }

        foreach (var xf in styles.CellFormats.Elements<CellFormat>())
        {
            yield return new CellXfEntry(
                (int)(xf.FontId?.Value ?? 0u),
                (int)(xf.FillId?.Value ?? 0u),
                (int)(xf.BorderId?.Value ?? 0u),
                (int)(xf.NumberFormatId?.Value ?? 0u),
                xf.ApplyFont?.Value ?? false,
                xf.ApplyFill?.Value ?? false,
                xf.ApplyBorder?.Value ?? false,
                xf.Alignment?.Horizontal?.Value,
                xf.Alignment?.Vertical?.Value,
                xf.Alignment?.WrapText?.Value ?? false);
        }
    }

    private static Dictionary<int, string> ReadCustomNumberFormats(Stylesheet styles)
    {
        var dict = new Dictionary<int, string>();
        if (styles.NumberingFormats is null)
        {
            return dict;
        }

        foreach (var nf in styles.NumberingFormats.Elements<NumberingFormat>())
        {
            if ((nf.NumberFormatId is null) || (nf.FormatCode is null))
            {
                continue;
            }
            dict[(int)nf.NumberFormatId.Value] = nf.FormatCode.Value ?? string.Empty;
        }

        return dict;
    }

    private static BorderEntry EmptyBorder() =>
        new(
            XLBorderStyleValues.None,
            null,
            XLBorderStyleValues.None,
            null,
            XLBorderStyleValues.None,
            null,
            XLBorderStyleValues.None,
            null);

    private static Color[] LoadThemeColors(ThemePart? themePart)
    {
        var defaults = new[]
        {
            Color.White,
            Color.Black,
            Color.FromArgb(0xFF, 0xEE, 0xEC, 0xE1),
            Color.FromArgb(0xFF, 0x1F, 0x49, 0x7D),
            Color.FromArgb(0xFF, 0x4F, 0x81, 0xBD),
            Color.FromArgb(0xFF, 0xC0, 0x50, 0x4D),
            Color.FromArgb(0xFF, 0x9B, 0xBB, 0x59),
            Color.FromArgb(0xFF, 0x80, 0x64, 0xA2),
            Color.FromArgb(0xFF, 0x4B, 0xAC, 0xC6),
            Color.FromArgb(0xFF, 0xF7, 0x96, 0x46),
            Color.FromArgb(0xFF, 0x00, 0x00, 0xFF),
            Color.FromArgb(0xFF, 0x80, 0x00, 0x80)
        };
        if (themePart is null)
        {
            return defaults;
        }

        using var stream = themePart.GetStream();
        var doc = XDocument.Load(stream);
        var clrScheme = doc.Descendants(A + "clrScheme").FirstOrDefault();
        if (clrScheme is null)
        {
            return defaults;
        }

        var byName = clrScheme.Elements().ToDictionary(
            e => e.Name.LocalName,
            ParseSchemeColor,
            StringComparer.OrdinalIgnoreCase);

        Color Get(string key, Color fallback) => byName.TryGetValue(key, out var c) ? c : fallback;

        return
        [
            Get("lt1", defaults[0]),
            Get("dk1", defaults[1]),
            Get("lt2", defaults[2]),
            Get("dk2", defaults[3]),
            Get("accent1", defaults[4]),
            Get("accent2", defaults[5]),
            Get("accent3", defaults[6]),
            Get("accent4", defaults[7]),
            Get("accent5", defaults[8]),
            Get("accent6", defaults[9]),
            Get("hlink", defaults[10]),
            Get("folHlink", defaults[11])
        ];
    }

    private static Color ParseSchemeColor(XElement element)
    {
        var srgb = element.Descendants(A + "srgbClr").FirstOrDefault();
        if (srgb is not null)
        {
            var val = srgb.Attribute("val")?.Value ?? "000000";
            var argb = Convert.ToInt32("FF" + val, 16);
            return Color.FromArgb(argb);
        }
        var sys = element.Descendants(A + "sysClr").FirstOrDefault();
        if (sys is not null)
        {
            var lastClr = sys.Attribute("lastClr")?.Value;
            if (!String.IsNullOrEmpty(lastClr))
            {
                var argb = Convert.ToInt32("FF" + lastClr, 16);
                return Color.FromArgb(argb);
            }
            var val = sys.Attribute("val")?.Value;
            if (val == "windowText")
            {
                return Color.Black;
            }
            if (val == "window")
            {
                return Color.White;
            }
        }
        return Color.Black;
    }
}

internal sealed record FontEntry(string Name, double Size, bool Bold, bool Italic, bool Underline, bool Strike, string ColorHex);

internal sealed record FillEntry(
    XLFillPatternValues Pattern,
    string ForegroundHex,
    string BackgroundHex,
    DocumentFormat.OpenXml.Spreadsheet.ForegroundColor? RawFg,
    DocumentFormat.OpenXml.Spreadsheet.BackgroundColor? RawBg,
    int Reserved);

internal sealed record BorderEntry(
    XLBorderStyleValues LeftStyle, DocumentFormat.OpenXml.Spreadsheet.ColorType? LeftColor,
    XLBorderStyleValues TopStyle, DocumentFormat.OpenXml.Spreadsheet.ColorType? TopColor,
    XLBorderStyleValues RightStyle, DocumentFormat.OpenXml.Spreadsheet.ColorType? RightColor,
    XLBorderStyleValues BottomStyle, DocumentFormat.OpenXml.Spreadsheet.ColorType? BottomColor);

internal sealed record CellXfEntry(
    int FontId, int FillId, int BorderId, int NumFmtId,
    bool ApplyFont, bool ApplyFill, bool ApplyBorder,
    HorizontalAlignmentValues? Horizontal,
    VerticalAlignmentValues? Vertical,
    bool WrapText);

internal static class BuiltInNumberFormat
{
    public static string GetCode(int id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => "General"
    };
}

internal static class NumberFormatCategorizer
{
    public static bool IsDateTime(string code)
    {
        if (String.IsNullOrEmpty(code) || code == "General" || code == "@")
        {
            return false;
        }

        var inBracket = false;
        foreach (var ch in code)
        {
            if (ch == '[')
            {
                inBracket = true;
            }
            else if (ch == ']')
            {
                inBracket = false;
            }
            else if (!inBracket)
            {
                if (ch is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M')
                {
                    return true;
                }
            }
        }
        return false;
    }

    public static string FormatValue(double numericValue, string formatCode)
    {
        if (String.IsNullOrEmpty(formatCode) || formatCode == "General")
        {
            return numericValue.ToString("G", CultureInfo.InvariantCulture);
        }

        try
        {
            var nf = new NumberFormat(formatCode);
            return nf.IsValid
                ? nf.Format(numericValue, CultureInfo.InvariantCulture)
                : numericValue.ToString("G", CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException)
        {
            return numericValue.ToString("G", CultureInfo.InvariantCulture);
        }
    }
}
