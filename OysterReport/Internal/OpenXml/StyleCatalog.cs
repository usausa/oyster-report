namespace OysterReport.Internal.OpenXml;

using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable NotAccessedPositionalProperty.Global
internal sealed record FontEntry(
    string Name,
    double Size,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    string ColorHex);

internal sealed record FillEntry(
    FillPattern Pattern,
    string ForegroundHex,
    string BackgroundHex,
    ForegroundColor? RawFg,
    BackgroundColor? RawBg,
    int Reserved);

internal sealed record BorderEntry(
    BorderLineStyle LeftStyle,
    ColorType? LeftColor,
    BorderLineStyle TopStyle,
    ColorType? TopColor,
    BorderLineStyle RightStyle,
    ColorType? RightColor,
    BorderLineStyle BottomStyle,
    ColorType? BottomColor);

internal sealed record CellXfEntry(
    int FontId,
    int FillId,
    int BorderId,
    int NumFmtId,
    bool ApplyFont,
    bool ApplyFill,
    bool ApplyBorder,
    HorizontalAlignment Horizontal,
    VerticalAlignment Vertical,
    bool WrapText);
// ReSharper restore NotAccessedPositionalProperty.Global

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
                [new FillEntry(FillPattern.None, "#00000000", "#00000000", null, null, 0)],
                [EmptyBorder()],
                [new CellXfEntry(0, 0, 0, 0, false, false, false, HorizontalAlignment.General, VerticalAlignment.Bottom, false)],
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

        return ResolveBuiltInNumberFormat(numFmtId);
    }

    public bool IsDateTimeFormat(int numFmtId) => ExcelFormatCode.IsDateTime(ResolveNumberFormat(numFmtId));

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
                yield return new FillEntry(FillPattern.None, "#00000000", "#00000000", null, null, 0);
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
                EnumMaps.ToHorizontalAlignment(xf.Alignment?.Horizontal?.Value),
                EnumMaps.ToVerticalAlignment(xf.Alignment?.Vertical?.Value),
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
            BorderLineStyle.None,
            null,
            BorderLineStyle.None,
            null,
            BorderLineStyle.None,
            null,
            BorderLineStyle.None,
            null);

    private static string ResolveBuiltInNumberFormat(int id) => id switch
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

    private static ArgbColor[] LoadThemeColors(ThemePart? themePart)
    {
        var defaults = new[]
        {
            ArgbColor.White,
            ArgbColor.Black,
            new ArgbColor(0xFF, 0xEE, 0xEC, 0xE1),
            new ArgbColor(0xFF, 0x1F, 0x49, 0x7D),
            new ArgbColor(0xFF, 0x4F, 0x81, 0xBD),
            new ArgbColor(0xFF, 0xC0, 0x50, 0x4D),
            new ArgbColor(0xFF, 0x9B, 0xBB, 0x59),
            new ArgbColor(0xFF, 0x80, 0x64, 0xA2),
            new ArgbColor(0xFF, 0x4B, 0xAC, 0xC6),
            new ArgbColor(0xFF, 0xF7, 0x96, 0x46),
            new ArgbColor(0xFF, 0x00, 0x00, 0xFF),
            new ArgbColor(0xFF, 0x80, 0x00, 0x80)
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

        return
        [
            byName.GetValueOrDefault("lt1", defaults[0]),
            byName.GetValueOrDefault("dk1", defaults[1]),
            byName.GetValueOrDefault("lt2", defaults[2]),
            byName.GetValueOrDefault("dk2", defaults[3]),
            byName.GetValueOrDefault("accent1", defaults[4]),
            byName.GetValueOrDefault("accent2", defaults[5]),
            byName.GetValueOrDefault("accent3", defaults[6]),
            byName.GetValueOrDefault("accent4", defaults[7]),
            byName.GetValueOrDefault("accent5", defaults[8]),
            byName.GetValueOrDefault("accent6", defaults[9]),
            byName.GetValueOrDefault("hlink", defaults[10]),
            byName.GetValueOrDefault("folHlink", defaults[11])
        ];
    }

    private static ArgbColor ParseSchemeColor(XElement element)
    {
        var srgb = element.Descendants(A + "srgbClr").FirstOrDefault();
        if (srgb is not null)
        {
            var val = srgb.Attribute("val")?.Value ?? "000000";
            var argb = Convert.ToUInt32("FF" + val, 16);
            return new ArgbColor(argb);
        }
        var sys = element.Descendants(A + "sysClr").FirstOrDefault();
        if (sys is not null)
        {
            var lastClr = sys.Attribute("lastClr")?.Value;
            if (!String.IsNullOrEmpty(lastClr))
            {
                var argb = Convert.ToUInt32("FF" + lastClr, 16);
                return new ArgbColor(argb);
            }
            var val = sys.Attribute("val")?.Value;
            if (val == "windowText")
            {
                return ArgbColor.Black;
            }
            if (val == "window")
            {
                return ArgbColor.White;
            }
        }
        return ArgbColor.Black;
    }
}
