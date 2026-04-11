namespace OysterReport.Internal;

using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;

using ClosedXML.Excel;

internal static class TableStyleCatalog
{
    private readonly record struct StyleBand(XLThemeColor ThemeColor, double Tint);

    private static readonly FrozenDictionary<string, StyleBand> Band1RowByStyleName =
        new Dictionary<string, StyleBand>(StringComparer.OrdinalIgnoreCase)
        {
            // ---- Light styles (21 total; L1, L8, L15 are neutral — omitted) ----
            ["TableStyleLight2"] = new(XLThemeColor.Accent1, 0.8),
            ["TableStyleLight3"] = new(XLThemeColor.Accent2, 0.8),
            ["TableStyleLight4"] = new(XLThemeColor.Accent3, 0.8),
            ["TableStyleLight5"] = new(XLThemeColor.Accent4, 0.8),
            ["TableStyleLight6"] = new(XLThemeColor.Accent5, 0.8),
            ["TableStyleLight7"] = new(XLThemeColor.Accent6, 0.8),
            ["TableStyleLight9"] = new(XLThemeColor.Accent1, 0.8),
            ["TableStyleLight10"] = new(XLThemeColor.Accent2, 0.8),
            ["TableStyleLight11"] = new(XLThemeColor.Accent3, 0.8),
            ["TableStyleLight12"] = new(XLThemeColor.Accent4, 0.8),
            ["TableStyleLight13"] = new(XLThemeColor.Accent5, 0.8),
            ["TableStyleLight14"] = new(XLThemeColor.Accent6, 0.8),
            ["TableStyleLight16"] = new(XLThemeColor.Accent1, 0.8),
            ["TableStyleLight17"] = new(XLThemeColor.Accent2, 0.8),
            ["TableStyleLight18"] = new(XLThemeColor.Accent3, 0.8),
            ["TableStyleLight19"] = new(XLThemeColor.Accent4, 0.8),
            ["TableStyleLight20"] = new(XLThemeColor.Accent5, 0.8),
            ["TableStyleLight21"] = new(XLThemeColor.Accent6, 0.8),
            // ---- Medium styles (28 total; M1, M8, M15, M22 are neutral — omitted) ----
            ["TableStyleMedium2"] = new(XLThemeColor.Accent1, 0.2),
            ["TableStyleMedium3"] = new(XLThemeColor.Accent2, 0.2),
            ["TableStyleMedium4"] = new(XLThemeColor.Accent3, 0.2),
            ["TableStyleMedium5"] = new(XLThemeColor.Accent4, 0.2),
            ["TableStyleMedium6"] = new(XLThemeColor.Accent5, 0.2),
            ["TableStyleMedium7"] = new(XLThemeColor.Accent6, 0.2),
            ["TableStyleMedium9"] = new(XLThemeColor.Accent1, 0.2),
            ["TableStyleMedium10"] = new(XLThemeColor.Accent2, 0.2),
            ["TableStyleMedium11"] = new(XLThemeColor.Accent3, 0.2),
            ["TableStyleMedium12"] = new(XLThemeColor.Accent4, 0.2),
            ["TableStyleMedium13"] = new(XLThemeColor.Accent5, 0.2),
            ["TableStyleMedium14"] = new(XLThemeColor.Accent6, 0.2),
            ["TableStyleMedium16"] = new(XLThemeColor.Accent1, 0.2),
            ["TableStyleMedium17"] = new(XLThemeColor.Accent2, 0.2),
            ["TableStyleMedium18"] = new(XLThemeColor.Accent3, 0.2),
            ["TableStyleMedium19"] = new(XLThemeColor.Accent4, 0.2),
            ["TableStyleMedium20"] = new(XLThemeColor.Accent5, 0.2),
            ["TableStyleMedium21"] = new(XLThemeColor.Accent6, 0.2),
            ["TableStyleMedium23"] = new(XLThemeColor.Accent1, 0.2),
            ["TableStyleMedium24"] = new(XLThemeColor.Accent2, 0.2),
            ["TableStyleMedium25"] = new(XLThemeColor.Accent3, 0.2),
            ["TableStyleMedium26"] = new(XLThemeColor.Accent4, 0.2),
            ["TableStyleMedium27"] = new(XLThemeColor.Accent5, 0.2),
            ["TableStyleMedium28"] = new(XLThemeColor.Accent6, 0.2)
        }.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);

    public static bool TryResolveBand1RowFillHex(string styleName, IXLWorkbook workbook, [NotNullWhen(true)] out string? hex)
    {
        if (!Band1RowByStyleName.TryGetValue(styleName, out var band))
        {
            hex = null;
            return false;
        }

        var themeColor = workbook.Theme.ResolveThemeColor(band.ThemeColor);
        if (!themeColor.HasValue)
        {
            hex = null;
            return false;
        }

        hex = ColorHelper.ToHex(ColorHelper.ApplyTint(themeColor.Color, band.Tint));
        return true;
    }
}
