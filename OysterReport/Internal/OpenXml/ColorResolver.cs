namespace OysterReport.Internal.OpenXml;

using System.Drawing;
using System.Globalization;

using DrawingColor = System.Drawing.Color;
using SsColor = DocumentFormat.OpenXml.Spreadsheet.ColorType;

internal sealed class ColorResolver
{
    // Standard Excel indexed colors (ECMA-376 §18.8.27)
    private static readonly uint[] IndexedPalette =
    [
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF800000, 0xFF008000, 0xFF000080, 0xFF808000, 0xFF800080, 0xFF008080, 0xFFC0C0C0, 0xFF808080,
        0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF, 0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF,
        0xFF000080, 0xFFFF00FF, 0xFFFFFF00, 0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF,
        0xFF00CCFF, 0xFFCCFFFF, 0xFFCCFFCC, 0xFFFFFF99, 0xFF99CCFF, 0xFFFF99CC, 0xFFCC99FF, 0xFFFFCC99,
        0xFF3366FF, 0xFF33CCCC, 0xFF99CC00, 0xFFFFCC00, 0xFFFF9900, 0xFFFF6600, 0xFF666699, 0xFF969696,
        0xFF003366, 0xFF339966, 0xFF003300, 0xFF333300, 0xFF993300, 0xFF993366, 0xFF333399, 0xFF333333,
        0xFF000000, 0xFFFFFFFF
    ];

    private readonly DrawingColor[] themeColors;

    public ColorResolver(DrawingColor[] themeColors)
    {
        this.themeColors = themeColors;
    }

    public bool TryGetThemeColor(int index, out DrawingColor color)
    {
        if ((index < 0) || (index >= themeColors.Length))
        {
            color = default;
            return false;
        }

        color = themeColors[index];
        return true;
    }

    public string Resolve(SsColor? color, string fallbackHex)
    {
        if (color is null)
        {
            return ColorHelper.NormalizeHex(fallbackHex);
        }

        if (color.Auto?.Value == true)
        {
            return ColorHelper.NormalizeHex(fallbackHex);
        }

        if (color.Rgb is not null)
        {
            var rgb = color.Rgb.Value!;
            return ColorHelper.NormalizeHex(NormalizeRgbHex(rgb));
        }

        if (color.Theme is not null)
        {
            var idx = (int)color.Theme.Value;
            if ((idx < 0) || (idx >= themeColors.Length))
            {
                return ColorHelper.NormalizeHex(fallbackHex);
            }

            var baseColor = themeColors[idx];
            var tint = color.Tint?.Value ?? 0d;
            var tinted = Math.Abs(tint) < Double.Epsilon ? baseColor : ColorHelper.ApplyTint(baseColor, tint);
            return ColorHelper.ToHex(tinted);
        }

        if (color.Indexed is not null)
        {
            var idx = (int)color.Indexed.Value;
            if ((idx >= 0) && (idx < IndexedPalette.Length))
            {
                return "#" + IndexedPalette[idx].ToString("X8", CultureInfo.InvariantCulture);
            }
        }

        return ColorHelper.NormalizeHex(fallbackHex);
    }

    private static string NormalizeRgbHex(string raw)
    {
        var trimmed = raw.TrimStart('#').Trim();
        if (trimmed.Length == 6)
        {
            trimmed = "FF" + trimmed;
        }

        return "#" + trimmed.ToUpperInvariant();
    }
}
