namespace OysterReport.Internal;

using System.Drawing;
using System.Globalization;

using ClosedXML.Excel;

internal static class ColorHelper
{
    public static string NormalizeHex(string? argb)
    {
        if (String.IsNullOrWhiteSpace(argb))
        {
            return "#00000000";
        }

        var trimmed = argb.Trim();
        if (trimmed.Length > 0 && trimmed[0] == '#')
        {
            return trimmed.ToUpperInvariant();
        }

        using var sb = new ValueStringBuilder(stackalloc char[12]);
        sb.Append('#');
        sb.Append(trimmed.ToUpperInvariant());
        return sb.ToString();
    }

    public static string ToHex(Color color) =>
        NormalizeHex(color.ToArgb().ToString("X8", CultureInfo.InvariantCulture));

    public static string ResolveHex(XLColor color, IXLWorkbook workbook, string defaultHex)
    {
        var fallback = NormalizeHex(defaultHex);
        if (!color.HasValue)
        {
            return fallback;
        }

        try
        {
            return color.ColorType switch
            {
                XLColorType.Theme => ResolveThemeHex(color, workbook, fallback),
                XLColorType.Indexed => ResolveIndexedHex(color, fallback),
                _ => ToHex(color.Color)
            };
        }
        catch (InvalidOperationException)
        {
            return fallback;
        }
    }

    public static Color ApplyTint(Color color, double tint)
    {
        if (Double.IsNaN(tint))
        {
            return color;
        }

        var clampedTint = Math.Clamp(tint, -1d, 1d);
        if (Math.Abs(clampedTint) < Double.Epsilon)
        {
            return color;
        }

        var (hue, saturation, lightness) = RgbToHsl(color);
        var tintedLightness = clampedTint < 0d
            ? lightness * (1d + clampedTint)
            : lightness + ((1d - lightness) * clampedTint);
        return HslToColor(hue, saturation, tintedLightness, color.A);
    }

    private static string ResolveThemeHex(XLColor color, IXLWorkbook workbook, string fallback)
    {
        var themeColor = workbook.Theme.ResolveThemeColor(color.ThemeColor);
        if (!themeColor.HasValue)
        {
            return fallback;
        }

        return ToHex(ApplyTint(themeColor.Color, color.ThemeTint));
    }

    private static string ResolveIndexedHex(XLColor color, string fallback)
    {
        if (XLColor.IndexedColors.TryGetValue(color.Indexed, out var indexedColor) && indexedColor.HasValue)
        {
            return ToHex(indexedColor.Color);
        }

        return fallback;
    }

    private static (double Hue, double Saturation, double Lightness) RgbToHsl(Color color)
    {
        var red = color.R / 255d;
        var green = color.G / 255d;
        var blue = color.B / 255d;
        var max = Math.Max(red, Math.Max(green, blue));
        var min = Math.Min(red, Math.Min(green, blue));
        var hue = 0d;
        var saturation = 0d;
        var lightness = (max + min) / 2d;

        if (Math.Abs(max - min) < Double.Epsilon)
        {
            return (hue, saturation, lightness);
        }

        var delta = max - min;
        saturation = lightness > 0.5d
            ? delta / (2d - max - min)
            : delta / (max + min);

        if (Math.Abs(max - red) < Double.Epsilon)
        {
            hue = ((green - blue) / delta) + (green < blue ? 6d : 0d);
        }
        else if (Math.Abs(max - green) < Double.Epsilon)
        {
            hue = ((blue - red) / delta) + 2d;
        }
        else
        {
            hue = ((red - green) / delta) + 4d;
        }

        hue /= 6d;
        return (hue, saturation, lightness);
    }

    private static Color HslToColor(double hue, double saturation, double lightness, int alpha)
    {
        if (saturation <= 0d)
        {
            var value = ToByte(lightness * 255d);
            return Color.FromArgb(alpha, value, value, value);
        }

        var q = lightness < 0.5d
            ? lightness * (1d + saturation)
            : lightness + saturation - (lightness * saturation);
        var p = (2d * lightness) - q;
        var red = HueToRgb(p, q, hue + (1d / 3d));
        var green = HueToRgb(p, q, hue);
        var blue = HueToRgb(p, q, hue - (1d / 3d));
        return Color.FromArgb(alpha, ToByte(red * 255d), ToByte(green * 255d), ToByte(blue * 255d));
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0d)
        {
            t += 1d;
        }

        if (t > 1d)
        {
            t -= 1d;
        }

        if (t < (1d / 6d))
        {
            return p + ((q - p) * 6d * t);
        }

        if (t < 0.5d)
        {
            return q;
        }

        if (t < (2d / 3d))
        {
            return p + ((q - p) * ((2d / 3d) - t) * 6d);
        }

        return p;
    }

    private static int ToByte(double value) =>
        (int)Math.Round(Math.Clamp(value, 0d, 255d), MidpointRounding.AwayFromZero);
}
