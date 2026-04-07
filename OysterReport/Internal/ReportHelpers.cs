namespace OysterReport.Internal;

using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

using ClosedXML.Excel;

using OysterReport;

internal static partial class PlaceholderParser
{
    public static bool TryParse(string? text, out string markerName)
    {
        markerName = string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var match = PlaceholderRegex().Match(text);
        if (!match.Success)
        {
            return false;
        }

        markerName = match.Groups["name"].Value.Trim();
        return markerName.Length > 0;
    }

    [GeneratedRegex(@"^\{\{(?<name>[^{}]+)\}\}$", RegexOptions.CultureInvariant)]
    private static partial Regex PlaceholderRegex();
}

internal static class ColumnWidthConverter
{
    public static double ToPoint(double excelWidth, double maxDigitWidth, double adjustment)
    {
        var normalizedWidth = Math.Max(0, excelWidth);
        var effectiveMaxDigitWidth = maxDigitWidth <= 0d ? 7d : maxDigitWidth;
        var pixelPadding = (2d * Math.Ceiling(effectiveMaxDigitWidth / 4d)) + 1d;
        double pixelWidth;
        if (normalizedWidth < 1d)
        {
            pixelWidth = normalizedWidth * (effectiveMaxDigitWidth + pixelPadding);
        }
        else
        {
            var normalizedCharacters = ((256d * normalizedWidth) + Math.Round(128d / effectiveMaxDigitWidth)) / 256d;
            pixelWidth = (normalizedCharacters * effectiveMaxDigitWidth) + pixelPadding;
        }

        return pixelWidth * 72d / 96d * adjustment;
    }
}

internal static class FontMeasurementHelper
{
    public static double ResolveMaxDigitWidth(string? fontName, double fontSize, double fallback = 7d)
    {
        if (string.IsNullOrWhiteSpace(fontName) || fontSize <= 0d)
        {
            return fallback;
        }

        var normalized = NormalizeFontName(fontName);
        var baseWidth = normalized switch
        {
            var value when value.Contains("meiryoui", StringComparison.Ordinal) => 8.28125d,
            var value when value.Contains("meiryo", StringComparison.Ordinal) => 8.28125d,
            var value when value.Contains("yugothicui", StringComparison.Ordinal) => 8.5d,
            var value when value.Contains("yugothic", StringComparison.Ordinal) => 8.5d,
            var value when value.Contains("mspgothic", StringComparison.Ordinal) => 6.66667d,
            var value when value.Contains("msgothic", StringComparison.Ordinal) => 6.66667d,
            var value when value.Contains("arial", StringComparison.Ordinal) => 7.41536d,
            _ => fallback
        };

        return Math.Max(fallback, baseWidth * (fontSize / 10d));
    }

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

internal static class ColorHelper
{
    public static string NormalizeHex(string? argb)
    {
        if (string.IsNullOrWhiteSpace(argb))
        {
            return "#00000000";
        }

        var trimmed = argb.Trim();
        return trimmed.Length > 0 && trimmed[0] == '#' ? trimmed.ToUpperInvariant() : $"#{trimmed.ToUpperInvariant()}";
    }

    public static string ToHex(Color color) =>
        NormalizeHex(color.ToArgb().ToString("X8", CultureInfo.InvariantCulture));

    public static string ResolveHex(XLColor color, IXLWorkbook workbook, string defaultHex)
    {
        ArgumentNullException.ThrowIfNull(workbook);

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
        if (double.IsNaN(tint))
        {
            return color;
        }

        var clampedTint = Math.Clamp(tint, -1d, 1d);
        if (Math.Abs(clampedTint) < double.Epsilon)
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

        if (Math.Abs(max - min) < double.Epsilon)
        {
            return (hue, saturation, lightness);
        }

        var delta = max - min;
        saturation = lightness > 0.5d
            ? delta / (2d - max - min)
            : delta / (max + min);

        if (Math.Abs(max - red) < double.Epsilon)
        {
            hue = ((green - blue) / delta) + (green < blue ? 6d : 0d);
        }
        else if (Math.Abs(max - green) < double.Epsilon)
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

internal static class PageSizeResolver
{
    public static (double Width, double Height) GetPageSize(ReportPaperSize paperSize)
    {
        return paperSize switch
        {
            ReportPaperSize.Letter => (612d, 792d),
            ReportPaperSize.Legal => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
    }
}
