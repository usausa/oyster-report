namespace OysterReport.Helpers;

using System.Text.RegularExpressions;
using OysterReport.Common;

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
        var pixelWidth = (((256 * normalizedWidth) + Math.Truncate(128 / maxDigitWidth)) / 256) * maxDigitWidth;
        return pixelWidth * 72d / 96d * adjustment;
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
}

internal static class PageSizeResolver
{
    public static (double Width, double Height) GetPageSize(ReportPaperSize paperSize)
    {
        return paperSize switch
        {
            ReportPaperSize.Letter => (612d, 792d),
            ReportPaperSize.Legal => (612d, 1008d),
            _ => (595.28d, 841.89d),
        };
    }
}
