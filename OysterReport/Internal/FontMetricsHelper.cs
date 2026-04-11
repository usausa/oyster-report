namespace OysterReport.Internal;

using SkiaSharp;

internal static class FontMetricsHelper
{
    private const double ReferenceScreenDpi = 96d;
    private const double PointsPerInch = 72d;

    public static double? MeasureMaxDigitWidth(string fontFamilyName, double fontSizePoints)
    {
        if (String.IsNullOrWhiteSpace(fontFamilyName) || fontSizePoints <= 0d)
        {
            return null;
        }

        try
        {
            using var typeface = SKTypeface.FromFamilyName(fontFamilyName);
            if (typeface is null)
            {
                return null;
            }

            // Excel MaxDigitWidth : 96dpi
            // SkiaSharp SKFont.Size : pixel unit
            // pt * (96/72) => pixel unit
            var pixelSize = (float)(fontSizePoints * ReferenceScreenDpi / PointsPerInch);
            using var font = new SKFont(typeface, pixelSize);

            var maxWidth = 0f;
            for (var ch = '0'; ch <= '9'; ch++)
            {
                var width = font.MeasureText(ch.ToString());
                if (width > maxWidth)
                {
                    maxWidth = width;
                }
            }

            return maxWidth > 0f ? maxWidth : null;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException)
        {
            return null;
        }
    }
}
