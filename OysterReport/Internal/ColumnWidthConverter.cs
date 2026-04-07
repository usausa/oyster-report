namespace OysterReport.Internal;

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
