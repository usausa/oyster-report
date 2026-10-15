namespace OysterReport.Internal;

using System.Collections.Concurrent;

using OysterReport.Internal.Pdf;

internal static class FontMetricsHelper
{
    private const double ReferenceScreenDpi = 96d;
    private const double PointsPerInch = 72d;

    private static readonly ConcurrentDictionary<string, TrueTypeFont?> FontCache = new(StringComparer.OrdinalIgnoreCase);

    public static double? MeasureMaxDigitWidth(string fontFamilyName, double fontSizePoints)
    {
        if (String.IsNullOrWhiteSpace(fontFamilyName) || (fontSizePoints <= 0d))
        {
            return null;
        }

        var font = FontCache.GetOrAdd(fontFamilyName, static name =>
        {
            if (!SystemFontCatalog.Default.TryResolve(name, false, false, out var face, out _, out _))
            {
                return null;
            }

            try
            {
                return TrueTypeFont.Load(SystemFontCatalog.Default.LoadFile(face), face.FaceIndex);
            }
            catch (Exception ex) when (ex is InvalidDataException or ArgumentException or IndexOutOfRangeException)
            {
                return null;
            }
        });
        if (font is null)
        {
            return null;
        }

        // Excel MaxDigitWidth is measured at 96 dpi: pt * (96/72) => pixel
        var pixelSize = fontSizePoints * ReferenceScreenDpi / PointsPerInch;
        var maxAdvance = 0;
        for (var ch = '0'; ch <= '9'; ch++)
        {
            maxAdvance = Math.Max(maxAdvance, font.GetAdvanceWidth(font.GetGlyphId(ch)));
        }

        // [Skia/GDI compat] hinted advances are whole pixels; rounding the linear advance reproduces them
        var width = Math.Round(maxAdvance * pixelSize / font.UnitsPerEm, MidpointRounding.AwayFromZero);
        return width > 0d ? width : null;
    }
}
