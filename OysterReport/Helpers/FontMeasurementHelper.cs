namespace OysterReport.Helpers;

using System.Text;

internal static class FontMeasurementHelper
{
    private const double DefaultFallbackDigitWidth = 7d;
    private const double ReferenceFontSizePoints = 10d;
    private const double MeiryoDigitWidthAt10Pt = 8.28125d;
    private const double YuGothicDigitWidthAt10Pt = 8.5d;
    private const double MsPgothicDigitWidthAt10Pt = 6.66667d;
    private const double ArialDigitWidthAt10Pt = 7.41536d;

    public static double ResolveMaxDigitWidth(string? fontName, double fontSize, double fallback = DefaultFallbackDigitWidth)
    {
        if (string.IsNullOrWhiteSpace(fontName) || fontSize <= 0d)
        {
            return fallback;
        }

        var normalized = NormalizeFontName(fontName);
        var baseWidth = normalized switch
        {
            var value when value.Contains("meiryoui", StringComparison.Ordinal) => MeiryoDigitWidthAt10Pt,
            var value when value.Contains("meiryo", StringComparison.Ordinal) => MeiryoDigitWidthAt10Pt,
            var value when value.Contains("yugothicui", StringComparison.Ordinal) => YuGothicDigitWidthAt10Pt,
            var value when value.Contains("yugothic", StringComparison.Ordinal) => YuGothicDigitWidthAt10Pt,
            var value when value.Contains("mspgothic", StringComparison.Ordinal) => MsPgothicDigitWidthAt10Pt,
            var value when value.Contains("msgothic", StringComparison.Ordinal) => MsPgothicDigitWidthAt10Pt,
            var value when value.Contains("arial", StringComparison.Ordinal) => ArialDigitWidthAt10Pt,
            _ => fallback
        };

        return Math.Max(fallback, baseWidth * (fontSize / ReferenceFontSizePoints));
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
