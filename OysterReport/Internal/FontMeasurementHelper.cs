namespace OysterReport.Internal;

using System.Text;

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
