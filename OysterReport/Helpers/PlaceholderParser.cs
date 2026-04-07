namespace OysterReport.Helpers;

using System.Text.RegularExpressions;

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
