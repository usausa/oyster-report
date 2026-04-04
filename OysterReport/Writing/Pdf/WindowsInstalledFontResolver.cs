namespace OysterReport.Writing.Pdf;

using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Versioning;
using Microsoft.Win32;
using PdfSharp.Fonts;

[SupportedOSPlatform("windows")]
internal sealed class WindowsInstalledFontResolver : IFontResolver
{
    private static readonly string WindowsFontsDirectory =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

    private readonly string[] fallbackFamilies;
    private readonly Dictionary<string, string> fontNameToPath;
    private readonly ConcurrentDictionary<string, byte[]> cache = new(StringComparer.OrdinalIgnoreCase);

    public WindowsInstalledFontResolver(params string[] fallbackFamilies)
    {
        this.fallbackFamilies = fallbackFamilies.Length > 0
            ? fallbackFamilies
            : ["Yu Gothic UI", "Meiryo UI", "Yu Gothic", "Meiryo", "MS UI Gothic", "Segoe UI"];
        fontNameToPath = LoadFontRegistryMap();
    }

    public byte[] GetFont(string faceName)
    {
        if (cache.TryGetValue(faceName, out var fontBytes))
        {
            return fontBytes;
        }

        ParseFaceName(faceName, out var family, out var wantBold, out var wantItalic);
        if (!TryFindFontPath(family, wantBold, wantItalic, out var path))
        {
            foreach (var fallbackFamily in fallbackFamilies)
            {
                if (TryFindFontPath(fallbackFamily, wantBold, wantItalic, out path))
                {
                    break;
                }
            }
        }

        if (path is null)
        {
            throw new FileNotFoundException(
                $"Installed font not found for '{faceName}' (family='{family}', bold={wantBold}, italic={wantItalic}).");
        }

        fontBytes = File.ReadAllBytes(path);
        cache[faceName] = fontBytes;
        return fontBytes;
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        return new FontResolverInfo(BuildFaceName(familyName, isBold, isItalic));
    }

    private static string BuildFaceName(string family, bool bold, bool italic)
    {
        var faceName = family;
        if (bold)
        {
            faceName += "#b";
        }

        if (italic)
        {
            faceName += "#i";
        }

        return faceName;
    }

    private static void ParseFaceName(string faceName, out string family, out bool bold, out bool italic)
    {
        bold = faceName.Contains("#b", StringComparison.OrdinalIgnoreCase);
        italic = faceName.Contains("#i", StringComparison.OrdinalIgnoreCase);
        family = faceName
            .Replace("#b", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("#i", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private static Dictionary<string, string> LoadFontRegistryMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts");
        if (key is null)
        {
            return map;
        }

        foreach (var valueName in key.GetValueNames())
        {
            if (key.GetValue(valueName) is not string registryValue || string.IsNullOrWhiteSpace(registryValue))
            {
                continue;
            }

            var path = Path.IsPathRooted(registryValue)
                ? registryValue
                : Path.Combine(WindowsFontsDirectory, registryValue);
            var extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!map.ContainsKey(valueName))
            {
                map[valueName] = path;
            }
        }

        return map;
    }

    private bool TryFindFontPath(string family, bool bold, bool italic, [NotNullWhen(true)] out string? path)
    {
        path = null;
        foreach (var candidate in GetCandidateNames(family, bold, italic))
        {
            var match = fontNameToPath.Keys.FirstOrDefault(key => key.StartsWith(candidate, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                continue;
            }

            var candidatePath = fontNameToPath[match];
            if (!File.Exists(candidatePath))
            {
                continue;
            }

            path = candidatePath;
            return true;
        }

        return false;
    }

    private static IEnumerable<string> GetCandidateNames(string family, bool bold, bool italic)
    {
        if (bold && italic)
        {
            yield return family + " Bold Italic";
        }

        if (bold)
        {
            yield return family + " Bold";
            yield return family + " SemiBold";
        }

        if (italic)
        {
            yield return family + " Italic";
        }

        yield return family;
    }
}
