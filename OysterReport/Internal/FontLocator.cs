namespace OysterReport.Internal;

internal sealed class FontLocator
{
    private static readonly Lazy<FontLocator> DefaultInstance = new(static () => Create(new ReportRenderOption()));

    private readonly List<SystemFontCatalog> catalogs = [];

    private readonly Dictionary<string, IReadOnlyList<string>> aliases;

    private FontLocator(ReportRenderOption option)
    {
        if (option.FontDirectories.Count > 0)
        {
            catalogs.Add(SystemFontCatalog.ForDirectories(option.FontDirectories));
        }

        if (option.UseSystemFonts)
        {
            catalogs.Add(SystemFontCatalog.Default);
        }

        aliases = new Dictionary<string, IReadOnlyList<string>>(option.FontAliases, StringComparer.OrdinalIgnoreCase);
    }

    public static FontLocator Default => DefaultInstance.Value;

    public static FontLocator Create(ReportRenderOption option) => new(option);

    public bool TryResolve(string family, bool bold, bool italic, out FontFaceEntry face, out bool simulateBold, out bool simulateItalic)
    {
        foreach (var candidate in Candidates(family))
        {
            foreach (var catalog in catalogs)
            {
                if (catalog.TryResolve(candidate, bold, italic, out face, out simulateBold, out simulateItalic))
                {
                    return true;
                }
            }
        }

        face = default!;
        simulateBold = false;
        simulateItalic = false;
        return false;
    }

    private IEnumerable<string> Candidates(string family)
    {
        yield return family;
        if (aliases.TryGetValue(family, out var substitutes))
        {
            foreach (var substitute in substitutes)
            {
                yield return substitute;
            }
        }
    }
}
