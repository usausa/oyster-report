namespace OysterReport.Internal;

using System.Collections.Concurrent;

using PdfSharp.Fonts;

#pragma warning disable CA1416
internal sealed class ReportFontResolverAdapter : IFontResolver
{
    private static readonly ConcurrentDictionary<string, FontResolveInfo> ResolvedTypefaceCache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly ConcurrentDictionary<string, byte[]> EmbeddedFontCache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly Lazy<WindowsFontResolver?> WindowsFallback = new(CreateWindowsFallback);

    //--------------------------------------------------------------------------------
    // Register
    //--------------------------------------------------------------------------------

    public static void RegisterEmbeddedFont(string fontName, ReadOnlyMemory<byte> fontData)
    {
        EmbeddedFontCache[fontName] = fontData.ToArray();
    }

    public static void RegisterResolvedTypeface(FontResolveInfo fontResolverInfo)
    {
        ResolvedTypefaceCache[fontResolverInfo.FaceName] = fontResolverInfo;
    }

    //--------------------------------------------------------------------------------
    // Bold simulation
    //--------------------------------------------------------------------------------

    public static bool NeedsBoldSimulationForInstalledFont(string faceName, bool isItalic)
    {
        return WindowsFallback.Value is not null && WindowsFallback.Value.NeedsBoldSimulation(faceName, isItalic);
    }

    //--------------------------------------------------------------------------------
    // IFontResolver
    //--------------------------------------------------------------------------------

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        // 事前登録された埋め込みフォントを優先する。
        // GetFont でベース名へのフォールバックを行うため、bold/italic の face 名でも登録不要。
        // Prioritizes pre-registered embedded fonts.
        // Fallback to the base name is handled in GetFont, so bold/italic face names need not be registered separately.
        if (ResolvedTypefaceCache.TryGetValue(familyName, out var resolvedTypeface))
        {
            return new FontResolverInfo(
                BuildFaceName(resolvedTypeface.FaceName, false, false),
                mustSimulateBold: false,
                mustSimulateItalic: resolvedTypeface.MustSimulateItalic);
        }

        if (EmbeddedFontCache.ContainsKey(familyName))
        {
            return new FontResolverInfo(BuildFaceName(familyName, false, false));
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.ResolveTypeface(familyName, isBold, isItalic);
        }

        return new FontResolverInfo(BuildFaceName(familyName, isBold, isItalic));
    }

    public byte[] GetFont(string faceName)
    {
        if (EmbeddedFontCache.TryGetValue(faceName, out var fontBytes))
        {
            return fontBytes;
        }

        // bold/italic バリアント ("familyName#b" 等) が個別登録されていない場合はベース名にフォールバックする。
        // Falls back to the base family name when bold/italic variants (e.g. "familyName#b") are not individually registered.
        var family = ExtractFamilyName(faceName);
        if (!String.Equals(family, faceName, StringComparison.OrdinalIgnoreCase) &&
            EmbeddedFontCache.TryGetValue(family, out fontBytes))
        {
            return fontBytes;
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.GetFont(faceName);
        }

        throw new InvalidOperationException(
            $"Font data not provided and no Windows fallback available. faceName=[{faceName}]");
    }

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private static string BuildFaceName(string familyName, bool bold, bool italic)
    {
        if (!bold && !italic)
        {
            return familyName;
        }

        using var sb = new ValueStringBuilder(stackalloc char[64]);
        sb.Append(familyName);
        if (bold)
        {
            sb.Append("#b");
        }

        if (italic)
        {
            sb.Append("#i");
        }

        return sb.ToString();
    }

    private static string ExtractFamilyName(string faceName) =>
        faceName
            .Replace("#b", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("#i", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

    private static WindowsFontResolver? CreateWindowsFallback() => OperatingSystem.IsWindows() ? new WindowsFontResolver() : null;
}
#pragma warning restore CA1416
