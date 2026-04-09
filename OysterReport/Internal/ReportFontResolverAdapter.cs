namespace OysterReport.Internal;

using System.Collections.Concurrent;

using PdfSharp.Fonts;

// PDFSharp の IFontResolver として登録されるアダプタ。
// PdfGenerator.ResolveFont が事前登録した埋め込みフォントを優先し、
// 未登録の場合は Windows インストール済みフォントへフォールバックする。
#pragma warning disable CA1416
internal sealed class ReportFontResolverAdapter : IFontResolver
{
    // 埋め込みフォントキャッシュ。キー: フォントファミリー名 (大文字小文字無視)
    private static readonly ConcurrentDictionary<string, byte[]> EmbeddedFontCache =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly ConcurrentDictionary<string, FontInfo> ResolvedTypefaceCache =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly Lazy<WindowsInstalledFontResolver?> WindowsFallback =
        new(CreateWindowsFallback);

    /// <summary>
    /// 埋め込みフォントのバイト列を事前登録する。
    /// <see cref="ResolveTypeface"/> より前に呼び出すこと。
    /// </summary>
    public static void RegisterEmbeddedFont(string fontName, ReadOnlyMemory<byte> fontData)
    {
        EmbeddedFontCache[fontName] = fontData.ToArray();
    }

    public static void RegisterResolvedTypeface(FontInfo fontResolverInfo)
    {
        ArgumentNullException.ThrowIfNull(fontResolverInfo);
        ResolvedTypefaceCache[fontResolverInfo.FaceName] = fontResolverInfo;
    }

    public static bool NeedsBoldSimulationForInstalledFont(string faceName, bool isItalic)
    {
        return WindowsFallback.Value is not null && WindowsFallback.Value.NeedsBoldSimulation(faceName, isItalic);
    }

    public byte[] GetFont(string faceName)
    {
        if (EmbeddedFontCache.TryGetValue(faceName, out var fontBytes))
        {
            return fontBytes;
        }

        // bold/italic バリアント ("familyName#b" 等) が個別登録されていない場合は
        // ベース名にフォールバックする。
        var family = ExtractFamilyName(faceName);
        if (!string.Equals(family, faceName, StringComparison.OrdinalIgnoreCase) &&
            EmbeddedFontCache.TryGetValue(family, out fontBytes))
        {
            return fontBytes;
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.GetFont(faceName);
        }

        throw new InvalidOperationException(
            $"Font data was not provided for '{faceName}', and no Windows font fallback is available.");
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        // 事前登録された埋め込みフォントを優先する。
        // GetFont でベース名へのフォールバックを行うため、bold/italic の face 名でも登録不要。
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

    private static string ExtractFamilyName(string faceName) =>
        faceName
            .Replace("#b", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("#i", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

    private static string BuildFaceName(string familyName, bool bold, bool italic)
    {
        var name = familyName;
        if (bold)
        {
            name += "#b";
        }

        if (italic)
        {
            name += "#i";
        }

        return name;
    }

    private static WindowsInstalledFontResolver? CreateWindowsFallback() =>
        OperatingSystem.IsWindows() ? new WindowsInstalledFontResolver() : null;
}
#pragma warning restore CA1416
