namespace OysterReport.Internal;

using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Threading;

using PdfSharp.Fonts;

#pragma warning disable CA1416
internal sealed class ReportFontResolverAdapter : IFontResolver
{
    private static readonly AsyncLocal<IReportFontResolver?> CurrentResolver = new();
    private static readonly ConcurrentDictionary<string, byte[]> EmbeddedFontCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Lazy<WindowsInstalledFontResolver?> WindowsFallback = new(CreateWindowsFallback);

    public static void SetCurrentResolver(IReportFontResolver? resolver)
    {
        CurrentResolver.Value = resolver;
    }

    public byte[] GetFont(string faceName)
    {
        if (EmbeddedFontCache.TryGetValue(faceName, out var fontBytes))
        {
            return fontBytes;
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.GetFont(faceName);
        }

        throw new InvalidOperationException($"Font data was not provided for '{faceName}', and no Windows font fallback is available.");
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var request = new ReportFontRequest
        {
            FontName = familyName,
            Bold = isBold,
            Italic = isItalic
        };

        var resolution = CurrentResolver.Value?.ResolveFont(request);
        if (resolution is not null)
        {
            var resolvedFontName = string.IsNullOrWhiteSpace(resolution.FontName)
                ? familyName
                : resolution.FontName;
            var faceName = resolution.FontData is ReadOnlyMemory<byte> fontData
                ? BuildEmbeddedFaceName(resolvedFontName, isBold, isItalic, fontData)
                : BuildFaceName(resolvedFontName, isBold, isItalic);

            if (resolution.FontData is ReadOnlyMemory<byte> embeddedFontData)
            {
                EmbeddedFontCache[faceName] = embeddedFontData.ToArray();
            }

            return new FontResolverInfo(faceName);
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.ResolveTypeface(familyName, isBold, isItalic);
        }

        return new FontResolverInfo(BuildFaceName(familyName, isBold, isItalic));
    }

    private static string BuildEmbeddedFaceName(string familyName, bool bold, bool italic, ReadOnlyMemory<byte> fontData)
    {
        var hash = Convert.ToHexString(SHA256.HashData(fontData.Span))[..12];
        return BuildFaceName($"{familyName}#{hash}", bold, italic);
    }

    private static string BuildFaceName(string familyName, bool bold, bool italic)
    {
        var faceName = familyName;
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

    private static WindowsInstalledFontResolver? CreateWindowsFallback()
    {
        return OperatingSystem.IsWindows()
            ? new WindowsInstalledFontResolver()
            : null;
    }
}
#pragma warning restore CA1416
