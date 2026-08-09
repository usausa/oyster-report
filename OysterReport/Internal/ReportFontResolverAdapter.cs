namespace OysterReport.Internal;

using System.Collections.Concurrent;
using System.Runtime.InteropServices;

using PdfSharp.Fonts;

#pragma warning disable CA1416
internal sealed class ReportFontResolverAdapter : IFontResolver
{
    private static readonly ConcurrentDictionary<string, FontResolveInfo> ResolvedTypefaceCache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly ConcurrentDictionary<string, EmbeddedFontEntry> EmbeddedFontCache = new(StringComparer.OrdinalIgnoreCase);

    // Data is the defensive copy handed to PDFsharp. Source* remember the identity of the
    // buffer the caller last registered, so re-registrations of the same buffer are O(1).
    private sealed record EmbeddedFontEntry(byte[] Data, object? SourceArray, int SourceOffset, int SourceLength);

    private static readonly Lazy<WindowsFontResolver?> WindowsFallback = new(() => OperatingSystem.IsWindows() ? new WindowsFontResolver() : null);

    //--------------------------------------------------------------------------------
    // Register
    //--------------------------------------------------------------------------------

    public static void RegisterEmbeddedFont(string fontName, ReadOnlyMemory<byte> fontData)
    {
        MemoryMarshal.TryGetArray(fontData, out var segment);

        if (EmbeddedFontCache.TryGetValue(fontName, out var entry))
        {
            // This is called on every font creation during rendering, typically with the same
            // buffer; without these checks each call would copy the whole font again.
            if ((segment.Array is not null) &&
                ReferenceEquals(entry.SourceArray, segment.Array) &&
                (entry.SourceOffset == segment.Offset) &&
                (entry.SourceLength == segment.Count))
            {
                return;
            }

            // Same content in a different buffer: keep the existing copy, remember the new source.
            if (fontData.Span.SequenceEqual(entry.Data))
            {
                EmbeddedFontCache[fontName] = new EmbeddedFontEntry(entry.Data, segment.Array, segment.Offset, segment.Count);
                return;
            }
        }

        // Re-registering an existing name with different data replaces it (last one wins).
        EmbeddedFontCache[fontName] = new EmbeddedFontEntry(fontData.ToArray(), segment.Array, segment.Offset, segment.Count);
    }

    public static void RegisterResolvedTypeface(FontResolveInfo fontResolverInfo)
    {
        ResolvedTypefaceCache[fontResolverInfo.FaceName] = fontResolverInfo;
    }

    //--------------------------------------------------------------------------------
    // IFontResolver
    //--------------------------------------------------------------------------------

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        // Fallback to the base name is handled in GetFont, so bold/italic face names need not be registered separately
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
        if (EmbeddedFontCache.TryGetValue(faceName, out var entry))
        {
            return entry.Data;
        }

        // Falls back to the base family name when bold/italic variants (e.g. "familyName#b") are not individually registered
        var family = ExtractFamilyName(faceName);
        if (!String.Equals(family, faceName, StringComparison.OrdinalIgnoreCase) &&
            EmbeddedFontCache.TryGetValue(family, out entry))
        {
            return entry.Data;
        }

        if (WindowsFallback.Value is not null)
        {
            return WindowsFallback.Value.GetFont(faceName);
        }

        throw new InvalidOperationException($"Font data not provided and no Windows fallback available. faceName=[{faceName}]");
    }

    //--------------------------------------------------------------------------------
    // Bold simulation
    //--------------------------------------------------------------------------------

    public static bool IsBoldSimulationRequired(string faceName, bool isItalic)
    {
        return WindowsFallback.Value is not null && WindowsFallback.Value.IsBoldSimulationRequired(faceName, isItalic);
    }

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private static string BuildFaceName(string familyName, bool bold, bool italic) =>
        (bold, italic) switch
        {
            (true, true) => familyName + "#b#i",
            (true, false) => familyName + "#b",
            (false, true) => familyName + "#i",
            _ => familyName
        };

    private static string ExtractFamilyName(string faceName) =>
        faceName
            .Replace("#b", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("#i", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
}
#pragma warning restore CA1416
