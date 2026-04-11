namespace OysterReport.Internal;

using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Versioning;
using System.Text;

using Microsoft.Win32;

using PdfSharp.Fonts;

[SupportedOSPlatform("windows")]
internal sealed class WindowsFontResolver : IFontResolver
{
    private static readonly string WindowsFontsDirectory =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

    private readonly Dictionary<string, (string Path, int FaceIndex)> fontNameToPathAndFace;
    private readonly ConcurrentDictionary<string, byte[]> cache = new(StringComparer.OrdinalIgnoreCase);

    public WindowsFontResolver()
    {
        fontNameToPathAndFace = LoadFontRegistry();
    }

    //--------------------------------------------------------------------------------
    // IFontResolver
    //--------------------------------------------------------------------------------

    public byte[] GetFont(string faceName)
    {
        if (cache.TryGetValue(faceName, out var fontBytes))
        {
            return fontBytes;
        }

        ParseFaceName(faceName, out var family, out var wantBold, out var wantItalic);
        if (!TryFindFont(family, wantBold, wantItalic, out var path, out var faceIndex))
        {
            throw new FileNotFoundException(
                $"Installed font not found. faceName=[{faceName}], family=[{family}], bold=[{wantBold}], italic=[{wantItalic}]");
        }

        var rawBytes = File.ReadAllBytes(path);

        if (String.Equals(Path.GetExtension(path), ".ttc", StringComparison.OrdinalIgnoreCase))
        {
            rawBytes = ExtractTtfFaceFromTtc(rawBytes, faceIndex);
        }

        cache[faceName] = rawBytes;
        return rawBytes;
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        ResolveFontMatch(familyName, isBold, isItalic, out var resolvedFamilyName, out _, out var hasItalicFace);
        return new FontResolverInfo(
            BuildFaceName(resolvedFamilyName, false, false),
            mustSimulateBold: false,
            mustSimulateItalic: isItalic && !hasItalicFace);
    }

    //--------------------------------------------------------------------------------
    // Bold simulation
    //--------------------------------------------------------------------------------

    internal bool NeedsBoldSimulation(string familyName, bool isItalic)
    {
        ResolveFontMatch(familyName, bold: true, isItalic, out _, out var hasBoldFace, out _);
        return !hasBoldFace;
    }

    //--------------------------------------------------------------------------------
    // TTC extraction
    //--------------------------------------------------------------------------------

    private static byte[] ExtractTtfFaceFromTtc(byte[] ttc, int faceIndex)
    {
        var numFonts = (int)BinaryPrimitives.ReadUInt32BigEndian(ttc.AsSpan(8, sizeof(uint)));
        if (faceIndex >= numFonts)
        {
            throw new ArgumentOutOfRangeException(nameof(faceIndex), $"TTC face index out of range. numFonts=[{numFonts}], faceIndex=[{faceIndex}]");
        }

        var faceOffset = (int)BinaryPrimitives.ReadUInt32BigEndian(ttc.AsSpan(12 + (4 * faceIndex), sizeof(uint)));
        var numTables = BinaryPrimitives.ReadUInt16BigEndian(ttc.AsSpan(faceOffset + 4, sizeof(ushort)));

        var tables = new (string Tag, uint CheckSum, int SrcOffset, int Length)[numTables];
        for (var i = 0; i < numTables; i++)
        {
            var ttcTableEntryOffset = faceOffset + 12 + (i * 16);
            tables[i] = (
                Tag: Encoding.ASCII.GetString(ttc, ttcTableEntryOffset, 4),
                CheckSum: BinaryPrimitives.ReadUInt32BigEndian(ttc.AsSpan(ttcTableEntryOffset + 4, sizeof(uint))),
                SrcOffset: (int)BinaryPrimitives.ReadUInt32BigEndian(ttc.AsSpan(ttcTableEntryOffset + 8, sizeof(uint))),
                Length: (int)BinaryPrimitives.ReadUInt32BigEndian(ttc.AsSpan(ttcTableEntryOffset + 12, sizeof(uint))));
        }

        var headerSize = 12 + (numTables * 16);
        var tableOffsets = new int[numTables];
        var totalSize = headerSize;
        for (var i = 0; i < numTables; i++)
        {
            tableOffsets[i] = totalSize;
            totalSize += tables[i].Length;
            var paddingNeeded = totalSize % 4;
            if (paddingNeeded != 0)
            {
                totalSize += 4 - paddingNeeded;
            }
        }

        var ttf = new byte[totalSize];
        Array.Copy(ttc, faceOffset, ttf, 0, 12);

        for (var i = 0; i < numTables; i++)
        {
            var ttfEntryOffset = 12 + (i * 16);
            Encoding.ASCII.GetBytes(tables[i].Tag, 0, 4, ttf, ttfEntryOffset);
            BinaryPrimitives.WriteUInt32BigEndian(ttf.AsSpan(ttfEntryOffset + 4, sizeof(uint)), tables[i].CheckSum);
            BinaryPrimitives.WriteUInt32BigEndian(ttf.AsSpan(ttfEntryOffset + 8, sizeof(uint)), (uint)tableOffsets[i]);
            BinaryPrimitives.WriteUInt32BigEndian(ttf.AsSpan(ttfEntryOffset + 12, sizeof(uint)), (uint)tables[i].Length);
            Array.Copy(ttc, tables[i].SrcOffset, ttf, tableOffsets[i], tables[i].Length);
        }

        return ttf;
    }

    //--------------------------------------------------------------------------------
    // Face name
    //--------------------------------------------------------------------------------

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

    //--------------------------------------------------------------------------------
    // Registry
    //--------------------------------------------------------------------------------

    private static readonly string[] CompoundNameSeparator = [" & "];

    private static Dictionary<string, (string Path, int FaceIndex)> LoadFontRegistry()
    {
        var map = new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase);
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts");
        if (key is null)
        {
            return map;
        }

        foreach (var valueName in key.GetValueNames())
        {
            if (key.GetValue(valueName) is not string registryValue || String.IsNullOrWhiteSpace(registryValue))
            {
                continue;
            }

            var path = Path.IsPathRooted(registryValue)
                ? registryValue
                : Path.Combine(WindowsFontsDirectory, registryValue);
            var extension = Path.GetExtension(path);
            if (!String.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(extension, ".ttc", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var namesPart = valueName;
            var parenIdx = namesPart.LastIndexOf('(');
            if (parenIdx > 0)
            {
                namesPart = namesPart[..parenIdx].TrimEnd(';', ' ');
            }

            var parts = namesPart.Split(CompoundNameSeparator, StringSplitOptions.RemoveEmptyEntries);
            for (var i = 0; i < parts.Length; i++)
            {
                var name = parts[i].Trim();
                if (!String.IsNullOrEmpty(name) && !map.ContainsKey(name))
                {
                    map[name] = (path, i);
                }
            }
        }

        return map;
    }

    //--------------------------------------------------------------------------------
    // Font matching
    //--------------------------------------------------------------------------------

    private bool TryFindFont(
        string family,
        bool bold,
        bool italic,
        [NotNullWhen(true)] out string? path,
        out int faceIndex)
    {
        path = null;
        faceIndex = 0;
        foreach (var candidate in GetCandidateNames(family, bold, italic))
        {
            if (!fontNameToPathAndFace.TryGetValue(candidate, out var info))
            {
                continue;
            }
            if (!File.Exists(info.Path))
            {
                continue;
            }

            path = info.Path;
            faceIndex = info.FaceIndex;
            return true;
        }
        return false;
    }

    private void ResolveFontMatch(
        string family,
        bool bold,
        bool italic,
        out string resolvedFamilyName,
        out bool hasBoldFace,
        out bool hasItalicFace)
    {
        resolvedFamilyName = family;
        hasBoldFace = false;
        hasItalicFace = false;

        if (TryFindFontName(GetCandidateNames(family, bold, italic, includeFallbackFamily: false), out var exactMatch))
        {
            resolvedFamilyName = exactMatch;
            hasBoldFace = bold;
            hasItalicFace = italic;
            return;
        }

        if (bold && TryFindFontName(GetCandidateNames(family, bold: true, italic: false, includeFallbackFamily: false), out var boldMatch))
        {
            resolvedFamilyName = boldMatch;
            hasBoldFace = true;
            return;
        }

        if (italic && TryFindFontName(GetCandidateNames(family, bold: false, italic: true, includeFallbackFamily: false), out var italicMatch))
        {
            resolvedFamilyName = italicMatch;
            hasItalicFace = true;
            return;
        }

        if (TryFindFontName(GetCandidateNames(family, bold: false, italic: false, includeFallbackFamily: true), out var regularMatch))
        {
            resolvedFamilyName = regularMatch;
        }
    }

    private bool TryFindFontName(IEnumerable<string> candidateNames, [NotNullWhen(true)] out string? matchedName)
    {
        matchedName = null;
        foreach (var candidate in candidateNames)
        {
            if (!fontNameToPathAndFace.TryGetValue(candidate, out var info))
            {
                continue;
            }

            if (!File.Exists(info.Path))
            {
                continue;
            }

            matchedName = candidate;
            return true;
        }

        return false;
    }

    private static IEnumerable<string> GetCandidateNames(string family, bool bold, bool italic, bool includeFallbackFamily = true)
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

        if (includeFallbackFamily)
        {
            yield return family;
        }
    }
}
