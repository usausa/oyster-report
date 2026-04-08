namespace OysterReport.Generator;

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

    private readonly Dictionary<string, (string Path, int FaceIndex)> fontNameToPathAndFace;
    private readonly ConcurrentDictionary<string, byte[]> cache = new(StringComparer.OrdinalIgnoreCase);

    public WindowsInstalledFontResolver()
    {
        fontNameToPathAndFace = LoadFontRegistry();
    }

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
                $"Installed font not found for '{faceName}' (family='{family}', bold={wantBold}, italic={wantItalic}).");
        }

        var rawBytes = File.ReadAllBytes(path);

        if (string.Equals(Path.GetExtension(path), ".ttc", StringComparison.OrdinalIgnoreCase))
        {
            rawBytes = ExtractTtfFaceFromTtc(rawBytes, faceIndex);
        }

        cache[faceName] = rawBytes;
        return rawBytes;
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        return new FontResolverInfo(BuildFaceName(familyName, isBold, isItalic));
    }

    private static byte[] ExtractTtfFaceFromTtc(byte[] ttc, int faceIndex)
    {
        uint ReadUInt32(int offset) =>
            ((uint)ttc[offset] << 24) | ((uint)ttc[offset + 1] << 16) | ((uint)ttc[offset + 2] << 8) | ttc[offset + 3];
        ushort ReadUInt16(int offset) =>
            (ushort)(((uint)ttc[offset] << 8) | ttc[offset + 1]);

        var numFonts = (int)ReadUInt32(8);
        if (faceIndex >= numFonts)
        {
            throw new ArgumentOutOfRangeException(nameof(faceIndex), $"TTC has {numFonts} faces, requested index {faceIndex}.");
        }

        var faceOffset = (int)ReadUInt32(12 + (4 * faceIndex));
        var numTables = ReadUInt16(faceOffset + 4);

        var tables = new (string Tag, uint CheckSum, int SrcOffset, int Length)[numTables];
        for (var i = 0; i < numTables; i++)
        {
            var rec = faceOffset + 12 + (i * 16);
            tables[i] = (
                Tag: System.Text.Encoding.ASCII.GetString(ttc, rec, 4),
                CheckSum: ReadUInt32(rec + 4),
                SrcOffset: (int)ReadUInt32(rec + 8),
                Length: (int)ReadUInt32(rec + 12));
        }

        var headerSize = 12 + (numTables * 16);
        var tableOffsets = new int[numTables];
        var totalSize = headerSize;
        for (var i = 0; i < numTables; i++)
        {
            tableOffsets[i] = totalSize;
            totalSize += tables[i].Length;
            if (totalSize % 4 != 0)
            {
                totalSize += 4 - (totalSize % 4);
            }
        }

        var ttf = new byte[totalSize];
        void WriteUInt32(int offset, uint value)
        {
            ttf[offset] = (byte)(value >> 24);
            ttf[offset + 1] = (byte)(value >> 16);
            ttf[offset + 2] = (byte)(value >> 8);
            ttf[offset + 3] = (byte)value;
        }

        Array.Copy(ttc, faceOffset, ttf, 0, 12);

        for (var i = 0; i < numTables; i++)
        {
            var rec = 12 + (i * 16);
            System.Text.Encoding.ASCII.GetBytes(tables[i].Tag, 0, 4, ttf, rec);
            WriteUInt32(rec + 4, tables[i].CheckSum);
            WriteUInt32(rec + 8, (uint)tableOffsets[i]);
            WriteUInt32(rec + 12, (uint)tables[i].Length);
            Array.Copy(ttc, tables[i].SrcOffset, ttf, tableOffsets[i], tables[i].Length);
        }

        return ttf;
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
            if (key.GetValue(valueName) is not string registryValue || string.IsNullOrWhiteSpace(registryValue))
            {
                continue;
            }

            var path = Path.IsPathRooted(registryValue)
                ? registryValue
                : Path.Combine(WindowsFontsDirectory, registryValue);
            var extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".ttc", StringComparison.OrdinalIgnoreCase))
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
                if (!string.IsNullOrEmpty(name) && !map.ContainsKey(name))
                {
                    map[name] = (path, i);
                }
            }
        }

        return map;
    }

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
