namespace OysterReport.Internal;

using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;

// Cross-platform, fully managed font index (replaces the registry-based Windows resolver).
// Scans font directories, reads only the sfnt table directory + name/OS/2/head tables of each face.
internal sealed class SystemFontCatalog
{
    private static readonly Lazy<SystemFontCatalog> DefaultInstance = new(static () => new SystemFontCatalog(GetDefaultDirectories()));

    private static readonly ConcurrentDictionary<string, Lazy<SystemFontCatalog>> DirectoryCatalogs = new(StringComparer.Ordinal);

    private static readonly ConcurrentDictionary<string, byte[]> FileCache = new(StringComparer.Ordinal);

    private readonly Dictionary<string, List<FontFaceEntry>> facesByFamily = new(StringComparer.OrdinalIgnoreCase);

    // Fonts in the OS font directories
    public static SystemFontCatalog Default => DefaultInstance.Value;

    public int FaceCount { get; private set; }

    public SystemFontCatalog(IEnumerable<string> directories)
    {
        foreach (var directory in directories.Where(Directory.Exists))
        {
            foreach (var path in Directory.EnumerateFiles(directory, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(path);
                if (!ext.Equals(".ttf", StringComparison.OrdinalIgnoreCase) &&
                    !ext.Equals(".ttc", StringComparison.OrdinalIgnoreCase) &&
                    !ext.Equals(".otf", StringComparison.OrdinalIgnoreCase) &&
                    !ext.Equals(".otc", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    IndexFile(path);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or ArgumentException or IndexOutOfRangeException)
                {
                    // Skip broken / unreadable font files
                }
            }
        }
    }

    // Catalog of user-supplied directories (scanned once per distinct directory list)
    public static SystemFontCatalog ForDirectories(IReadOnlyList<string> directories)
    {
        var key = String.Join('\n', directories);
        return DirectoryCatalogs.GetOrAdd(key, _ => new Lazy<SystemFontCatalog>(() => new SystemFontCatalog(directories))).Value;
    }

    public static byte[] LoadFile(FontFaceEntry face) => FileCache.GetOrAdd(face.Path, File.ReadAllBytes);

    public static IEnumerable<string> GetDefaultDirectories()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (OperatingSystem.IsWindows())
        {
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Fonts");
        }
        else if (OperatingSystem.IsMacOS())
        {
            yield return "/System/Library/Fonts";
            yield return "/Library/Fonts";
            yield return Path.Combine(home, "Library", "Fonts");
        }
        else
        {
            yield return "/usr/share/fonts";
            yield return "/usr/local/share/fonts";
            yield return Path.Combine(home, ".local", "share", "fonts");
            yield return Path.Combine(home, ".fonts");
        }
    }

    public bool TryResolve(string family, bool bold, bool italic, out FontFaceEntry face, out bool simulateBold, out bool simulateItalic)
    {
        if (!facesByFamily.TryGetValue(family, out var faces))
        {
            face = default!;
            simulateBold = false;
            simulateItalic = false;
            return false;
        }

        // Exact style first, then the closest face with simulation (same policy as the former WindowsFontResolver)
        var best = faces
            .OrderByDescending(f => (f.Bold == bold ? 2 : 0) + (f.Italic == italic ? 1 : 0))
            .ThenBy(static f => f.Bold || f.Italic ? 1 : 0)
            .First();
        face = best;
        simulateBold = bold && !best.Bold;
        simulateItalic = italic && !best.Italic;
        return true;
    }

    private void IndexFile(string path)
    {
        using var stream = File.OpenRead(path);
        var header = Read(stream, 0, 12);
        var faceOffsets = new List<long>();
        // 'ttcf' = TrueType collection
        if (BinaryPrimitives.ReadUInt32BigEndian(header) == 0x74746366)
        {
            var count = (int)BinaryPrimitives.ReadUInt32BigEndian(header.AsSpan(8));
            var offsets = Read(stream, 12, count * 4);
            for (var i = 0; i < count; i++)
            {
                faceOffsets.Add(BinaryPrimitives.ReadUInt32BigEndian(offsets.AsSpan(i * 4)));
            }
        }
        else
        {
            faceOffsets.Add(0);
        }

        for (var faceIndex = 0; faceIndex < faceOffsets.Count; faceIndex++)
        {
            var dirHeader = Read(stream, faceOffsets[faceIndex], 12);
            var isCff = BinaryPrimitives.ReadUInt32BigEndian(dirHeader) == 0x4F54544F; // 'OTTO'
            int numTables = BinaryPrimitives.ReadUInt16BigEndian(dirHeader.AsSpan(4));
            var records = Read(stream, faceOffsets[faceIndex] + 12, numTables * 16);
            var tables = new Dictionary<string, (long Offset, int Length)>(StringComparer.Ordinal);
            for (var i = 0; i < numTables; i++)
            {
                var r = records.AsSpan(i * 16);
                tables[Encoding.ASCII.GetString(r[..4])] = (BinaryPrimitives.ReadUInt32BigEndian(r[8..]), (int)BinaryPrimitives.ReadUInt32BigEndian(r[12..]));
            }

            if (!tables.TryGetValue("name", out var nameTable) || !tables.TryGetValue("head", out var headTable))
            {
                continue;
            }

            var macStyle = BinaryPrimitives.ReadUInt16BigEndian(Read(stream, headTable.Offset + 44, 2));
            ushort fsSelection = 0;
            if (tables.TryGetValue("OS/2", out var os2))
            {
                fsSelection = BinaryPrimitives.ReadUInt16BigEndian(Read(stream, os2.Offset + 62, 2));
            }

            var bold = (macStyle & 1) != 0 || (fsSelection & 0x20) != 0;
            var italic = (macStyle & 2) != 0 || (fsSelection & 0x01) != 0;
            var families = ReadFamilyNames(Read(stream, nameTable.Offset, nameTable.Length));
            var entry = new FontFaceEntry(path, faceIndex, families.FirstOrDefault() ?? Path.GetFileNameWithoutExtension(path), bold, italic, isCff);
            foreach (var family in families)
            {
                if (!facesByFamily.TryGetValue(family, out var list))
                {
                    list = [];
                    facesByFamily[family] = list;
                }

                list.Add(entry);
            }

            FaceCount++;
        }
    }

    // nameID 1 (family) and 16 (typographic family), all platforms/languages (e.g. "ＭＳ Ｐゴシック" and "MS PGothic")
    private static List<string> ReadFamilyNames(byte[] name)
    {
        var result = new List<string>();
        int count = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(2));
        int stringOffset = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(4));
        for (var i = 0; i < count; i++)
        {
            var rec = name.AsSpan(6 + (i * 12));
            int platform = BinaryPrimitives.ReadUInt16BigEndian(rec);
            int nameId = BinaryPrimitives.ReadUInt16BigEndian(rec[6..]);
            int length = BinaryPrimitives.ReadUInt16BigEndian(rec[8..]);
            int offset = BinaryPrimitives.ReadUInt16BigEndian(rec[10..]);
            if ((nameId != 1 && nameId != 16) || (stringOffset + offset + length > name.Length))
            {
                continue;
            }

            var bytes = name.AsSpan(stringOffset + offset, length);
            var value = platform is 0 or 3 ? Encoding.BigEndianUnicode.GetString(bytes) : Encoding.Latin1.GetString(bytes);
            if (!String.IsNullOrWhiteSpace(value) && !result.Contains(value, StringComparer.OrdinalIgnoreCase))
            {
                result.Add(value);
            }
        }

        return result;
    }

    private static byte[] Read(FileStream stream, long offset, int length)
    {
        var buffer = new byte[length];
        stream.Position = offset;
        stream.ReadExactly(buffer);
        return buffer;
    }
}

internal sealed record FontFaceEntry(string Path, int FaceIndex, string Family, bool Bold, bool Italic, bool IsCff);
