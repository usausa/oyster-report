namespace OysterReport.Internal.Pdf;

using System.Buffers.Binary;
using System.Text;

internal sealed class TrueTypeFont
{
    private readonly byte[] data;
    private readonly Dictionary<string, (int Offset, int Length)> tables = [with(StringComparer.Ordinal)];
    private readonly CharacterMap cmap = new();
    private uint[] loca = [];
    private bool hasPostScriptName;
    private ushort[] advanceWidths = [];
    private short[] leftSideBearings = [];

    public int UnitsPerEm { get; private set; }

    public int NumGlyphs { get; private set; }

    public bool IsCompactFont { get; private set; }

    public CompactFont? CompactFont { get; private set; }

    public short XMin { get; private set; }

    public short YMin { get; private set; }

    public short XMax { get; private set; }

    public short YMax { get; private set; }

    public double ItalicAngle { get; private set; }

    public bool IsFixedPitch { get; private set; }

    public short UnderlinePosition { get; private set; }

    public short UnderlineThickness { get; private set; }

    public short StrikeoutPosition { get; private set; }

    public short StrikeoutSize { get; private set; }

    public int CapHeight { get; private set; }

    public int XHeight { get; private set; }

    public int FamilyClass { get; private set; }

    public ushort FsType { get; private set; }

    // PDFsharp-compatible cell metrics (font units)
    public int CellAscent { get; private set; }

    public int CellDescent { get; private set; }

    public int CellSpace { get; private set; }

    public string PostScriptName { get; private set; } = "Font";

    private TrueTypeFont(byte[] data) => this.data = data;

    public static TrueTypeFont Load(byte[] data, int faceIndex = 0)
    {
        var font = new TrueTypeFont(data);
        var offset = 0;
        if (U32(data, 0) == 0x74746366)
        {
            var count = (int)U32(data, 8);
            if ((uint)faceIndex >= (uint)count)
            {
                throw new InvalidDataException($"TTC face index out of range. faceIndex=[{faceIndex}], count=[{count}]");
            }

            offset = (int)U32(data, 12 + (faceIndex * 4));
        }

        font.IsCompactFont = U32(data, offset) == 0x4F54544F;
        int numTables = U16(data, offset + 4);
        for (var i = 0; i < numTables; i++)
        {
            var rec = offset + 12 + (i * 16);
            font.tables[Encoding.ASCII.GetString(data, rec, 4)] = ((int)U32(data, rec + 8), (int)U32(data, rec + 12));
        }

        font.Parse();
        return font;
    }

    private void Parse()
    {
        var head = Table("head");
        UnitsPerEm = U16(data, head + 18);
        XMin = S16(data, head + 36);
        YMin = S16(data, head + 38);
        XMax = S16(data, head + 40);
        YMax = S16(data, head + 42);
        var indexToLocFormat = S16(data, head + 50);

        NumGlyphs = U16(data, Table("maxp") + 4);

        var hhea = Table("hhea");
        var hheaAscender = S16(data, hhea + 4);
        var hheaDescender = S16(data, hhea + 6);
        var hheaLineGap = S16(data, hhea + 8);
        int numberOfHMetrics = U16(data, hhea + 34);

        var hmtx = Table("hmtx");
        advanceWidths = new ushort[NumGlyphs];
        leftSideBearings = new short[NumGlyphs];
        for (var g = 0; g < NumGlyphs; g++)
        {
            if (g < numberOfHMetrics)
            {
                advanceWidths[g] = U16(data, hmtx + (g * 4));
                leftSideBearings[g] = S16(data, hmtx + (g * 4) + 2);
            }
            else
            {
                advanceWidths[g] = advanceWidths[numberOfHMetrics - 1];
                leftSideBearings[g] = S16(data, hmtx + (numberOfHMetrics * 4) + ((g - numberOfHMetrics) * 2));
            }
        }

        if (!IsCompactFont)
        {
            var locaOffset = Table("loca");
            loca = new uint[NumGlyphs + 1];
            for (var g = 0; g <= NumGlyphs; g++)
            {
                loca[g] = indexToLocFormat == 0 ? (uint)U16(data, locaOffset + (g * 2)) * 2 : U32(data, locaOffset + (g * 4));
            }
        }

        if (tables.TryGetValue("post", out var post))
        {
            ItalicAngle = S16(data, post.Offset + 4) + (U16(data, post.Offset + 6) / 65536d);
            UnderlinePosition = S16(data, post.Offset + 8);
            UnderlineThickness = S16(data, post.Offset + 10);
            IsFixedPitch = U32(data, post.Offset + 12) != 0;
        }

        short typoAscender = 0, typoDescender = 0, typoLineGap = 0;
        ushort winAscent = 0, winDescent = 0, fsSelection = 0;
        if (tables.TryGetValue("OS/2", out var os2))
        {
            var o = os2.Offset;
            FsType = U16(data, o + 8);
            StrikeoutSize = S16(data, o + 26);
            StrikeoutPosition = S16(data, o + 28);
            FamilyClass = S16(data, o + 30);
            fsSelection = U16(data, o + 62);
            typoAscender = S16(data, o + 68);
            typoDescender = S16(data, o + 70);
            typoLineGap = S16(data, o + 72);
            winAscent = U16(data, o + 74);
            winDescent = U16(data, o + 76);
            if (U16(data, o) >= 2 && os2.Length >= 90)
            {
                XHeight = S16(data, o + 86);
                CapHeight = S16(data, o + 88);
            }
        }

        var os2SeemsToBeEmpty = typoAscender == 0 && typoDescender == 0 && typoLineGap == 0;
        var useTypoMetrics = (fsSelection & 128) != 0;
        if (!os2SeemsToBeEmpty && useTypoMetrics)
        {
            CellAscent = typoAscender + typoLineGap;
            CellDescent = -typoDescender;
            CellSpace = typoAscender + typoLineGap - typoDescender;
        }
        else
        {
            int ascender = hheaAscender;
            var descender = Math.Abs((int)hheaDescender);
            var lineGap = Math.Max((short)0, hheaLineGap);
            if (!os2SeemsToBeEmpty)
            {
                CellAscent = winAscent;
                CellDescent = winDescent;
                CellSpace = Math.Max(lineGap + ascender + descender, winAscent + winDescent);
            }
            else
            {
                CellAscent = ascender;
                CellDescent = descender;
                CellSpace = ascender + descender + lineGap;
            }
        }

        if (CapHeight == 0)
        {
            CapHeight = CellAscent;
        }

        if (XHeight == 0)
        {
            XHeight = (int)(0.66 * CellAscent);
        }

        ParseName();
        ParseCmap();
        if (IsCompactFont)
        {
            var (offset, length) = tables["CFF "];
            CompactFont = new CompactFont(data, offset, length);
            if (!hasPostScriptName)
            {
                // Fall back to the Compact Font name index when the name table has no PostScript name
                PostScriptName = Encoding.ASCII.GetString(CompactFont.Get(CompactFont.Names, 0));
            }
        }
    }

    private int Table(string tag) =>
        tables.TryGetValue(tag, out var t) ? t.Offset : throw new InvalidDataException($"Font table not found. tag=[{tag}]");

    private void ParseName()
    {
        if (!tables.TryGetValue("name", out var name))
        {
            return;
        }

        int count = U16(data, name.Offset + 2);
        int stringOffset = U16(data, name.Offset + 4);
        for (var i = 0; i < count; i++)
        {
            var rec = name.Offset + 6 + (i * 12);
            int platform = U16(data, rec);
            int nameId = U16(data, rec + 6);
            int length = U16(data, rec + 8);
            int off = U16(data, rec + 10);
            if (nameId != 6)
            {
                continue;
            }

            var start = name.Offset + stringOffset + off;
            PostScriptName = platform is 0 or 3 ? Encoding.BigEndianUnicode.GetString(data, start, length) : Encoding.Latin1.GetString(data, start, length);
            hasPostScriptName = true;
            if (platform == 3)
            {
                break;
            }
        }
    }

    private void ParseCmap()
    {
        // Subset fonts produced for PDF embedding have no cmap (glyphs are addressed by id)
        if (!tables.TryGetValue("cmap", out var cmapTable))
        {
            return;
        }

        var cmapOffset = cmapTable.Offset;
        int numSub = U16(data, cmapOffset + 2);
        int best = -1, bestScore = -1;
        for (var i = 0; i < numSub; i++)
        {
            int platform = U16(data, cmapOffset + 4 + (i * 8));
            int encoding = U16(data, cmapOffset + 6 + (i * 8));
            var sub = cmapOffset + (int)U32(data, cmapOffset + 8 + (i * 8));
            int format = U16(data, sub);
            var score = (platform, encoding, format) switch
            {
                (3, 10, 12) => 4,
                (0, _, 12) => 3,
                (3, 1, 4) => 2,
                (0, _, 4) => 1,
                (3, 0, 4) => 0,
                _ => -1
            };
            if (score > bestScore)
            {
                bestScore = score;
                best = sub;
            }
        }

        if (best < 0)
        {
            throw new InvalidDataException("Unsupported cmap.");
        }

        if (U16(data, best) == 12)
        {
            var groups = (int)U32(data, best + 12);
            for (var i = 0; i < groups; i++)
            {
                var g = best + 16 + (i * 12);
                uint start = U32(data, g), end = U32(data, g + 4), gid = U32(data, g + 8);
                for (var c = start; c <= end; c++)
                {
                    cmap.Set((int)c, (ushort)(gid + (c - start)));
                }
            }

            return;
        }

        var segCount = U16(data, best + 6) / 2;
        var endCodes = best + 14;
        var startCodes = endCodes + (segCount * 2) + 2;
        var idDeltas = startCodes + (segCount * 2);
        var idRangeOffsets = idDeltas + (segCount * 2);
        for (var s = 0; s < segCount; s++)
        {
            int end = U16(data, endCodes + (s * 2));
            int start = U16(data, startCodes + (s * 2));
            var delta = S16(data, idDeltas + (s * 2));
            int rangeOffset = U16(data, idRangeOffsets + (s * 2));
            for (var c = start; c <= end && c != 0xFFFF; c++)
            {
                int gid;
                if (rangeOffset == 0)
                {
                    gid = (c + delta) & 0xFFFF;
                }
                else
                {
                    gid = U16(data, idRangeOffsets + (s * 2) + rangeOffset + ((c - start) * 2));
                    if (gid != 0)
                    {
                        gid = (gid + delta) & 0xFFFF;
                    }
                }

                if (gid != 0)
                {
                    cmap.Set(c, (ushort)gid);
                }
            }
        }
    }

    public static bool IsCollection(ReadOnlySpan<byte> data) => (data.Length >= 12) && (U32(data, 0) == 0x74746366);

    // Index of the face in a collection whose family / full / PostScript name equals name (case-insensitive)
    public static int? FindFaceIndex(byte[] data, string name)
    {
        if (!IsCollection(data))
        {
            return null;
        }

        var count = (int)U32(data, 8);
        for (var face = 0; face < count; face++)
        {
            var offset = (int)U32(data, 12 + (face * 4));
            int numTables = U16(data, offset + 4);
            for (var i = 0; i < numTables; i++)
            {
                var rec = offset + 12 + (i * 16);
                if (Encoding.ASCII.GetString(data, rec, 4) != "name")
                {
                    continue;
                }

                if (ReadNames(data, (int)U32(data, rec + 8)).Contains(name, StringComparer.OrdinalIgnoreCase))
                {
                    return face;
                }
            }
        }

        return null;
    }

    // nameID 1 (family), 4 (full), 6 (PostScript), 16 (typographic family)
    private static List<string> ReadNames(byte[] data, int nameOffset)
    {
        var result = new List<string>();
        int count = U16(data, nameOffset + 2);
        int stringOffset = U16(data, nameOffset + 4);
        for (var i = 0; i < count; i++)
        {
            var rec = nameOffset + 6 + (i * 12);
            int platform = U16(data, rec);
            int nameId = U16(data, rec + 6);
            if (nameId is not (1 or 4 or 6 or 16))
            {
                continue;
            }

            var start = nameOffset + stringOffset + U16(data, rec + 10);
            int length = U16(data, rec + 8);
            result.Add(platform is 0 or 3 ? Encoding.BigEndianUnicode.GetString(data, start, length) : Encoding.Latin1.GetString(data, start, length));
        }

        return result;
    }

    public ushort GetGlyphId(int codePoint) => cmap.TryGet(codePoint, out var glyphId) ? glyphId : (ushort)0;

    public bool HasGlyph(int codePoint) => cmap.TryGet(codePoint, out _);

    public int GetAdvanceWidth(int glyphId) => advanceWidths[glyphId];

    public short GetLeftSideBearing(int glyphId) => leftSideBearings[glyphId];

    public ReadOnlySpan<byte> GetGlyphData(int glyphId)
    {
        var glyf = Table("glyf");
        return data.AsSpan(glyf + (int)loca[glyphId], (int)(loca[glyphId + 1] - loca[glyphId]));
    }

    public ReadOnlySpan<byte> GetTable(string tag) => tables.TryGetValue(tag, out var t) ? data.AsSpan(t.Offset, t.Length) : [];

    internal static ushort U16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadUInt16BigEndian(d[o..]);

    internal static short S16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt16BigEndian(d[o..]);

    internal static uint U32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadUInt32BigEndian(d[o..]);
}

internal sealed class CharacterMap
{
    private const int PageShift = 8;
    private const int PageMask = (1 << PageShift) - 1;
    private const int MaxCodePoint = 0x10FFFF;

    private readonly ushort[]?[] pages = new ushort[]?[(MaxCodePoint >> PageShift) + 1];

    private Dictionary<int, ushort>? overflow;

    public void Set(int codePoint, ushort glyphId)
    {
        if (((uint)codePoint > MaxCodePoint) || (glyphId == UInt16.MaxValue))
        {
            overflow ??= [];
            overflow[codePoint] = glyphId;
            if ((uint)codePoint <= MaxCodePoint)
            {
                var existing = pages[codePoint >> PageShift];
                existing?[codePoint & PageMask] = 0;
            }

            return;
        }

        overflow?.Remove(codePoint);
        var page = pages[codePoint >> PageShift] ??= new ushort[PageMask + 1];
        page[codePoint & PageMask] = (ushort)(glyphId + 1);
    }

    public bool TryGet(int codePoint, out ushort glyphId)
    {
        if ((uint)codePoint <= MaxCodePoint)
        {
            var page = pages[codePoint >> PageShift];
            if (page is not null)
            {
                var value = page[codePoint & PageMask];
                if (value != 0)
                {
                    glyphId = (ushort)(value - 1);
                    return true;
                }
            }
        }

        if ((overflow is not null) && overflow.TryGetValue(codePoint, out glyphId))
        {
            return true;
        }

        glyphId = 0;
        return false;
    }
}
