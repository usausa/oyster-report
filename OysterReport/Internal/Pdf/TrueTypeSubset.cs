// Subset logic ported from fontkit (MIT, Copyright (c) 2014 Devon Govett) src/subset/TTFSubset.js and src/tables/directory.js.
namespace OysterReport.Internal.Pdf;

using System.Buffers.Binary;
using System.Text;

// Port of fontkit TTFSubset: new glyph ids are allocated in first-use order, composite components appended.
internal sealed class TrueTypeSubset(TrueTypeFont font)
{
    private const ushort ArgsAreWords = 0x0001;
    private const ushort WeHaveAScale = 0x0008;
    private const ushort MoreComponents = 0x0020;
    private const ushort WeHaveAnXAndYScale = 0x0040;
    private const ushort WeHaveATwoByTwo = 0x0080;

    private readonly List<int> glyphs = [0];
    private readonly Dictionary<int, int> mapping = new() { [0] = 0 };

    public int IncludeGlyph(int oldGlyphId)
    {
        if (!mapping.TryGetValue(oldGlyphId, out var newId))
        {
            newId = glyphs.Count;
            glyphs.Add(oldGlyphId);
            mapping[oldGlyphId] = newId;
        }

        return newId;
    }

    public byte[] Encode()
    {
        var glyfParts = new List<byte[]>();
        var offsets = new List<uint>();
        var metrics = new List<(ushort Advance, short Bearing)>();
        uint offset = 0;

        // The list can grow while iterating (composite components), same as the JS while-loop
        for (var i = 0; i < glyphs.Count; i++)
        {
            var old = glyphs[i];
            var buffer = font.GetGlyphData(old).ToArray();
            if (buffer.Length >= 10 && TrueTypeFont.S16(buffer, 0) < 0)
            {
                var pos = 10;
                ushort flags;
                do
                {
                    flags = TrueTypeFont.U16(buffer, pos);
                    var component = TrueTypeFont.U16(buffer, pos + 2);
                    BinaryPrimitives.WriteUInt16BigEndian(buffer.AsSpan(pos + 2), (ushort)IncludeGlyph(component));
                    pos += 4 + ((flags & ArgsAreWords) != 0 ? 4 : 2);
                    pos += (flags & WeHaveAScale) != 0 ? 2 : (flags & WeHaveAnXAndYScale) != 0 ? 4 : (flags & WeHaveATwoByTwo) != 0 ? 8 : 0;
                }
                while ((flags & MoreComponents) != 0);
            }

            if ((buffer.Length & 3) != 0)
            {
                Array.Resize(ref buffer, (buffer.Length + 3) & ~3);
            }

            glyfParts.Add(buffer);
            offsets.Add(offset);
            metrics.Add(((ushort)font.GetAdvanceWidth(old), font.GetLeftSideBearing(old)));
            offset += (uint)buffer.Length;
        }

        offsets.Add(offset);

        var head = font.GetTable("head").ToArray();
        BinaryPrimitives.WriteInt16BigEndian(head.AsSpan(50), 1); // long loca
        BinaryPrimitives.WriteUInt32BigEndian(head.AsSpan(8), 0); // checkSumAdjustment (fixed below)

        var hhea = font.GetTable("hhea").ToArray();
        BinaryPrimitives.WriteUInt16BigEndian(hhea.AsSpan(34), (ushort)metrics.Count);

        var maxp = font.GetTable("maxp").ToArray();
        BinaryPrimitives.WriteUInt16BigEndian(maxp.AsSpan(4), (ushort)glyfParts.Count);

        var hmtx = new byte[metrics.Count * 4];
        for (var i = 0; i < metrics.Count; i++)
        {
            BinaryPrimitives.WriteUInt16BigEndian(hmtx.AsSpan(i * 4), metrics[i].Advance);
            BinaryPrimitives.WriteInt16BigEndian(hmtx.AsSpan((i * 4) + 2), metrics[i].Bearing);
        }

        var loca = new byte[offsets.Count * 4];
        for (var i = 0; i < offsets.Count; i++)
        {
            BinaryPrimitives.WriteUInt32BigEndian(loca.AsSpan(i * 4), offsets[i]);
        }

        var glyf = glyfParts.SelectMany(static x => x).ToArray();

        // tables required by PDF spec: head, hhea, loca, maxp, cvt, prep, glyf, hmtx, fpgm
        var outTables = new SortedDictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["head"] = head,
            ["hhea"] = hhea,
            ["loca"] = loca,
            ["maxp"] = maxp,
            ["glyf"] = glyf,
            ["hmtx"] = hmtx,
        };
        foreach (var tag in new[] { "cvt ", "prep", "fpgm" })
        {
            var t = font.GetTable(tag);
            if (!t.IsEmpty)
            {
                outTables[tag] = t.ToArray();
            }
        }

        return WriteDirectory(outTables);
    }

    // Port of fontkit tables/directory.js preEncode/encode
    private static byte[] WriteDirectory(SortedDictionary<string, byte[]> tables)
    {
        var numTables = tables.Count;
        var entrySelector = (int)Math.Floor(Math.Log2(numTables));
        var searchRange = (1 << entrySelector) * 16;
        using var ms = new MemoryStream();
        var header = new byte[12 + (numTables * 16)];
        BinaryPrimitives.WriteUInt32BigEndian(header, 0x00010000);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(4), (ushort)numTables);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(6), (ushort)searchRange);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), (ushort)entrySelector);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(10), (ushort)((numTables * 16) - searchRange));
        ms.Write(header);

        var headOffset = 0;
        var index = 0;
        foreach (var (tag, table) in tables)
        {
            var rec = 12 + (index++ * 16);
            Encoding.ASCII.GetBytes(tag, header.AsSpan(rec));
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(rec + 4), CheckSum(table));
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(rec + 8), (uint)ms.Position);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(rec + 12), (uint)table.Length);
            if (tag == "head")
            {
                headOffset = (int)ms.Position;
            }

            ms.Write(table);
            while ((ms.Position & 3) != 0)
            {
                ms.WriteByte(0);
            }
        }

        var result = ms.ToArray();
        header.CopyTo(result, 0);
        BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(headOffset + 8), 0xB1B0AFBA - CheckSum(result));
        return result;
    }

    private static uint CheckSum(ReadOnlySpan<byte> d)
    {
        uint sum = 0;
        for (var i = 0; i < d.Length; i += 4)
        {
            uint v = 0;
            for (var j = 0; j < 4; j++)
            {
                v = (v << 8) | (i + j < d.Length ? d[i + j] : 0u);
            }

            sum += v;
        }

        return sum;
    }
}
