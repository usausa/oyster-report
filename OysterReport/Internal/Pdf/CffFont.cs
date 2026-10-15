namespace OysterReport.Internal.Pdf;

// CFF (Compact Font Format) reader for subsetting. Structures follow fontkit (MIT) src/cff.
// DICT entries keep their raw operand bytes so unchanged values are copied verbatim.
internal sealed class CffFont
{
    public const int StandardStringCount = 391;

    private readonly byte[] data;
    private readonly int start;

    public CffFont(byte[] data, int start, int length)
    {
        this.data = data;
        this.start = start;
        Length = length;

        var pos = (int)Byte(2); // hdrSize
        Names = ReadIndex(ref pos);
        var topDicts = ReadIndex(ref pos);
        Strings = ReadIndex(ref pos);
        GlobalSubrs = ReadIndex(ref pos);

        var (topOffset, topLength) = topDicts.Get(0);
        TopDict = CffDictEntry.Parse(Slice(topOffset, topLength));
        IsCidKeyed = TopDict.Any(static x => x.Op == CffOperators.Ros);

        var charStringsOffset = (int)Find(TopDict, CffOperators.CharStrings)![0];
        var charStringsPos = charStringsOffset;
        CharStrings = ReadIndex(ref charStringsPos);

        if (IsCidKeyed)
        {
            var fdArrayPos = (int)Find(TopDict, CffOperators.FDArray)![0];
            var fdArray = ReadIndex(ref fdArrayPos);
            for (var i = 0; i < fdArray.Count; i++)
            {
                var (offset, length2) = fdArray.Get(i);
                FontDicts.Add(ReadFontDict(CffDictEntry.Parse(Slice(offset, length2))));
            }

            FdSelect = ReadFdSelect((int)Find(TopDict, CffOperators.FDSelect)![0], CharStrings.Count);
        }
        else
        {
            // Name-keyed: the Top DICT's Private DICT acts as the single font dict
            FontDicts.Add(ReadFontDict(TopDict));
            FdSelect = new byte[CharStrings.Count];
        }
    }

    public int Length { get; }

    public CffIndex Names { get; }

    public CffIndex Strings { get; }

    public CffIndex GlobalSubrs { get; }

    public CffIndex CharStrings { get; }

    public IReadOnlyList<CffDictEntry> TopDict { get; }

    public bool IsCidKeyed { get; }

    public List<CffFontDict> FontDicts { get; } = [];

    // gid -> index into FontDicts
    public byte[] FdSelect { get; }

    public ReadOnlySpan<byte> Slice(int offset, int length) => data.AsSpan(start + offset, length);

    public ReadOnlySpan<byte> Get(CffIndex index, int i)
    {
        var (offset, length) = index.Get(i);
        return Slice(offset, length);
    }

    internal static double[]? Find(IEnumerable<CffDictEntry> dict, int op) =>
        dict.FirstOrDefault(x => x.Op == op)?.Values;

    private byte Byte(int offset) => data[start + offset];

    private CffFontDict ReadFontDict(IReadOnlyList<CffDictEntry> dict)
    {
        var result = new CffFontDict { Entries = [.. dict.Where(static x => x.Op != CffOperators.Private)] };
        if (Find(dict, CffOperators.Private) is not [var size, var offset])
        {
            return result;
        }

        result.Private = [.. CffDictEntry.Parse(Slice((int)offset, (int)size))];
        if (Find(result.Private, CffOperators.Subrs) is [var subrsOffset])
        {
            var subrsPos = (int)offset + (int)subrsOffset;
            result.Subrs = ReadIndex(ref subrsPos);
        }

        return result;
    }

    private byte[] ReadFdSelect(int offset, int glyphCount)
    {
        var result = new byte[glyphCount];
        var format = Byte(offset);
        if (format == 0)
        {
            Slice(offset + 1, glyphCount).CopyTo(result);
            return result;
        }

        if (format != 3)
        {
            throw new InvalidDataException($"Unsupported FDSelect format. format=[{format}]");
        }

        int ranges = TrueTypeFont.U16(data, start + offset + 1);
        for (var i = 0; i < ranges; i++)
        {
            var rec = offset + 3 + (i * 3);
            int first = TrueTypeFont.U16(data, start + rec);
            var fd = Byte(rec + 2);
            int next = TrueTypeFont.U16(data, start + rec + 3); // next range's first, or the sentinel
            for (var gid = first; gid < next && gid < glyphCount; gid++)
            {
                result[gid] = fd;
            }
        }

        return result;
    }

    private CffIndex ReadIndex(ref int pos)
    {
        int count = TrueTypeFont.U16(data, start + pos);
        if (count == 0)
        {
            pos += 2;
            return new CffIndex([0]);
        }

        int offSize = Byte(pos + 2);
        var offsets = new int[count + 1];
        var dataStart = pos + 3 + ((count + 1) * offSize) - 1;
        for (var i = 0; i <= count; i++)
        {
            var v = 0;
            for (var b = 0; b < offSize; b++)
            {
                v = (v << 8) | Byte(pos + 3 + (i * offSize) + b);
            }

            offsets[i] = dataStart + v;
        }

        pos = offsets[count];
        return new CffIndex(offsets);
    }
}

// Absolute offsets (from the start of the CFF table) of each item, plus the end offset
internal sealed class CffIndex(int[] offsets)
{
    public int Count => offsets.Length - 1;

    public (int Offset, int Length) Get(int i) => (offsets[i], offsets[i + 1] - offsets[i]);
}

internal sealed class CffFontDict
{
    public List<CffDictEntry> Entries { get; init; } = [];

    public List<CffDictEntry> Private { get; set; } = [];

    public CffIndex? Subrs { get; set; }
}

internal static class CffOperators
{
    public const int Version = 0;
    public const int Notice = 1;
    public const int FullName = 2;
    public const int FamilyName = 3;
    public const int Weight = 4;
    public const int UniqueId = 13;
    public const int Xuid = 14;
    public const int Charset = 15;
    public const int Encoding = 16;
    public const int CharStrings = 17;
    public const int Private = 18;
    public const int Subrs = 19;
    public const int Copyright = 1200;
    public const int SyntheticBase = 1220;
    public const int Ros = 1230;
    public const int CidCount = 1234;
    public const int UidBase = 1235;
    public const int FDArray = 1236;
    public const int FDSelect = 1237;
    public const int FontName = 1238;
}

internal sealed record CffDictEntry(int Op, byte[] Operands, double[] Values)
{
    public static List<CffDictEntry> Parse(ReadOnlySpan<byte> d)
    {
        var result = new List<CffDictEntry>();
        var values = new List<double>();
        var operandStart = 0;
        var i = 0;
        while (i < d.Length)
        {
            int b0 = d[i];
            if (b0 <= 21)
            {
                var op = b0 == 12 ? 1200 + d[i + 1] : b0;
                result.Add(new CffDictEntry(op, d[operandStart..i].ToArray(), [.. values]));
                i += b0 == 12 ? 2 : 1;
                operandStart = i;
                values.Clear();
            }
            else if (b0 == 28)
            {
                values.Add((short)((d[i + 1] << 8) | d[i + 2]));
                i += 3;
            }
            else if (b0 == 29)
            {
                values.Add((d[i + 1] << 24) | (d[i + 2] << 16) | (d[i + 3] << 8) | d[i + 4]);
                i += 5;
            }
            else if (b0 == 30)
            {
                // Real number: nibbles until 0xF (value is copied verbatim, decoding is not needed)
                i++;
                while ((i < d.Length) && ((d[i] & 0x0F) != 0x0F) && ((d[i] & 0xF0) != 0xF0))
                {
                    i++;
                }

                i++;
                values.Add(0);
            }
            else if (b0 <= 246)
            {
                values.Add(b0 - 139);
                i++;
            }
            else if (b0 <= 250)
            {
                values.Add(((b0 - 247) * 256) + d[i + 1] + 108);
                i += 2;
            }
            else
            {
                values.Add((-(b0 - 251) * 256) - d[i + 1] - 108);
                i += 2;
            }
        }

        return result;
    }

    // Integer operands are written as 5-byte values so offsets can be patched without changing sizes
    public static CffDictEntry Integers(int op, params int[] values)
    {
        var operands = new byte[values.Length * 5];
        for (var i = 0; i < values.Length; i++)
        {
            operands[i * 5] = 29;
            System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(operands.AsSpan((i * 5) + 1), values[i]);
        }

        return new CffDictEntry(op, operands, [.. values.Select(static x => (double)x)]);
    }

    public static byte[] Encode(IEnumerable<CffDictEntry> entries)
    {
        using var ms = new MemoryStream();
        foreach (var entry in entries)
        {
            ms.Write(entry.Operands);
            if (entry.Op >= 1200)
            {
                ms.WriteByte(12);
                ms.WriteByte((byte)(entry.Op - 1200));
            }
            else
            {
                ms.WriteByte((byte)entry.Op);
            }
        }

        return ms.ToArray();
    }
}
