namespace OysterReport.Internal.Pdf;

internal sealed class CompactFont
{
    public const int StandardStringCount = 391;

    private readonly byte[] data;
    private readonly int start;

    public CompactFont(byte[] data, int start, int length)
    {
        this.data = data;
        this.start = start;
        Length = length;

        var pos = (int)Byte(2);
        Names = ReadIndex(ref pos);
        var topDictionaries = ReadIndex(ref pos);
        Strings = ReadIndex(ref pos);
        GlobalSubroutines = ReadIndex(ref pos);

        var (topOffset, topLength) = topDictionaries.Get(0);
        TopDictionary = CompactFontDictionaryEntry.Parse(Slice(topOffset, topLength));
        IsCidKeyed = TopDictionary.Any(static x => x.Operator == CompactFontOperators.RegistryOrderingSupplement);

        var charStringsOffset = (int)Find(TopDictionary, CompactFontOperators.CharStrings)![0];
        var charStringsPos = charStringsOffset;
        CharStrings = ReadIndex(ref charStringsPos);

        if (IsCidKeyed)
        {
            var fontDictionaryArrayPosition = (int)Find(TopDictionary, CompactFontOperators.FontDictionaryArray)![0];
            var fontDictionaryArray = ReadIndex(ref fontDictionaryArrayPosition);
            for (var i = 0; i < fontDictionaryArray.Count; i++)
            {
                var (offset, dictionaryLength) = fontDictionaryArray.Get(i);
                FontDictionaries.Add(ReadFontDictionary(CompactFontDictionaryEntry.Parse(Slice(offset, dictionaryLength))));
            }

            FontDictionarySelection = ReadFontDictionarySelection((int)Find(TopDictionary, CompactFontOperators.FontDictionarySelection)![0], CharStrings.Count);
        }
        else
        {
            FontDictionaries.Add(ReadFontDictionary(TopDictionary));
            FontDictionarySelection = new byte[CharStrings.Count];
        }
    }

    public int Length { get; }

    public CompactFontIndex Names { get; }

    public CompactFontIndex Strings { get; }

    public CompactFontIndex GlobalSubroutines { get; }

    public CompactFontIndex CharStrings { get; }

    public IReadOnlyList<CompactFontDictionaryEntry> TopDictionary { get; }

    public bool IsCidKeyed { get; }

    public List<CompactFontDictionary> FontDictionaries { get; } = [];

    public byte[] FontDictionarySelection { get; }

    public ReadOnlySpan<byte> Slice(int offset, int length) => data.AsSpan(start + offset, length);

    public ReadOnlySpan<byte> Get(CompactFontIndex index, int i)
    {
        var (offset, length) = index.Get(i);
        return Slice(offset, length);
    }

    internal static double[]? Find(IEnumerable<CompactFontDictionaryEntry> dictionary, int @operator) =>
        dictionary.FirstOrDefault(entry => entry.Operator == @operator)?.Values;

    private byte Byte(int offset) => data[start + offset];

    private CompactFontDictionary ReadFontDictionary(IReadOnlyList<CompactFontDictionaryEntry> dictionary)
    {
        var result = new CompactFontDictionary { Entries = [.. dictionary.Where(static entry => entry.Operator != CompactFontOperators.Private)] };
        if (Find(dictionary, CompactFontOperators.Private) is not [var size, var offset])
        {
            return result;
        }

        result.Private = [.. CompactFontDictionaryEntry.Parse(Slice((int)offset, (int)size))];
        if (Find(result.Private, CompactFontOperators.Subroutines) is [var subroutinesOffset])
        {
            var subroutinesPosition = (int)offset + (int)subroutinesOffset;
            result.Subroutines = ReadIndex(ref subroutinesPosition);
        }

        return result;
    }

    private byte[] ReadFontDictionarySelection(int offset, int glyphCount)
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
            throw new InvalidDataException($"Unsupported font dictionary selection format. format=[{format}]");
        }

        int ranges = TrueTypeFont.U16(data, start + offset + 1);
        for (var i = 0; i < ranges; i++)
        {
            var rec = offset + 3 + (i * 3);
            int first = TrueTypeFont.U16(data, start + rec);
            var fontDictionaryIndex = Byte(rec + 2);
            int next = TrueTypeFont.U16(data, start + rec + 3);
            for (var gid = first; gid < next && gid < glyphCount; gid++)
            {
                result[gid] = fontDictionaryIndex;
            }
        }

        return result;
    }

    private CompactFontIndex ReadIndex(ref int pos)
    {
        int count = TrueTypeFont.U16(data, start + pos);
        if (count == 0)
        {
            pos += 2;
            return new CompactFontIndex([0]);
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
        return new CompactFontIndex(offsets);
    }
}

internal sealed class CompactFontIndex(int[] offsets)
{
    public int Count => offsets.Length - 1;

    public (int Offset, int Length) Get(int i) => (offsets[i], offsets[i + 1] - offsets[i]);
}

internal sealed class CompactFontDictionary
{
    public List<CompactFontDictionaryEntry> Entries { get; init; } = [];

    public List<CompactFontDictionaryEntry> Private { get; set; } = [];

    public CompactFontIndex? Subroutines { get; set; }
}

#pragma warning disable IDE0051
internal static class CompactFontOperators
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
    public const int Subroutines = 19;
    public const int Copyright = 1200;
    public const int SyntheticBase = 1220;
    public const int RegistryOrderingSupplement = 1230;
    public const int CidCount = 1234;
    public const int UidBase = 1235;
    public const int FontDictionaryArray = 1236;
    public const int FontDictionarySelection = 1237;
    public const int FontName = 1238;
}
#pragma warning restore IDE0051

internal sealed record CompactFontDictionaryEntry(int Operator, byte[] Operands, double[] Values)
{
    public static List<CompactFontDictionaryEntry> Parse(ReadOnlySpan<byte> data)
    {
        var result = new List<CompactFontDictionaryEntry>();
        var values = new List<double>();
        var operandStart = 0;
        var i = 0;
        while (i < data.Length)
        {
            int b0 = data[i];
            if (b0 <= 21)
            {
                var @operator = b0 == 12 ? 1200 + data[i + 1] : b0;
                result.Add(new CompactFontDictionaryEntry(@operator, [.. data[operandStart..i]], [.. values]));
                i += b0 == 12 ? 2 : 1;
                operandStart = i;
                values.Clear();
            }
            else if (b0 == 28)
            {
                values.Add((short)((data[i + 1] << 8) | data[i + 2]));
                i += 3;
            }
            else if (b0 == 29)
            {
                values.Add((data[i + 1] << 24) | (data[i + 2] << 16) | (data[i + 3] << 8) | data[i + 4]);
                i += 5;
            }
            else if (b0 == 30)
            {
                // Real number: nibbles until 0xF (value is copied verbatim, decoding is not needed)
                i++;
                while ((i < data.Length) && ((data[i] & 0x0F) != 0x0F) && ((data[i] & 0xF0) != 0xF0))
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
                values.Add(((b0 - 247) * 256) + data[i + 1] + 108);
                i += 2;
            }
            else
            {
                values.Add((-(b0 - 251) * 256) - data[i + 1] - 108);
                i += 2;
            }
        }

        return result;
    }

    // Integer operands are written as 5-byte values so offsets can be patched without changing sizes
    public static CompactFontDictionaryEntry Integers(int @operator, params int[] values)
    {
        var operands = new byte[values.Length * 5];
        for (var i = 0; i < values.Length; i++)
        {
            operands[i * 5] = 29;
            System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(operands.AsSpan((i * 5) + 1), values[i]);
        }

        return new CompactFontDictionaryEntry(@operator, operands, [.. values.Select(static x => (double)x)]);
    }

    public static byte[] Encode(IEnumerable<CompactFontDictionaryEntry> entries)
    {
        using var ms = new MemoryStream();
        foreach (var entry in entries)
        {
            ms.Write(entry.Operands);
            if (entry.Operator >= 1200)
            {
                ms.WriteByte(12);
                ms.WriteByte((byte)(entry.Operator - 1200));
            }
            else
            {
                ms.WriteByte((byte)entry.Operator);
            }
        }

        return ms.ToArray();
    }
}
