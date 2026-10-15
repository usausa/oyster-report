namespace OysterReport.Internal.Pdf;

// Minimal CFF reader: enough to embed an OpenType/CFF face as CIDFontType0 (FontFile3 /CIDFontType0C).
// Reads the Top DICT (ROS / charset / CharStrings) and the charset to map glyph ids to CIDs.
internal sealed class CffFont
{
    private readonly ushort[] gidToCid;

    public CffFont(ReadOnlySpan<byte> cff, int glyphCount)
    {
        Data = cff.ToArray();
        int hdrSize = cff[2];
        var pos = hdrSize;
        pos = SkipIndex(cff, pos); // Name INDEX
        var (topOffset, topLength) = ReadIndexEntry(cff, pos, 0);
        var dict = ParseDict(cff.Slice(topOffset, topLength));
        IsCidKeyed = dict.ContainsKey(1230); // ROS (12 30)
        gidToCid = new ushort[glyphCount];
        if (dict.TryGetValue(15, out var charsetOffset) && charsetOffset[0] > 2)
        {
            ReadCharset(cff, (int)charsetOffset[0], glyphCount, gidToCid);
        }
        else
        {
            for (var i = 0; i < glyphCount; i++)
            {
                gidToCid[i] = (ushort)i; // predefined charset: only meaningful for name-keyed fonts
            }
        }
    }

    public byte[] Data { get; }

    public bool IsCidKeyed { get; }

    // For CID-keyed CFF the PDF content uses CIDs; for name-keyed CFF CID == GID
    public ushort ToCid(ushort gid) => IsCidKeyed ? gidToCid[gid] : gid;

    private static void ReadCharset(ReadOnlySpan<byte> cff, int offset, int glyphCount, ushort[] map)
    {
        var format = cff[offset];
        var pos = offset + 1;
        var gid = 1;
        switch (format)
        {
            case 0:
                while (gid < glyphCount)
                {
                    map[gid++] = TrueTypeFont.U16(cff, pos);
                    pos += 2;
                }

                break;
            case 1:
            case 2:
                while (gid < glyphCount)
                {
                    int first = TrueTypeFont.U16(cff, pos);
                    int left = format == 1 ? cff[pos + 2] : TrueTypeFont.U16(cff, pos + 2);
                    pos += format == 1 ? 3 : 4;
                    for (var i = 0; i <= left && gid < glyphCount; i++)
                    {
                        map[gid++] = (ushort)(first + i);
                    }
                }

                break;
            default:
                throw new InvalidDataException($"Unsupported CFF charset format. format=[{format}]");
        }
    }

    private static int SkipIndex(ReadOnlySpan<byte> cff, int pos)
    {
        int count = TrueTypeFont.U16(cff, pos);
        if (count == 0)
        {
            return pos + 2;
        }

        int offSize = cff[pos + 2];
        var last = ReadOffset(cff, pos + 3 + (count * offSize), offSize);
        return pos + 3 + ((count + 1) * offSize) + last - 1;
    }

    private static (int Offset, int Length) ReadIndexEntry(ReadOnlySpan<byte> cff, int pos, int index)
    {
        int offSize = cff[pos + 2];
        var start = ReadOffset(cff, pos + 3 + (index * offSize), offSize);
        var end = ReadOffset(cff, pos + 3 + ((index + 1) * offSize), offSize);
        var count = TrueTypeFont.U16(cff, pos);
        var dataStart = pos + 3 + ((count + 1) * offSize) - 1;
        return (dataStart + start, end - start);
    }

    private static int ReadOffset(ReadOnlySpan<byte> d, int pos, int size)
    {
        var v = 0;
        for (var i = 0; i < size; i++)
        {
            v = (v << 8) | d[pos + i];
        }

        return v;
    }

    private static Dictionary<int, List<double>> ParseDict(ReadOnlySpan<byte> d)
    {
        var result = new Dictionary<int, List<double>>();
        var operands = new List<double>();
        var i = 0;
        while (i < d.Length)
        {
            int b0 = d[i];
            if (b0 <= 21)
            {
                var op = b0 == 12 ? 1200 + d[++i] : b0;
                result[op] = operands;
                operands = [];
                i++;
            }
            else if (b0 == 28)
            {
                operands.Add((short)((d[i + 1] << 8) | d[i + 2]));
                i += 3;
            }
            else if (b0 == 29)
            {
                operands.Add((d[i + 1] << 24) | (d[i + 2] << 16) | (d[i + 3] << 8) | d[i + 4]);
                i += 5;
            }
            else if (b0 == 30)
            {
                // real: skip nibbles until 0xf terminator (value not needed here)
                i++;
                while (i < d.Length && (d[i] & 0x0F) != 0x0F && (d[i] & 0xF0) != 0xF0)
                {
                    i++;
                }

                i++;
                operands.Add(0);
            }
            else if (b0 <= 246)
            {
                operands.Add(b0 - 139);
                i++;
            }
            else if (b0 <= 250)
            {
                operands.Add(((b0 - 247) * 256) + d[i + 1] + 108);
                i += 2;
            }
            else
            {
                operands.Add((-(b0 - 251) * 256) - d[i + 1] - 108);
                i += 2;
            }
        }

        return result;
    }
}
