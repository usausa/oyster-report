// CFF subsetting ported from fontkit (MIT, Copyright (c) 2014 Devon Govett) src/subset/CFFSubset.js,
// with the subroutine usage scan of src/glyph/CFFGlyph.js.
// Output is always CID-keyed (ROS Adobe-Identity-0) with an identity charset, so CID == new glyph id.
// Subroutine calls are inlined (desubroutinized, like hb-subset's option) so no Subrs are written; when a
// call index is computed at run time the fontkit approach (unused subrs become "return") is used instead.
namespace OysterReport.Internal.Pdf;

internal sealed class CffSubset(CffFont cff, bool desubroutinize = true) : IFontSubset
{
    private const int MaxSubrDepth = 10;

    private static readonly byte[] ReturnOnly = [11];

    private static readonly HashSet<int> DroppedTopOperators =
    [
        CffOperators.Ros, CffOperators.Charset, CffOperators.Encoding, CffOperators.CharStrings, CffOperators.Private,
        CffOperators.FDArray, CffOperators.FDSelect, CffOperators.CidCount, CffOperators.UidBase, CffOperators.UniqueId,
        CffOperators.Xuid, CffOperators.SyntheticBase
    ];

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
        // Font dicts in first-use order and the subroutines each glyph actually calls
        var fdMap = new Dictionary<int, int>();
        var newFdSelect = new byte[glyphs.Count];
        var usedGlobal = new HashSet<int>();
        var usedLocal = new List<HashSet<int>>();
        for (var i = 0; i < glyphs.Count; i++)
        {
            var oldFd = cff.FdSelect[glyphs[i]];
            if (!fdMap.TryGetValue(oldFd, out var newFd))
            {
                newFd = fdMap.Count;
                fdMap[oldFd] = newFd;
                usedLocal.Add([]);
            }

            newFdSelect[i] = (byte)newFd;
            var scanner = new CharStringScanner(cff, cff.FontDicts[oldFd].Subrs, usedGlobal, usedLocal[newFd]);
            scanner.Scan(cff.Get(cff.CharStrings, glyphs[i]), 0);
        }

        var fontDicts = fdMap.OrderBy(static x => x.Value).Select(x => cff.FontDicts[x.Key]).ToList();
        var flattened = desubroutinize ? Flatten() : null;

        // Strings: keep the original String INDEX (SIDs stay valid) and append the ROS strings
        var strings = Enumerable.Range(0, cff.Strings.Count).Select(i => cff.Get(cff.Strings, i).ToArray()).ToList();
        var sidAdobe = CffFont.StandardStringCount + strings.Count;
        strings.Add("Adobe"u8.ToArray());
        strings.Add("Identity"u8.ToArray());

        var nameIndex = WriteIndex([cff.Get(cff.Names, 0).ToArray()]);
        var stringIndex = WriteIndex(strings);
        var globalSubrIndex = WriteIndex(flattened is null ? SubsetSubrs(cff.GlobalSubrs, usedGlobal) : []);
        var charStringIndex = WriteIndex(flattened ?? glyphs.Select(g => cff.Get(cff.CharStrings, g).ToArray()).ToList());
        var charset = glyphs.Count > 1
            ? new byte[] { 2, 0, 1, (byte)((glyphs.Count - 2) >> 8), (byte)(glyphs.Count - 2) }
            : [0];
        var fdSelect = new byte[1 + newFdSelect.Length];
        newFdSelect.CopyTo(fdSelect, 1);

        // Private DICT (+ local Subrs right after it, so the Subrs offset equals the DICT size) for each font dict
        var privates = new List<(byte[] Dict, byte[] Subrs)>();
        for (var i = 0; i < fontDicts.Count; i++)
        {
            var entries = fontDicts[i].Private.Where(static x => x.Op != CffOperators.Subrs).ToList();
            byte[] subrs = [];
            if ((flattened is null) && (fontDicts[i].Subrs is { } localSubrs))
            {
                subrs = WriteIndex(SubsetSubrs(localSubrs, usedLocal[i]));
                var dictSize = CffDictEntry.Encode(entries).Length + 5 + 1;
                entries.Add(CffDictEntry.Integers(CffOperators.Subrs, dictSize));
            }

            privates.Add((CffDictEntry.Encode(entries), subrs));
        }

        // Layout: header, Name, Top DICT, String, Global Subr, charset, FDSelect, CharStrings, FDArray, Private(s)
        var topDict = BuildTopDict(sidAdobe, 0, 0, 0, 0);
        var topDictIndex = WriteIndex([CffDictEntry.Encode(topDict)]);
        var offset = 4 + nameIndex.Length + topDictIndex.Length + stringIndex.Length + globalSubrIndex.Length;
        var charsetOffset = offset;
        var fdSelectOffset = charsetOffset + charset.Length;
        var charStringsOffset = fdSelectOffset + fdSelect.Length;
        var fdArrayOffset = charStringsOffset + charStringIndex.Length;

        var privateOffsets = new List<int>();
        var fdArrayIndex = WriteIndex(fontDicts.Select((fd, i) => BuildFontDict(fd, privates[i].Dict.Length, 0)).ToList());
        var privateOffset = fdArrayOffset + fdArrayIndex.Length;
        for (var i = 0; i < fontDicts.Count; i++)
        {
            privateOffsets.Add(privateOffset);
            privateOffset += privates[i].Dict.Length + privates[i].Subrs.Length;
        }

        fdArrayIndex = WriteIndex(fontDicts.Select((fd, i) => BuildFontDict(fd, privates[i].Dict.Length, privateOffsets[i])).ToList());
        topDictIndex = WriteIndex([CffDictEntry.Encode(BuildTopDict(sidAdobe, charsetOffset, fdSelectOffset, charStringsOffset, fdArrayOffset))]);

        using var output = new MemoryStream();
        output.Write([1, 0, 4, 4]);
        output.Write(nameIndex);
        output.Write(topDictIndex);
        output.Write(stringIndex);
        output.Write(globalSubrIndex);
        output.Write(charset);
        output.Write(fdSelect);
        output.Write(charStringIndex);
        output.Write(fdArrayIndex);
        foreach (var (dict, subrs) in privates)
        {
            output.Write(dict);
            output.Write(subrs);
        }

        return output.ToArray();
    }

    private static int Bias(int count) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768;

    private List<byte[]>? Flatten()
    {
        var result = new List<byte[]>(glyphs.Count);
        foreach (var gid in glyphs)
        {
            var flattener = new CharStringFlattener(cff, cff.FontDicts[cff.FdSelect[gid]].Subrs);
            if (flattener.Flatten(cff.Get(cff.CharStrings, gid)) is not { } code)
            {
                return null;
            }

            result.Add(code);
        }

        return result;
    }

    private List<CffDictEntry> BuildTopDict(int sidAdobe, int charsetOffset, int fdSelectOffset, int charStringsOffset, int fdArrayOffset)
    {
        // ROS must be the first operator of a CIDFont Top DICT
        List<CffDictEntry> dict = [CffDictEntry.Integers(CffOperators.Ros, sidAdobe, sidAdobe + 1, 0)];
        dict.AddRange(cff.TopDict.Where(static x => !DroppedTopOperators.Contains(x.Op)));
        dict.Add(CffDictEntry.Integers(CffOperators.CidCount, glyphs.Count));
        dict.Add(CffDictEntry.Integers(CffOperators.Charset, charsetOffset));
        dict.Add(CffDictEntry.Integers(CffOperators.FDSelect, fdSelectOffset));
        dict.Add(CffDictEntry.Integers(CffOperators.CharStrings, charStringsOffset));
        dict.Add(CffDictEntry.Integers(CffOperators.FDArray, fdArrayOffset));
        return dict;
    }

    private byte[] BuildFontDict(CffFontDict fontDict, int privateSize, int privateOffset)
    {
        // A name-keyed source has no font dict entries of its own (its Top DICT is kept as the CIDFont Top DICT)
        var entries = cff.IsCidKeyed ? fontDict.Entries.Where(static x => x.Op != CffOperators.FontName).ToList() : [];
        entries.Add(CffDictEntry.Integers(CffOperators.Private, privateSize, privateOffset));
        return CffDictEntry.Encode(entries);
    }

    // Unused subroutines are replaced with a bare "return" so the remaining indices stay valid (fontkit subsetSubrs)
    private List<byte[]> SubsetSubrs(CffIndex subrs, HashSet<int> used) =>
        Enumerable.Range(0, subrs.Count).Select(i => used.Contains(i) ? cff.Get(subrs, i).ToArray() : ReturnOnly).ToList();

    private static byte[] WriteIndex(List<byte[]> items)
    {
        if (items.Count == 0)
        {
            return [0, 0];
        }

        var total = items.Sum(static x => x.Length) + 1;
        var offSize = total <= 0xFF ? 1 : total <= 0xFFFF ? 2 : total <= 0xFFFFFF ? 3 : 4;
        using var ms = new MemoryStream();
        ms.WriteByte((byte)(items.Count >> 8));
        ms.WriteByte((byte)items.Count);
        ms.WriteByte((byte)offSize);
        var offset = 1;
        void WriteOffset(int value)
        {
            for (var b = offSize - 1; b >= 0; b--)
            {
                ms.WriteByte((byte)(value >> (b * 8)));
            }
        }

        WriteOffset(offset);
        foreach (var item in items)
        {
            offset += item.Length;
            WriteOffset(offset);
        }

        foreach (var item in items)
        {
            ms.Write(item);
        }

        return ms.ToArray();
    }

    // Inlines callsubr/callgsubr whose index is the literal number right before the call
    private sealed class CharStringFlattener(CffFont cff, CffIndex? localSubrs)
    {
        private readonly List<byte> output = [];
        private int lastNumberStart = -1;
        private double lastNumber;
        private int depth;
        private int stems;

        private enum Result
        {
            Continue,
            EndChar,
            Failed
        }

        public byte[]? Flatten(ReadOnlySpan<byte> code) =>
            Inline(code, 0) == Result.Failed ? null : [.. output];

        private Result Inline(ReadOnlySpan<byte> code, int nesting)
        {
            if (nesting > MaxSubrDepth)
            {
                return Result.Failed;
            }

            var i = 0;
            while (i < code.Length)
            {
                int b0 = code[i];
                if ((b0 >= 32) || (b0 == 28))
                {
                    var length = b0 == 28 ? 3 : b0 <= 246 ? 1 : b0 <= 254 ? 2 : 5;
                    lastNumberStart = output.Count;
                    lastNumber = DecodeNumber(code.Slice(i, length));
                    output.AddRange(code.Slice(i, length));
                    depth++;
                    i += length;
                    continue;
                }

                if (b0 is 10 or 29)
                {
                    var subrs = b0 == 10 ? localSubrs : cff.GlobalSubrs;
                    if ((lastNumberStart < 0) || (subrs is null))
                    {
                        return Result.Failed;
                    }

                    // Drop the index operand; the subroutine body replaces the call
                    output.RemoveRange(lastNumberStart, output.Count - lastNumberStart);
                    depth--;
                    lastNumberStart = -1;
                    var index = (int)lastNumber + Bias(subrs.Count);
                    if ((uint)index >= (uint)subrs.Count)
                    {
                        return Result.Failed;
                    }

                    var result = Inline(cff.Get(subrs, index), nesting + 1);
                    if (result != Result.Continue)
                    {
                        return result;
                    }

                    i++;
                    continue;
                }

                if (b0 == 11)
                {
                    return Result.Continue;
                }

                var opLength = b0 == 12 ? 2 : 1;
                if (b0 is 1 or 3 or 18 or 23)
                {
                    stems += depth / 2;
                }
                else if (b0 is 19 or 20)
                {
                    stems += depth / 2;
                    opLength += (stems + 7) / 8;
                }

                output.AddRange(code.Slice(i, opLength));
                depth = b0 == 12 ? EscapeDepth(code[i + 1], depth) : 0;
                lastNumberStart = -1;
                i += opLength;
                if (b0 == 14)
                {
                    return Result.EndChar;
                }
            }

            return Result.Continue;
        }

        private static double DecodeNumber(ReadOnlySpan<byte> n) =>
            n[0] switch
            {
                28 => (short)((n[1] << 8) | n[2]),
                <= 246 => n[0] - 139,
                <= 250 => ((n[0] - 247) * 256) + n[1] + 108,
                <= 254 => (-(n[0] - 251) * 256) - n[1] - 108,
                _ => ((n[1] << 24) | (n[2] << 16) | (n[3] << 8) | n[4]) / 65536d
            };

        // Operand count after an escape operator (needed to count stems for hintmask)
        private static int EscapeDepth(byte op, int depth) =>
            op switch
            {
                3 or 4 or 10 or 11 or 12 or 15 or 24 or 18 => Math.Max(depth - 1, 0),
                5 or 9 or 14 or 26 or 21 or 28 or 29 => depth,
                20 or 30 => Math.Max(depth - 2, 0),
                22 => Math.Max(depth - 3, 0),
                23 or 27 => depth + 1,
                _ => 0
            };
    }

    // Type 2 charstring walker that only tracks subroutine calls (stack effects of path operators are "clear")
    private sealed class CharStringScanner(CffFont cff, CffIndex? localSubrs, HashSet<int> usedGlobal, HashSet<int> usedLocal)
    {
        private readonly List<double> stack = [];
        private readonly double[] transient = new double[32];
        private int stems;

        // Returns true when endchar was reached
        public bool Scan(ReadOnlySpan<byte> code, int depth)
        {
            if (depth > MaxSubrDepth)
            {
                throw new InvalidDataException("CFF subroutine nesting is too deep.");
            }

            var i = 0;
            while (i < code.Length)
            {
                int b0 = code[i];
                if (b0 >= 32)
                {
                    i += PushNumber(code, i);
                    continue;
                }

                i++;
                switch (b0)
                {
                    case 28:
                        stack.Add((short)((code[i] << 8) | code[i + 1]));
                        i += 2;
                        break;
                    case 1 or 3 or 18 or 23: // hstem vstem hstemhm vstemhm
                        stems += stack.Count / 2;
                        stack.Clear();
                        break;
                    case 19 or 20: // hintmask cntrmask (may carry implicit vstem operands)
                        stems += stack.Count / 2;
                        stack.Clear();
                        i += (stems + 7) / 8;
                        break;
                    case 10: // callsubr
                        if (localSubrs is not null && CallSubr(localSubrs, usedLocal, depth))
                        {
                            return true;
                        }

                        break;
                    case 29: // callgsubr
                        if (CallSubr(cff.GlobalSubrs, usedGlobal, depth))
                        {
                            return true;
                        }

                        break;
                    case 11: // return
                        return false;
                    case 14: // endchar
                        return true;
                    case 12:
                        Escape(code[i++]);
                        break;
                    default: // path construction operators consume the whole stack
                        stack.Clear();
                        break;
                }
            }

            return false;
        }

        private bool CallSubr(CffIndex subrs, HashSet<int> used, int depth)
        {
            if (stack.Count == 0)
            {
                return false;
            }

            var index = (int)Pop() + Bias(subrs.Count);
            if ((uint)index >= (uint)subrs.Count)
            {
                return false;
            }

            used.Add(index);
            return Scan(cff.Get(subrs, index), depth + 1);
        }

        private int PushNumber(ReadOnlySpan<byte> code, int i)
        {
            int b0 = code[i];
            if (b0 <= 246)
            {
                stack.Add(b0 - 139);
                return 1;
            }

            if (b0 <= 250)
            {
                stack.Add(((b0 - 247) * 256) + code[i + 1] + 108);
                return 2;
            }

            if (b0 <= 254)
            {
                stack.Add((-(b0 - 251) * 256) - code[i + 1] - 108);
                return 2;
            }

            stack.Add(((code[i + 1] << 24) | (code[i + 2] << 16) | (code[i + 3] << 8) | code[i + 4]) / 65536d);
            return 5;
        }

        private double Pop()
        {
            if (stack.Count == 0)
            {
                return 0;
            }

            var v = stack[^1];
            stack.RemoveAt(stack.Count - 1);
            return v;
        }

        // Arithmetic / storage operators keep the stack meaningful for later subroutine calls
        private void Escape(byte op)
        {
            switch (op)
            {
                case 3: // and
                case 4: // or
                {
                    var b = Pop();
                    var a = Pop();
                    stack.Add((op == 3 ? (a != 0) && (b != 0) : (a != 0) || (b != 0)) ? 1 : 0);
                    break;
                }

                case 5: // not
                    stack.Add(Pop() == 0 ? 1 : 0);
                    break;
                case 9: // abs
                    stack.Add(Math.Abs(Pop()));
                    break;
                case 10: // add
                    stack.Add(Pop() + Pop());
                    break;
                case 11: // sub
                {
                    var b = Pop();
                    stack.Add(Pop() - b);
                    break;
                }

                case 12: // div
                {
                    var b = Pop();
                    var a = Pop();
                    stack.Add(b == 0 ? 0 : a / b);
                    break;
                }

                case 14: // neg
                    stack.Add(-Pop());
                    break;
                case 15: // eq
                    stack.Add(Pop() == Pop() ? 1 : 0);
                    break;
                case 18: // drop
                    Pop();
                    break;
                case 20: // put
                {
                    var index = (int)Pop();
                    var value = Pop();
                    if ((uint)index < (uint)transient.Length)
                    {
                        transient[index] = value;
                    }

                    break;
                }

                case 21: // get
                {
                    var index = (int)Pop();
                    stack.Add((uint)index < (uint)transient.Length ? transient[index] : 0);
                    break;
                }

                case 22: // ifelse
                {
                    var v2 = Pop();
                    var v1 = Pop();
                    var s2 = Pop();
                    var s1 = Pop();
                    stack.Add(v1 <= v2 ? s1 : s2);
                    break;
                }

                case 23: // random (any value in (0, 1])
                    stack.Add(0.5);
                    break;
                case 24: // mul
                    stack.Add(Pop() * Pop());
                    break;
                case 26: // sqrt
                    stack.Add(Math.Sqrt(Pop()));
                    break;
                case 27: // dup
                {
                    var a = Pop();
                    stack.Add(a);
                    stack.Add(a);
                    break;
                }

                case 28: // exch
                {
                    var b = Pop();
                    var a = Pop();
                    stack.Add(b);
                    stack.Add(a);
                    break;
                }

                case 29: // index (negative index copies the top element)
                {
                    var index = Math.Max((int)Pop(), 0);
                    stack.Add(stack.Count == 0 ? 0 : stack[Math.Max(stack.Count - 1 - index, 0)]);
                    break;
                }

                case 30: // roll
                {
                    var shift = (int)Pop();
                    var count = (int)Pop();
                    if ((count > 0) && (count <= stack.Count))
                    {
                        var items = stack.GetRange(stack.Count - count, count);
                        stack.RemoveRange(stack.Count - count, count);
                        for (var k = 0; k < count; k++)
                        {
                            stack.Add(items[(((k - shift) % count) + count) % count]);
                        }
                    }

                    break;
                }

                default: // flex family and reserved operators clear the stack
                    stack.Clear();
                    break;
            }
        }
    }
}

internal interface IFontSubset
{
    int IncludeGlyph(int oldGlyphId);

    byte[] Encode();
}
