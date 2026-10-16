namespace OysterReport.Internal.Pdf;

internal sealed class CompactFontSubset(CompactFont compactFont, bool inlineSubroutines = true)
    : IFontSubset
{
    private const int MaxSubrDepth = 10;

    private static readonly byte[] ReturnOnly = [11];

    private static readonly HashSet<int> DroppedTopOperators =
    [
        CompactFontOperators.RegistryOrderingSupplement, CompactFontOperators.Charset, CompactFontOperators.Encoding, CompactFontOperators.CharStrings, CompactFontOperators.Private,
        CompactFontOperators.FontDictionaryArray, CompactFontOperators.FontDictionarySelection, CompactFontOperators.CidCount, CompactFontOperators.UidBase, CompactFontOperators.UniqueId,
        CompactFontOperators.Xuid, CompactFontOperators.SyntheticBase
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
        // Font dictionary in first-use order and the subroutines each glyph actually calls
        var fdMap = new Dictionary<int, int>();
        var newFontDictionarySelection = new byte[glyphs.Count];
        var usedGlobal = new HashSet<int>();
        var usedLocal = new List<HashSet<int>>();
        for (var i = 0; i < glyphs.Count; i++)
        {
            var sourceFontDictionaryIndex = compactFont.FontDictionarySelection[glyphs[i]];
            if (!fdMap.TryGetValue(sourceFontDictionaryIndex, out var fontDictionaryIndex))
            {
                fontDictionaryIndex = fdMap.Count;
                fdMap[sourceFontDictionaryIndex] = fontDictionaryIndex;
                usedLocal.Add([]);
            }

            newFontDictionarySelection[i] = (byte)fontDictionaryIndex;
            var scanner = new CharStringScanner(compactFont, compactFont.FontDictionaries[sourceFontDictionaryIndex].Subroutines, usedGlobal, usedLocal[fontDictionaryIndex]);
            scanner.Scan(compactFont.Get(compactFont.CharStrings, glyphs[i]), 0);
        }

        var fontDictionaries = fdMap.OrderBy(static entry => entry.Value).Select(entry => compactFont.FontDictionaries[entry.Key]).ToList();
        var flattened = inlineSubroutines ? Flatten() : null;

        // Strings: keep the original String INDEX (SIDs stay valid) and append the ROS strings
        var strings = Enumerable.Range(0, compactFont.Strings.Count).Select(i => compactFont.Get(compactFont.Strings, i).ToArray()).ToList();
        var sidAdobe = CompactFont.StandardStringCount + strings.Count;
        strings.Add([.. "Adobe"u8]);
        strings.Add([.. "Identity"u8]);

        var nameIndex = WriteIndex([[.. compactFont.Get(compactFont.Names, 0)]]);
        var stringIndex = WriteIndex(strings);
        var globalSubroutineIndex = WriteIndex(flattened is null ? SubsetSubroutines(compactFont.GlobalSubroutines, usedGlobal) : []);
        var charStringIndex = WriteIndex(flattened ?? [.. glyphs.Select(glyph => compactFont.Get(compactFont.CharStrings, glyph).ToArray())]);
        var charset = glyphs.Count > 1
            ? new byte[] { 2, 0, 1, (byte)((glyphs.Count - 2) >> 8), (byte)(glyphs.Count - 2) }
            : [0];
        var fontDictionarySelection = new byte[1 + newFontDictionarySelection.Length];
        newFontDictionarySelection.CopyTo(fontDictionarySelection, 1);

        // Private dictionary (+ local subroutines immediately after it, so the subroutine offset equals the dictionary size) for each font dictionary
        var privateDictionaries = new List<(byte[] Dictionary, byte[] Subroutines)>();
        for (var i = 0; i < fontDictionaries.Count; i++)
        {
            var entries = fontDictionaries[i].Private.Where(static entry => entry.Operator != CompactFontOperators.Subroutines).ToList();
            byte[] subroutines = [];
            if ((flattened is null) && (fontDictionaries[i].Subroutines is { } localSubroutines))
            {
                subroutines = WriteIndex(SubsetSubroutines(localSubroutines, usedLocal[i]));
                var dictionarySize = CompactFontDictionaryEntry.Encode(entries).Length + 5 + 1;
                entries.Add(CompactFontDictionaryEntry.Integers(CompactFontOperators.Subroutines, dictionarySize));
            }

            privateDictionaries.Add((CompactFontDictionaryEntry.Encode(entries), subroutines));
        }

        // Layout: header, name, top dictionary, string, global subroutines, charset, font dictionary selection, character strings, font dictionary array, private dictionaries
        var topDictionary = BuildTopDictionary(sidAdobe, 0, 0, 0, 0);
        var topDictionaryIndex = WriteIndex([CompactFontDictionaryEntry.Encode(topDictionary)]);
        var offset = 4 + nameIndex.Length + topDictionaryIndex.Length + stringIndex.Length + globalSubroutineIndex.Length;
        var charsetOffset = offset;
        var fdSelectOffset = charsetOffset + charset.Length;
        var charStringsOffset = fdSelectOffset + fontDictionarySelection.Length;
        var fdArrayOffset = charStringsOffset + charStringIndex.Length;

        var privateOffsets = new List<int>();
        var fontDictionaryArrayIndex = WriteIndex([.. fontDictionaries.Select((dictionary, i) => BuildFontDictionary(dictionary, privateDictionaries[i].Dictionary.Length, 0))]);
        var privateOffset = fdArrayOffset + fontDictionaryArrayIndex.Length;
        for (var i = 0; i < fontDictionaries.Count; i++)
        {
            privateOffsets.Add(privateOffset);
            privateOffset += privateDictionaries[i].Dictionary.Length + privateDictionaries[i].Subroutines.Length;
        }

        fontDictionaryArrayIndex = WriteIndex([.. fontDictionaries.Select((dictionary, i) => BuildFontDictionary(dictionary, privateDictionaries[i].Dictionary.Length, privateOffsets[i]))]);
        topDictionaryIndex = WriteIndex([CompactFontDictionaryEntry.Encode(BuildTopDictionary(sidAdobe, charsetOffset, fdSelectOffset, charStringsOffset, fdArrayOffset))]);

        using var output = new MemoryStream();
        output.Write([1, 0, 4, 4]);
        output.Write(nameIndex);
        output.Write(topDictionaryIndex);
        output.Write(stringIndex);
        output.Write(globalSubroutineIndex);
        output.Write(charset);
        output.Write(fontDictionarySelection);
        output.Write(charStringIndex);
        output.Write(fontDictionaryArrayIndex);
        foreach (var (dictionary, subroutines) in privateDictionaries)
        {
            output.Write(dictionary);
            output.Write(subroutines);
        }

        return output.ToArray();
    }

    private static int Bias(int count) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768;

    private List<byte[]>? Flatten()
    {
        var result = new List<byte[]>(glyphs.Count);
        foreach (var gid in glyphs)
        {
            var flattener = new CharStringFlattener(compactFont, compactFont.FontDictionaries[compactFont.FontDictionarySelection[gid]].Subroutines);
            if (flattener.Flatten(compactFont.Get(compactFont.CharStrings, gid)) is not { } code)
            {
                return null;
            }

            result.Add(code);
        }

        return result;
    }

    private List<CompactFontDictionaryEntry> BuildTopDictionary(int sidAdobe, int charsetOffset, int fdSelectOffset, int charStringsOffset, int fdArrayOffset)
    {
        // The registry-ordering-supplement operator must be the first operator of the CIDFont top dictionary.
        List<CompactFontDictionaryEntry> dictionary =
        [
            CompactFontDictionaryEntry.Integers(CompactFontOperators.RegistryOrderingSupplement, sidAdobe, sidAdobe + 1, 0), .. compactFont.TopDictionary.Where(static entry => !DroppedTopOperators.Contains(entry.Operator)),
            CompactFontDictionaryEntry.Integers(CompactFontOperators.CidCount, glyphs.Count),
            CompactFontDictionaryEntry.Integers(CompactFontOperators.Charset, charsetOffset),
            CompactFontDictionaryEntry.Integers(CompactFontOperators.FontDictionarySelection, fdSelectOffset),
            CompactFontDictionaryEntry.Integers(CompactFontOperators.CharStrings, charStringsOffset),
            CompactFontDictionaryEntry.Integers(CompactFontOperators.FontDictionaryArray, fdArrayOffset)
        ];
        return dictionary;
    }

    private byte[] BuildFontDictionary(CompactFontDictionary fontDictionary, int privateSize, int privateOffset)
    {
        // A name-keyed source has no font dictionary entries of its own (its top dictionary is kept as the CIDFont top dictionary).
        var entries = compactFont.IsCidKeyed ? fontDictionary.Entries.Where(static entry => entry.Operator != CompactFontOperators.FontName).ToList() : [];
        entries.Add(CompactFontDictionaryEntry.Integers(CompactFontOperators.Private, privateSize, privateOffset));
        return CompactFontDictionaryEntry.Encode(entries);
    }

    // Unused subroutines are replaced with a bare "return" so the remaining indices stay valid
    private List<byte[]> SubsetSubroutines(CompactFontIndex subroutines, HashSet<int> used) =>
    [
        .. Enumerable.Range(0, subroutines.Count).Select(i => used.Contains(i) ? [.. compactFont.Get(subroutines, i)] : ReturnOnly)
    ];

    private static byte[] WriteIndex(List<byte[]> items)
    {
        if (items.Count == 0)
        {
#pragma warning disable IDE0230
            return [0, 0];
#pragma warning restore IDE0230
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
    private sealed class CharStringFlattener(CompactFont compactFont, CompactFontIndex? localSubroutines)
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
                    var subroutines = b0 == 10 ? localSubroutines : compactFont.GlobalSubroutines;
                    if ((lastNumberStart < 0) || (subroutines is null))
                    {
                        return Result.Failed;
                    }

                    // Drop the index operand; the subroutine body replaces the call
                    output.RemoveRange(lastNumberStart, output.Count - lastNumberStart);
                    depth--;
                    lastNumberStart = -1;
                    var index = (int)lastNumber + Bias(subroutines.Count);
                    if ((uint)index >= (uint)subroutines.Count)
                    {
                        return Result.Failed;
                    }

                    var result = Inline(compactFont.Get(subroutines, index), nesting + 1);
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
    private sealed class CharStringScanner(CompactFont compactFont, CompactFontIndex? localSubroutines, HashSet<int> usedGlobal, HashSet<int> usedLocal)
    {
        private readonly List<double> stack = [];
        private readonly double[] transient = new double[32];
        private int stems;

        // Returns true when endchar was reached
        public bool Scan(ReadOnlySpan<byte> code, int depth)
        {
            if (depth > MaxSubrDepth)
            {
                throw new InvalidDataException("Compact Font subroutine nesting is too deep.");
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
                        if (localSubroutines is not null && CallSubroutine(localSubroutines, usedLocal, depth))
                        {
                            return true;
                        }

                        break;
                    case 29: // callgsubr
                        if (CallSubroutine(compactFont.GlobalSubroutines, usedGlobal, depth))
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

        private bool CallSubroutine(CompactFontIndex subroutines, HashSet<int> used, int depth)
        {
            if (stack.Count == 0)
            {
                return false;
            }

            var index = (int)Pop() + Bias(subroutines.Count);
            if ((uint)index >= (uint)subroutines.Count)
            {
                return false;
            }

            used.Add(index);
            return Scan(compactFont.Get(subroutines, index), depth + 1);
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
                // ReSharper disable once IntVariableOverflowInUncheckedContext
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
                    // ReSharper disable once CompareOfFloatsByEqualityOperator
                    // ReSharper disable once EqualExpressionComparison
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
