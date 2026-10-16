// ReSharper disable UseUtf8StringLiteral
namespace OysterReport.Tests;

using System.Buffers.Binary;
using System.IO.Compression;

using OysterReport.Internal.Pdf;

using UglyToad.PdfPig;

public sealed class PdfInfrastructureTests
{
    private static readonly Lazy<byte[]> IpaExGothic = new(static () => File.ReadAllBytes(TestHelper.IpaExGothicFontPath));

    //--------------------------------------------------------------------------------
    // TrueType
    //--------------------------------------------------------------------------------

    [Fact]
    public void TrueTypeFontShouldLoadMetrics()
    {
        // Act
        var font = TrueTypeFont.Load(IpaExGothic.Value);

        // Assert
        Assert.False(font.IsCompactFont);
        Assert.Equal(2048, font.UnitsPerEm);
        Assert.NotEqual(0, font.GetGlyphId('日'));
        Assert.True(font.CellSpace >= font.CellAscent + font.CellDescent);
    }

    [Fact]
    public void TrueTypeFontShouldSelectFaceFromCollection()
    {
        // Arrange
        var ttc = BuildCollection(IpaExGothic.Value, 2);
        var standalone = TrueTypeFont.Load(IpaExGothic.Value);

        // Act
        var face = TrueTypeFont.Load(ttc, 1);

        // Assert
        Assert.Equal(standalone.GetGlyphId('本'), face.GetGlyphId('本'));
        Assert.Equal(standalone.GetAdvanceWidth(standalone.GetGlyphId('本')), face.GetAdvanceWidth(face.GetGlyphId('本')));
    }

    [Fact]
    public void TrueTypeFontShouldThrowWhenFaceIndexIsOutOfRange()
    {
        // Arrange
        var ttc = BuildCollection(IpaExGothic.Value, 1);

        // Act / Assert
        Assert.Throws<InvalidDataException>(() => TrueTypeFont.Load(ttc, 1));
    }

    [Fact]
    public void TrueTypeSubsetShouldKeepOnlyUsedGlyphs()
    {
        // Arrange
        var font = TrueTypeFont.Load(IpaExGothic.Value);
        var subset = new TrueTypeSubset(font);
        var hi = subset.IncludeGlyph(font.GetGlyphId('日'));
        var hon = subset.IncludeGlyph(font.GetGlyphId('本'));

        // Act
        var subsetFont = TrueTypeFont.Load(subset.Encode());

        // Assert
        Assert.Equal(3, subsetFont.NumGlyphs);
        Assert.Equal(font.GetAdvanceWidth(font.GetGlyphId('日')), subsetFont.GetAdvanceWidth(hi));
        Assert.Equal(font.GetAdvanceWidth(font.GetGlyphId('本')), subsetFont.GetAdvanceWidth(hon));
    }

    //--------------------------------------------------------------------------------
    // Compact Font
    //--------------------------------------------------------------------------------

    [Fact]
    public void CompactFontSubsetShouldInlineSubroutines()
    {
        // Arrange
        var font = TrueTypeFont.Load(File.ReadAllBytes(AssetPath("OysterTestCompactFontCid.otf")));
        var subset = new CompactFontSubset(font.CompactFont!);
        subset.IncludeGlyph(font.GetGlyphId('A'));
        subset.IncludeGlyph(font.GetGlyphId('本'));

        // Act
        var bytes = subset.Encode();
        var result = new CompactFont(bytes, 0, bytes.Length);

        // Assert
        Assert.True(result.IsCidKeyed);
        Assert.Equal(3, result.CharStrings.Count);
        Assert.Equal(0, result.GlobalSubroutines.Count);
        Assert.All(result.FontDictionaries, static dictionary => Assert.Null(dictionary.Subroutines));

        // "A" calls gsubr 0 and local subr 0; both bodies are inlined (no operand byte of this font equals 10 or 29)
        var a = result.Get(result.CharStrings, 1).ToArray();
        Assert.Equal(14, a[^1]);
        Assert.DoesNotContain((byte)10, a);
        Assert.DoesNotContain((byte)29, a);
    }

    [Fact]
    public void CompactFontSubsetShouldConvertNameKeyedFontToCidKeyed()
    {
        // Arrange
        var font = TrueTypeFont.Load(File.ReadAllBytes(AssetPath("OysterTestCompactFont.otf")));
        var source = font.CompactFont!;
        var subset = new CompactFontSubset(source, inlineSubroutines: false);
        subset.IncludeGlyph(font.GetGlyphId('A'));
        subset.IncludeGlyph(font.GetGlyphId('日'));

        // Act
        var bytes = subset.Encode();
        var result = new CompactFont(bytes, 0, bytes.Length);

        // Assert
        Assert.False(source.IsCidKeyed);
        Assert.True(result.IsCidKeyed);
        Assert.Equal(3, result.CharStrings.Count);
        Assert.Single(result.FontDictionaries);
        Assert.Equal(source.Get(source.CharStrings, font.GetGlyphId('日')).ToArray(), result.Get(result.CharStrings, 2).ToArray());

        // A uses global subr 0 and local subr 0, 日 uses local subr 0; the others become "return"
        Assert.Equal(source.Get(source.GlobalSubroutines, 0).ToArray(), result.Get(result.GlobalSubroutines, 0).ToArray());
        Assert.Equal(new byte[] { 11 }, result.Get(result.GlobalSubroutines, 1).ToArray());
        var subroutines = result.FontDictionaries[0].Subroutines!;
        Assert.Equal(source.Get(source.FontDictionaries[0].Subroutines!, 0).ToArray(), result.Get(subroutines, 0).ToArray());
        Assert.Equal(new byte[] { 11 }, result.Get(subroutines, 1).ToArray());
    }

    [Fact]
    public void CompactFontSubsetShouldRemapFontDictionariesOfCidKeyedFont()
    {
        // Arrange
        var font = TrueTypeFont.Load(File.ReadAllBytes(AssetPath("OysterTestCompactFontCid.otf")));
        var source = font.CompactFont!;
        var subset = new CompactFontSubset(source, inlineSubroutines: false);
        subset.IncludeGlyph(font.GetGlyphId('本'));
        subset.IncludeGlyph(font.GetGlyphId('A'));

        // Act
        var bytes = subset.Encode();
        var result = new CompactFont(bytes, 0, bytes.Length);

        // Assert
        Assert.True(source.IsCidKeyed);
        Assert.Equal(2, source.FontDictionaries.Count);
        Assert.Equal(2, result.FontDictionaries.Count);
        Assert.Equal(new byte[] { 0, 1, 0 }, result.FontDictionarySelection);
        Assert.Equal(source.Get(source.CharStrings, font.GetGlyphId('本')).ToArray(), result.Get(result.CharStrings, 1).ToArray());

        // 本 (FD 1 in the source, FD 1 here since .notdef keeps FD 0) calls local subr 1 and global subr 0
        Assert.Equal(source.Get(source.FontDictionaries[1].Subroutines!, 1).ToArray(), result.Get(result.FontDictionaries[1].Subroutines!, 1).ToArray());
        Assert.Equal(new byte[] { 11 }, result.Get(result.FontDictionaries[1].Subroutines!, 0).ToArray());
    }

    [Theory]
    [InlineData("OysterTestCompactFont.otf")]
    [InlineData("OysterTestCompactFontCid.otf")]
    public void PdfDocumentWriterShouldEmbedCompactFontSubset(string fileName)
    {
        // Act
        var pdf = WritePdf(compress: true, font: File.ReadAllBytes(AssetPath(fileName)), draw: static (page, font) =>
            page.ShowText(font, 12, PdfColor.Black, 20, 60, "AB日本", 0));
        TestHelper.SavePdf($"{nameof(PdfDocumentWriterShouldEmbedCompactFontSubset)}-{Path.GetFileNameWithoutExtension(fileName)}", pdf);

        // Assert
        using var document = PdfDocument.Open(pdf);
        Assert.Equal("AB日本", String.Concat(document.GetPage(1).Letters.Select(static x => x.Value)));
        Assert.Contains("/CIDFontType0C", Encoding.Latin1.GetString(pdf), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Writer
    //--------------------------------------------------------------------------------

    [Fact]
    public void PdfDocumentWriterShouldEmbedSubsetFontWithToUnicode()
    {
        // Act
        var pdf = WritePdf(compress: true, draw: static (page, font) =>
            page.ShowText(font, 12, PdfColor.Black, 40, 60, "日本語 目次", 0));

        // Assert
        using var document = PdfDocument.Open(pdf);
        var letters = document.GetPage(1).Letters;
        var text = String.Concat(letters.Select(static x => x.Value));
        Assert.Equal("日本語 目次", text);
        Assert.Contains(letters, static x => (x.FontName is not null) && x.FontName.EndsWith("+IPAexGothic", StringComparison.Ordinal));
        Assert.True(pdf.Length < 20000);
    }

    [Fact]
    public void PdfDocumentWriterShouldWritePlainContentWhenCompressionIsDisabled()
    {
        // Act
        var pdf = WritePdf(compress: false, draw: static (page, font) =>
            page.ShowText(font, 12, PdfColor.Black, 40, 60, "ABC", 0));

        // Assert
        Assert.Contains("BT /F1 12 Tf", Encoding.Latin1.GetString(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentWriterShouldNotEmbedUnusedFont()
    {
        // Act
        var pdf = WritePdf(compress: false, draw: static (page, _) =>
            page.FillRect(PdfColor.Black, 10, 10, 10, 10));

        // Assert
        Assert.DoesNotContain("/FontFile2", Encoding.Latin1.GetString(pdf), StringComparison.Ordinal);
        Assert.Equal(1, TestHelper.GetPageCount(pdf));
    }

    //--------------------------------------------------------------------------------
    // Images
    //--------------------------------------------------------------------------------

    [Fact]
    public void PngDecoderShouldDecodeInterlacedRgba()
    {
        // Arrange
        static byte[] Pixel(int x, int y) => [(byte)(x * 40), (byte)(y * 40), 7, (byte)(255 - (x * 10))];
        var png = BuildPng(5, 4, colorType: 6, depth: 8, interlaced: true, Pixel);

        // Act
        var image = PngDecoder.Decode(png);

        // Assert
        Assert.Equal(3, image.Colors);
        Assert.NotNull(image.Alpha);
        for (var y = 0; y < 4; y++)
        {
            for (var x = 0; x < 5; x++)
            {
                var i = (y * 5) + x;
                var expected = Pixel(x, y);
                Assert.Equal(expected[..3], image.Color.AsSpan(i * 3, 3).ToArray());
                Assert.Equal(expected[3], image.Alpha[i]);
            }
        }
    }

    [Fact]
    public void PngDecoderShouldDownconvertSixteenBitSamples()
    {
        // Arrange
        var png = BuildPng(1, 1, colorType: 0, depth: 16, interlaced: false, static (_, _) => [0xAB, 0xCD]);

        // Act
        var image = PngDecoder.Decode(png);

        // Assert
        Assert.Equal(1, image.Colors);
        Assert.Equal(0xAB, image.Color[0]);
        Assert.Null(image.Alpha);
    }

    [Fact]
    public void PngDecoderShouldApplyPaletteTransparency()
    {
        // Arrange
        var png = BuildPng(2, 1, colorType: 3, depth: 8, interlaced: false, static (x, _) => [(byte)x], palette: [255, 0, 0, 0, 0, 255], transparency: [0]);

        // Act
        var image = PngDecoder.Decode(png);

        // Assert
        Assert.Equal(new byte[] { 255, 0, 0, 0, 0, 255 }, image.Color);
        Assert.Equal(new byte[] { 0, 255 }, image.Alpha);
    }

    [Fact]
    public void BmpDecoderShouldDecodeTrueColorAndPalette()
    {
        // Arrange
        var bmp24 = Convert.FromBase64String("Qk1OAAAAAAAAADYAAAAoAAAAAwAAAAIAAAABABgAAAAAABgAAADEDgAAxA4AAAAAAAAAAAAAAP////8A/wD/AAAAAAD/AP8A/wAAAAAA");
        var bmp8 = Convert.FromBase64String("Qk1WAAAAAAAAAE4AAAAoAAAAAwAAAAIAAAABAAgAAAAAAAgAAADEDgAAxA4AAAYAAAAGAAAAAP//AP//AAAA/wAA/wD/AAAA/wD/AAAAAAEDAAQCBQA=");
        byte[] expected = [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0, 0, 255, 255, 255, 0, 255];

        // Act
        var trueColor = BmpDecoder.Decode(bmp24);
        var palette = BmpDecoder.Decode(bmp8);

        // Assert
        Assert.Equal(expected, trueColor.Color);
        Assert.Equal(expected, palette.Color);
    }

    [Fact]
    public void GifDecoderShouldDecodeTransparency()
    {
        // Arrange
        var gif = Convert.FromBase64String("R0lGODlhAwACAIIAAP//AAD//wD/AP8A//8AAAAA/wAAAAAAACH5BAEAAAQALAAAAAADAAIAAAgJAAkIKAAgwICAADs=");

        // Act
        var image = GifDecoder.Decode(gif);

        // Assert
        Assert.Equal(new byte[] { 255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0, 0, 255, 255, 255, 0, 255 }, image.Color);
        Assert.Equal(new byte[] { 0, 255, 255, 255, 255, 255 }, image.Alpha);
    }

    [Fact]
    public void PdfImageWriterShouldRejectUnknownFormat()
    {
        // Arrange
        using var output = new MemoryStream();
        var writer = new PdfDocumentWriter(output, compress: true);

        // Act / Assert
        Assert.Throws<NotSupportedException>(() => PdfImageWriter.Write(writer, "RIFF0000WEBP"u8));
    }

    //--------------------------------------------------------------------------------
    // System fonts
    //--------------------------------------------------------------------------------

    [Fact]
    public void SystemFontCatalogShouldResolveFamilyFromDirectory()
    {
        // Arrange
        var directory = Directory.CreateTempSubdirectory("oyster-fonts-");
        try
        {
            File.Copy(TestHelper.IpaExGothicFontPath, Path.Combine(directory.FullName, "ipaexg.ttf"));
            var catalog = new SystemFontCatalog([directory.FullName]);

            // Act
            var regular = catalog.TryResolve("IPAexGothic", false, false, out var face, out var simulateBold, out _);
            var bold = catalog.TryResolve("IPAexGothic", true, false, out _, out var simulateBoldForBold, out _);
            var unknown = catalog.TryResolve("__NoSuchFont__", false, false, out _, out _, out _);

            // Assert
            Assert.True(regular);
            Assert.Equal(0, face.FaceIndex);
            Assert.False(simulateBold);
            Assert.True(bold);
            Assert.True(simulateBoldForBold);
            Assert.False(unknown);
            Assert.Equal(1, catalog.FaceCount);
        }
        finally
        {
            directory.Delete(true);
        }
    }

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private static string AssetPath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Assets", fileName);

    private static byte[] WritePdf(bool compress, Action<PdfPageCanvas, PdfEmbeddedFont> draw, byte[]? font = null)
    {
        using var output = new MemoryStream();
        var writer = new PdfDocumentWriter(output, compress);
        var embedded = writer.AddFont(TrueTypeFont.Load(font ?? IpaExGothic.Value));
        var page = writer.BeginPage(200, 100);
        draw(page, embedded);
        writer.EndPage(page);
        writer.Close();
        return output.ToArray();
    }

    // Wraps a single sfnt into a collection with the given number of faces sharing the same tables
    private static byte[] BuildCollection(byte[] sfnt, int faces)
    {
        int numTables = BinaryPrimitives.ReadUInt16BigEndian(sfnt.AsSpan(4));
        var directorySize = 12 + (numTables * 16);
        var headerSize = 12 + (faces * 4);
        var shift = headerSize + (faces * directorySize) - directorySize;
        var result = new byte[headerSize + (faces * directorySize) + sfnt.Length - directorySize];
        "ttcf"u8.CopyTo(result);
        BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(4), 0x00010000);
        BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(8), (uint)faces);
        for (var face = 0; face < faces; face++)
        {
            var faceOffset = headerSize + (face * directorySize);
            BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(12 + (face * 4)), (uint)faceOffset);
            sfnt.AsSpan(0, directorySize).CopyTo(result.AsSpan(faceOffset));
            for (var table = 0; table < numTables; table++)
            {
                var record = result.AsSpan(faceOffset + 12 + (table * 16) + 8);
                BinaryPrimitives.WriteUInt32BigEndian(record, BinaryPrimitives.ReadUInt32BigEndian(record) + (uint)shift);
            }
        }

        sfnt.AsSpan(directorySize).CopyTo(result.AsSpan(headerSize + (faces * directorySize)));
        return result;
    }

    // Minimal PNG encoder (filter 0, optional Adam7); CRCs are not validated by the decoder
    private static byte[] BuildPng(int width, int height, int colorType, int depth, bool interlaced, Func<int, int, byte[]> pixel, byte[]? palette = null, byte[]? transparency = null)
    {
        using var raw = new MemoryStream();
        void Rows(IEnumerable<int> xs, IEnumerable<int> ys)
        {
            var columns = xs.ToList();
            foreach (var y in ys)
            {
                raw.WriteByte(0);
                foreach (var x in columns)
                {
                    raw.Write(pixel(x, y));
                }
            }
        }

        if (interlaced)
        {
            foreach (var (x0, y0, dx, dy) in new[] { (0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2) })
            {
                var xs = Enumerable.Range(0, width).Where(x => (x >= x0) && ((x - x0) % dx == 0)).ToList();
                var ys = Enumerable.Range(0, height).Where(y => (y >= y0) && ((y - y0) % dy == 0)).ToList();
                if ((xs.Count > 0) && (ys.Count > 0))
                {
                    Rows(xs, ys);
                }
            }
        }
        else
        {
            Rows(Enumerable.Range(0, width), Enumerable.Range(0, height));
        }

        using var compressed = new MemoryStream();
        using (var z = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
        {
            raw.Position = 0;
            raw.CopyTo(z);
        }

        using var png = new MemoryStream();
        png.Write([0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A]);

        var header = new byte[13];
        BinaryPrimitives.WriteUInt32BigEndian(header, (uint)width);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), (uint)height);
        header[8] = (byte)depth;
        header[9] = (byte)colorType;
        header[12] = interlaced ? (byte)1 : (byte)0;
        Chunk("IHDR", header);
        if (palette is not null)
        {
            Chunk("PLTE", palette);
        }

        if (transparency is not null)
        {
            Chunk("tRNS", transparency);
        }

        Chunk("IDAT", compressed.ToArray());
        Chunk("IEND", []);
        return png.ToArray();

        void Chunk(string type, ReadOnlySpan<byte> data)
        {
            Span<byte> length = stackalloc byte[4];
            BinaryPrimitives.WriteUInt32BigEndian(length, (uint)data.Length);
            png.Write(length);
            png.Write(Encoding.ASCII.GetBytes(type));
            png.Write(data);
#pragma warning disable IDE0230
            png.Write([0, 0, 0, 0]);
#pragma warning restore IDE0230
        }
    }
}
