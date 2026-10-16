namespace OysterReport.Tests;

using System.Buffers.Binary;
using System.IO.Compression;

using OysterReport.Internal.Pdf;

using UglyToad.PdfPig;

public sealed class RenderOptionTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    // Not decodable by the built-in decoders (RIFF/WEBP container)
#pragma warning disable IDE0028
    private static readonly byte[] FakeWebp = "RIFF\u0010\0\0\0WEBPVP8 \u0004\0\0\0\0\0\0\0"u8.ToArray();
#pragma warning restore IDE0028

    //--------------------------------------------------------------------------------
    // Image converter
    //--------------------------------------------------------------------------------

    [Fact]
    public void ImageConverterShouldEmbedConvertedImage()
    {
        // Arrange
        var converted = new List<int>();
        var options = new ReportRenderOption
        {
            ImageConverter = bytes =>
            {
                converted.Add(bytes.Length);
                return OnePxPng;
            }
        };
        var warnings = new List<ReportRenderWarning>();
        options.OnRenderWarning = warnings.Add;

        // Act
        var pdf = Render(CreateWorkbookWithImage(FakeWebp), options);

        // Assert
        Assert.Equal(new[] { FakeWebp.Length }, converted);
        Assert.Empty(warnings);
        Assert.Equal(1, TestHelper.GetImageCount(pdf));
    }

    [Fact]
    public void ImageConverterShouldWarnWhenItReturnsNull()
    {
        // Arrange
        var warnings = new List<ReportRenderWarning>();
        var options = new ReportRenderOption { ImageConverter = static _ => null, OnRenderWarning = warnings.Add };

        // Act
        var pdf = Render(CreateWorkbookWithImage(FakeWebp), options);

        // Assert
        Assert.Contains(warnings, static x => x.Kind == ReportRenderWarningKind.ImageDecodeFailed);
        Assert.Equal(0, TestHelper.GetImageCount(pdf));
    }

    //--------------------------------------------------------------------------------
    // Warnings
    //--------------------------------------------------------------------------------

    [Fact]
    public void MissingGlyphShouldRaiseWarning()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var cell = workbook.AddWorksheet("Report").Cell("A1");
            cell.Value = "日本語\U0001F600";
            cell.Style.Font.FontName = "ＭＳ ゴシック";
        });
        var warnings = new List<ReportRenderWarning>();
        var options = new ReportRenderOption { OnRenderWarning = warnings.Add };

        // Act
        Render(stream, options, new TestFontResolver(TestHelper.IpaExGothicFontPath));

        // Assert
        var warning = Assert.Single(warnings, static x => x.Kind == ReportRenderWarningKind.MissingGlyph);
        Assert.Equal("Report", warning.SheetName);
        Assert.Equal("IPAexGothic", warning.Source);
        Assert.Contains("\U0001F600", warning.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("日", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RestrictedFontShouldRaiseWarningAndStillEmbed()
    {
        // Arrange: OS/2 fsType = 0x0002 (Restricted License embedding)
        var font = File.ReadAllBytes(TestHelper.IpaExGothicFontPath);
        BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(FindTable(font, "OS/2") + 8), 0x0002);
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var cell = workbook.AddWorksheet("Report").Cell("A1");
            cell.Value = "埋め込み";
            cell.Style.Font.FontName = "Restricted";
        });
        var warnings = new List<ReportRenderWarning>();
        var options = new ReportRenderOption { OnRenderWarning = warnings.Add };

        // Act
        var pdf = Render(stream, options, new SingleFontResolver("Restricted", font));

        // Assert
        var warning = Assert.Single(warnings, static x => x.Kind == ReportRenderWarningKind.FontEmbeddingRestricted);
        Assert.Contains("restricted license", warning.Message, StringComparison.Ordinal);
        Assert.Contains("埋め込み", TestHelper.ExtractAllText(pdf), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Font collections
    //--------------------------------------------------------------------------------

    [Fact]
    public void FontCollectionFaceShouldBeSelectedByNameOrIndex()
    {
        // Arrange
        var collection = BuildCollection(
            File.ReadAllBytes(AssetPath("OysterTestCompactFont.otf")),
            File.ReadAllBytes(AssetPath("OysterTestCompactFontCid.otf")));

        // Act
        var byName = ResolveFace(collection, "Oyster Test CFF CID", null);
        var byIndex = ResolveFace(collection, "Oyster Test CFF CID", 0);
        var fallback = ResolveFace(collection, "Unknown Face", null);

        // Assert
        Assert.Equal(1, TrueTypeFont.FindFaceIndex(collection, "Oyster Test CFF CID"));
        Assert.True(byName.CompactFont!.IsCidKeyed);
        Assert.False(byIndex.CompactFont!.IsCidKeyed);
        Assert.False(fallback.CompactFont!.IsCidKeyed);
    }

    //--------------------------------------------------------------------------------
    // Font discovery
    //--------------------------------------------------------------------------------

    [Fact]
    public void FontDirectoriesAndAliasesShouldBeUsedWithoutSystemFonts()
    {
        // Arrange
        var directory = CreateFontDirectory();
        try
        {
            var options = new ReportRenderOption
            {
                UseSystemFonts = false,
                FontDirectories = [directory.FullName],
                FontAliases = new Dictionary<string, IReadOnlyList<string>> { ["My Report Font"] = ["IPAexGothic"] },
                HeaderFooterFallbackFonts = ["IPAexGothic"]
            };
            using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                sheet.Cell("A1").Value = "直接指定";
                sheet.Cell("A1").Style.Font.FontName = "IPAexGothic";
                sheet.Cell("A2").Value = "別名";
                sheet.Cell("A2").Style.Font.FontName = "my report font";
            });

            // Act
            var pdf = Render(stream, options);

            // Assert
            var text = TestHelper.ExtractAllText(pdf);
            Assert.Contains("直接指定", text, StringComparison.Ordinal);
            Assert.Contains("別名", text, StringComparison.Ordinal);
        }
        finally
        {
            directory.Delete(true);
        }
    }

    [Fact]
    public void DisabledSystemFontsShouldFailWhenNoFontIsAvailable()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook => workbook.AddWorksheet("Report").Cell("A1").Value = "Text");
        var options = new ReportRenderOption { UseSystemFonts = false };

        // Act / Assert
        Assert.Throws<InvalidOperationException>(() => Render(stream, options));
    }

    [Fact]
    public void TemplateWorkbookShouldMeasureColumnsWithConfiguredFonts()
    {
        // Arrange
        var directory = CreateFontDirectory();
        try
        {
            // The workbook default font is Calibri (11pt); it is mapped to IPAexGothic in the configured directory
            using var stream = TestWorkbookFactory.CreateWorkbook(static workbook => workbook.AddWorksheet("Report").Cell("A1").Value = "Text");
            var bytes = stream.ToArray();
            var aliases = new Dictionary<string, IReadOnlyList<string>> { ["Calibri"] = ["IPAexGothic"] };

            // Act
            using var configured = new TemplateWorkbook(new MemoryStream(bytes), new ReportRenderOption { UseSystemFonts = false, FontDirectories = [directory.FullName], FontAliases = aliases });
            using var unavailable = new TemplateWorkbook(new MemoryStream(bytes), new ReportRenderOption { UseSystemFonts = false });

            // Assert: 11pt IPAexGothic digits are 9 px at 96 dpi; without any font the 7 px fallback is used
            Assert.Equal(9d, configured.ReportWorkbook.MeasurementProfile.MaxDigitWidth);
            Assert.Equal(7d, unavailable.ReportWorkbook.MeasurementProfile.MaxDigitWidth);
        }
        finally
        {
            directory.Delete(true);
        }
    }

    //--------------------------------------------------------------------------------
    // Document
    //--------------------------------------------------------------------------------

    [Fact]
    public void MetadataShouldIncludeCreationDateAndFileIdentifier()
    {
        // Arrange
        var bytes = CreateTextWorkbook();

        // Act
        var pdf = Render(new MemoryStream(bytes), new ReportRenderOption(), new TestFontResolver(TestHelper.IpaExGothicFontPath));

        // Assert
        using var document = PdfDocument.Open(pdf);
        Assert.False(String.IsNullOrEmpty(document.Information.CreationDate));
        Assert.Contains("/ID [<", Encoding.Latin1.GetString(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void OutputWithoutMetadataShouldBeDeterministic()
    {
        // Arrange
        var bytes = CreateTextWorkbook();

        // Act
        var first = Render(new MemoryStream(bytes), new ReportRenderOption(), new TestFontResolver(TestHelper.IpaExGothicFontPath), embedMetadata: false);
        var second = Render(new MemoryStream(bytes), new ReportRenderOption(), new TestFontResolver(TestHelper.IpaExGothicFontPath), embedMetadata: false);

        // Assert
        Assert.Equal(first, second);
        using var document = PdfDocument.Open(first);
        Assert.True(String.IsNullOrEmpty(document.Information.CreationDate));
        Assert.Contains("/ID [<", Encoding.Latin1.GetString(first), StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratePdfShouldWriteToNonSeekableStream()
    {
        // Arrange
        using var template = new TemplateWorkbook(new MemoryStream(CreateTextWorkbook()));
        var engine = new OysterReportEngine { FontResolver = new TestFontResolver(TestHelper.IpaExGothicFontPath) };
        using var buffer = new MemoryStream();
        using var output = new ForwardOnlyStream(buffer);

        // Act
        engine.GeneratePdf(template, output);

        // Assert
        Assert.True(TestHelper.IsValidPdf(buffer.ToArray()));
        Assert.Contains("帳票", TestHelper.ExtractAllText(buffer.ToArray()), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private static string AssetPath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Assets", fileName);

    private static byte[] Render(MemoryStream workbookStream, ReportRenderOption options, IReportFontResolver? resolver = null, bool embedMetadata = true)
    {
        workbookStream.Position = 0;
        using var template = new TemplateWorkbook(workbookStream, options);
        var engine = new OysterReportEngine { FontResolver = resolver, RenderingOptions = options, EmbedDocumentMetadata = embedMetadata };
        using var output = new MemoryStream();
        engine.GeneratePdf(template, output);
        return output.ToArray();
    }

    private static byte[] CreateTextWorkbook()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var cell = workbook.AddWorksheet("Report").Cell("A1");
            cell.Value = "帳票";
            cell.Style.Font.FontName = "ＭＳ ゴシック";
        });
        return stream.ToArray();
    }

    // Workbook whose picture part holds the given bytes (ClosedXML only accepts known formats, so the part is replaced)
    private static MemoryStream CreateWorkbookWithImage(byte[] imageBytes)
    {
        var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Image";
            sheet.Cell("D6").Value = "End";
            using var image = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(image, XLPictureFormat.Png, "Picture").MoveTo(sheet.Cell("B2")).WithSize(40, 30);
        });

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.First(static x => x.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
            using var entryStream = entry.Open();
            entryStream.SetLength(0);
            entryStream.Write(imageBytes);
        }

        stream.Position = 0;
        return stream;
    }

    private static DirectoryInfo CreateFontDirectory()
    {
        var directory = Directory.CreateTempSubdirectory("oyster-fonts-");
        File.Copy(TestHelper.IpaExGothicFontPath, Path.Combine(directory.FullName, "ipaexg.ttf"));
        return directory;
    }

    private static TrueTypeFont ResolveFace(byte[] collection, string faceName, int? faceIndex)
    {
        using var output = new MemoryStream();
        var writer = new PdfDocumentWriter(output, compress: true);
        var manager = new PdfFontManager(writer, new SingleFontResolver("Requested", collection, faceName, faceIndex), FontLocator.Create(new ReportRenderOption { UseSystemFonts = false }));
        return manager.Resolve(new ReportFont { Name = "Requested" }, 11).Face.Font;
    }

    private static int FindTable(byte[] font, string tag)
    {
        int numTables = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
        for (var i = 0; i < numTables; i++)
        {
            var record = 12 + (i * 16);
            if (Encoding.ASCII.GetString(font, record, 4) == tag)
            {
                return (int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(record + 8));
            }
        }

        throw new InvalidDataException(tag);
    }

    // Combines standalone sfnt fonts into a collection (each face keeps its own tables)
    private static byte[] BuildCollection(params byte[][] fonts)
    {
        var headerSize = 12 + (fonts.Length * 4);
        using var result = new MemoryStream();
        result.Write(new byte[headerSize]);
        var directoryOffsets = new List<long>();
        foreach (var font in fonts)
        {
            directoryOffsets.Add(result.Position);
            int numTables = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
            result.Write(font.AsSpan(0, 12 + (numTables * 16)));
        }

        for (var face = 0; face < fonts.Length; face++)
        {
            var font = fonts[face];
            int numTables = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
            var directorySize = 12 + (numTables * 16);
            var shift = (int)result.Position - directorySize;
            result.Write(font.AsSpan(directorySize));
            var buffer = result.GetBuffer();
            for (var table = 0; table < numTables; table++)
            {
                var record = buffer.AsSpan((int)directoryOffsets[face] + 12 + (table * 16) + 8);
                BinaryPrimitives.WriteUInt32BigEndian(record, BinaryPrimitives.ReadUInt32BigEndian(record) + (uint)shift);
            }
        }

        var output = result.ToArray();
        "ttcf"u8.CopyTo(output);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(4), 0x00010000);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(8), (uint)fonts.Length);
        for (var face = 0; face < fonts.Length; face++)
        {
            BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(12 + (face * 4)), (uint)directoryOffsets[face]);
        }

        return output;
    }

    private sealed class SingleFontResolver(string family, byte[] data, string? face = null, int? faceIndex = null) : IReportFontResolver
    {
        public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
            String.Equals(familyName, family, StringComparison.Ordinal) ? new FontResolveInfo(face ?? family) { FaceIndex = faceIndex } : null;

        public ReadOnlyMemory<byte>? GetFont(string faceName) => data;
    }

    private sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
    }
}
