namespace OysterReport.Tests;

using System.Buffers.Binary;
using System.Text;

using OysterReport.Internal;

#pragma warning disable CA1416
public sealed class FontInfrastructureTests
{
    //--------------------------------------------------------------------------------
    // TTC extraction
    //--------------------------------------------------------------------------------

    // Builds a minimal single-face TTC: header(12) + one face offset(4) + sfnt header(12) +
    // one table directory entry(16) + 4 bytes of table data.
    private static byte[] BuildMinimalTtc()
    {
        var ttc = new byte[48];
        Encoding.ASCII.GetBytes("ttcf", 0, 4, ttc, 0);
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(4, 4), 0x00010000);
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(8, 4), 1);            // numFonts
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(12, 4), 16);          // face offset

        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(16, 4), 0x00010000);  // sfnt version
        BinaryPrimitives.WriteUInt16BigEndian(ttc.AsSpan(20, 2), 1);           // numTables

        Encoding.ASCII.GetBytes("cmap", 0, 4, ttc, 28);                        // table entry
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(32, 4), 0x12345678);  // checksum
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(36, 4), 44);          // src offset
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(40, 4), 4);           // length
        Encoding.ASCII.GetBytes("ABCD", 0, 4, ttc, 44);                        // table data
        return ttc;
    }

    [Fact]
    public void ExtractTtc_ValidData_RebuildsTtf()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var ttf = WindowsFontResolver.ExtractTtfFaceFromTtc(BuildMinimalTtc(), 0, "test.ttc");

        // header(12) + directory(16) + data(4) padded to 4 = 32
        Assert.Equal(32, ttf.Length);
        Assert.Equal("cmap", Encoding.ASCII.GetString(ttf, 12, 4));
        Assert.Equal(28u, BinaryPrimitives.ReadUInt32BigEndian(ttf.AsSpan(20, 4)));
        Assert.Equal("ABCD", Encoding.ASCII.GetString(ttf, 28, 4));
    }

    [Fact]
    public void ExtractTtc_TruncatedHeader_ReportsCorruption()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<InvalidDataException>(static () => WindowsFontResolver.ExtractTtfFaceFromTtc(new byte[8], 0, "test.ttc"));
        Assert.Contains("header is truncated", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTtc_FaceIndexOutOfRange_ReportsCorruption()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<InvalidDataException>(static () => WindowsFontResolver.ExtractTtfFaceFromTtc(BuildMinimalTtc(), 1, "test.ttc"));
        Assert.Contains("face index is out of range", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTtc_FaceOffsetOutOfBounds_ReportsCorruption()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var ttc = BuildMinimalTtc();
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(12, 4), 0xFFFFFFF0);

        var ex = Assert.Throws<InvalidDataException>(() => WindowsFontResolver.ExtractTtfFaceFromTtc(ttc, 0, "test.ttc"));
        Assert.Contains("face offset is out of bounds", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTtc_TableLengthBeyondFile_ReportsCorruption()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var ttc = BuildMinimalTtc();
        // Would have become a negative int length (and an obscure AsSpan failure) before validation
        BinaryPrimitives.WriteUInt32BigEndian(ttc.AsSpan(40, 4), 0x80000004);

        var ex = Assert.Throws<InvalidDataException>(() => WindowsFontResolver.ExtractTtfFaceFromTtc(ttc, 0, "test.ttc"));
        Assert.Contains("table is out of bounds", ex.Message, StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Embedded font registration
    //--------------------------------------------------------------------------------

    [Fact]
    public void RegisterEmbeddedFont_SameBuffer_KeepsExistingCopy()
    {
        var name = "ReRegSame-" + Guid.NewGuid().ToString("N");
        var data = new byte[] { 1, 2, 3, 4 };
        var adapter = new ReportFontResolverAdapter();

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, data);
        var first = adapter.GetFont(name);

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, data);
        Assert.Same(first, adapter.GetFont(name));
    }

    [Fact]
    public void RegisterEmbeddedFont_SameContentDifferentBuffer_KeepsExistingCopy()
    {
        var name = "ReRegClone-" + Guid.NewGuid().ToString("N");
        var data = new byte[] { 1, 2, 3, 4 };
        var adapter = new ReportFontResolverAdapter();

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, data);
        var first = adapter.GetFont(name);

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, (byte[])data.Clone());
        Assert.Same(first, adapter.GetFont(name));
    }

    [Fact]
    public void RegisterEmbeddedFont_DifferentContent_ReplacesData()
    {
        var name = "ReRegNew-" + Guid.NewGuid().ToString("N");
        var adapter = new ReportFontResolverAdapter();

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, new byte[] { 1, 2, 3, 4 });
        ReportFontResolverAdapter.RegisterEmbeddedFont(name, new byte[] { 9, 9 });

        Assert.Equal([9, 9], adapter.GetFont(name));
    }

    [Fact]
    public void RegisterEmbeddedFont_IsDefensiveCopy()
    {
        var name = "ReRegCopy-" + Guid.NewGuid().ToString("N");
        var data = new byte[] { 1, 2, 3, 4 };
        var adapter = new ReportFontResolverAdapter();

        ReportFontResolverAdapter.RegisterEmbeddedFont(name, data);
        data[0] = 99;

        Assert.Equal(1, adapter.GetFont(name)[0]);
    }
}
#pragma warning restore CA1416
