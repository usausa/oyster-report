namespace OysterReport.Benchmark;

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

// Photo-like RGBA PNG (gradients + noise + a transparent band) so that decoding and deflate do realistic work
internal static class BenchmarkImageFactory
{
    private static readonly uint[] CrcTable = CreateCrcTable();

    public static byte[] CreateRgbaPng(int width, int height, uint seed)
    {
        var state = seed | 1u;
        var raw = new byte[height * (1 + (width * 4))];
        var pos = 0;
        for (var y = 0; y < height; y++)
        {
            raw[pos++] = 0;
            for (var x = 0; x < width; x++)
            {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                var noise = (int)(state % 25) - 12;
                raw[pos++] = Clamp(128 + (127 * Math.Sin((x + seed) / 37.0)) + noise);
                raw[pos++] = Clamp(128 + (127 * Math.Cos(y / 23.0)) + noise);
                var pattern = ((x * y) / 97) % 256;
                raw[pos++] = Clamp(pattern + noise);
                raw[pos++] = x < width / 8 ? (byte)0 : (byte)255;
            }
        }

        using var output = new MemoryStream();
        output.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        var header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        header[9] = 6;
        WriteChunk(output, "IHDR", header);
        using (var compressed = new MemoryStream())
        {
            using (var z = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            {
                z.Write(raw);
            }

            WriteChunk(output, "IDAT", compressed.ToArray());
        }

        WriteChunk(output, "IEND", []);
        return output.ToArray();
    }

    private static byte Clamp(double value) => Clamp((int)value);

    private static byte Clamp(int value) => (byte)Math.Clamp(value, 0, 255);

    private static void WriteChunk(Stream output, string type, byte[] data)
    {
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        output.Write(length);
        var typeBytes = Encoding.ASCII.GetBytes(type);
        output.Write(typeBytes);
        output.Write(data);
        var crc = Crc(Crc(0xFFFFFFFFu, typeBytes), data) ^ 0xFFFFFFFFu;
        Span<byte> crcBytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crcBytes, crc);
        output.Write(crcBytes);
    }

    private static uint Crc(uint crc, ReadOnlySpan<byte> data)
    {
        foreach (var b in data)
        {
            crc = CrcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }

        return crc;
    }

    private static uint[] CreateCrcTable()
    {
        var table = new uint[256];
        for (var n = 0u; n < 256; n++)
        {
            var c = n;
            for (var k = 0; k < 8; k++)
            {
                c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            }

            table[n] = c;
        }

        return table;
    }
}
