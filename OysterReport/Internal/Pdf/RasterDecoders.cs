namespace OysterReport.Internal.Pdf;

using System.Buffers;
using System.IO.Compression;
using System.Text;

// 8-bit samples; Colors = 1 (gray) or 3 (RGB); Alpha null when opaque
internal sealed record RasterImage(int Width, int Height, int Colors, byte[] Color, byte[]? Alpha);

internal static class PngDecoder
{
    public static RasterImage Decode(ReadOnlySpan<byte> png)
    {
        int width = 0, height = 0, depth = 8, colorType = 0, interlace = 0;
        byte[] palette = [];
        byte[]? trns = null;
        using var idat = new MemoryStream();
        var pos = 8;
        while (pos + 8 <= png.Length)
        {
            var len = (int)TrueTypeFont.U32(png, pos);
            var type = Encoding.ASCII.GetString(png.Slice(pos + 4, 4));
            var body = png.Slice(pos + 8, len);
            switch (type)
            {
                case "IHDR":
                    width = (int)TrueTypeFont.U32(body, 0);
                    height = (int)TrueTypeFont.U32(body, 4);
                    depth = body[8];
                    colorType = body[9];
                    interlace = body[12];
                    break;
                case "PLTE": palette = [.. body]; break;
                case "tRNS": trns = [.. body]; break;
                case "IDAT": idat.Write(body); break;
            }

            pos += 12 + len;
        }

        var channels = colorType switch { 0 => 1, 2 => 3, 3 => 1, 4 => 2, 6 => 4, _ => throw new InvalidDataException($"Unsupported PNG color type. colorType=[{colorType}]") };
        var bitsPerPixel = channels * depth;
        var filterBpp = Math.Max(1, bitsPerPixel / 8);

        using var z = new ZLibStream(new MemoryStream(idat.ToArray()), CompressionMode.Decompress);
        using var raw = new MemoryStream();
        z.CopyTo(raw);
        var src = raw.GetBuffer();
        var srcPos = 0;

        // samples[y, x, c] as 16-bit values (or palette index)
        var samples = new ushort[width * height * channels];
        if (interlace == 0)
        {
            Unfilter(src, ref srcPos, width, height, channels, depth, filterBpp, (x, y) => (x, y), samples, width);
        }
        else
        {
            ReadOnlySpan<(int X0, int Y0, int Dx, int Dy)> passes = [(0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2)];
            foreach (var (x0, y0, dx, dy) in passes)
            {
                var pw = (width - x0 + dx - 1) / dx;
                var ph = (height - y0 + dy - 1) / dy;
                if (pw <= 0 || ph <= 0)
                {
                    continue;
                }

                Unfilter(src, ref srcPos, pw, ph, channels, depth, filterBpp, (x, y) => (x0 + (x * dx), y0 + (y * dy)), samples, width);
            }
        }

        var max = (1 << depth) - 1;
        byte To8(ushort v) => depth == 16 ? (byte)(v >> 8) : depth == 8 ? (byte)v : (byte)(v * 255 / max);
        var colors = colorType is 0 or 4 ? 1 : 3;
        var color = new byte[width * height * colors];
        var alpha = (colorType is 4 or 6) || (trns is not null) ? new byte[width * height] : null;
        for (var i = 0; i < width * height; i++)
        {
            var s = i * channels;
            switch (colorType)
            {
                case 0:
                    color[i] = To8(samples[s]);
                    alpha?[i] = samples[s] == TrueTypeFont.U16(trns, 0) ? (byte)0 : (byte)255;

                    break;
                case 2:
                    for (var c = 0; c < 3; c++)
                    {
                        color[(i * 3) + c] = To8(samples[s + c]);
                    }

                    if (alpha is not null)
                    {
                        var key = samples[s] == TrueTypeFont.U16(trns, 0) && samples[s + 1] == TrueTypeFont.U16(trns, 2) && samples[s + 2] == TrueTypeFont.U16(trns, 4);
                        alpha[i] = key ? (byte)0 : (byte)255;
                    }

                    break;
                case 3:
                    var index = samples[s];
                    color[i * 3] = palette[index * 3];
                    color[(i * 3) + 1] = palette[(index * 3) + 1];
                    color[(i * 3) + 2] = palette[(index * 3) + 2];
                    alpha?[i] = index < trns!.Length ? trns[index] : (byte)255;

                    break;
                case 4:
                    color[i] = To8(samples[s]);
                    alpha![i] = To8(samples[s + 1]);
                    break;
                default:
                    for (var c = 0; c < 3; c++)
                    {
                        color[(i * 3) + c] = To8(samples[s + c]);
                    }

                    alpha![i] = To8(samples[s + 3]);
                    break;
            }
        }

        return new RasterImage(width, height, colors, color, alpha);
    }

    private static void Unfilter(byte[] src, ref int pos, int w, int h, int channels, int depth, int bpp, Func<int, int, (int X, int Y)> map, ushort[] samples, int fullWidth)
    {
        var stride = ((w * channels * depth) + 7) / 8;
        var prev = new byte[stride];
        var row = new byte[stride];
        for (var y = 0; y < h; y++)
        {
            var filter = src[pos++];
            for (var x = 0; x < stride; x++)
            {
                var a = x >= bpp ? row[x - bpp] : 0;
                var b = prev[x];
                var c = x >= bpp ? prev[x - bpp] : 0;
                var v = src[pos++];
                row[x] = filter switch
                {
                    0 => v,
                    1 => (byte)(v + a),
                    2 => (byte)(v + b),
                    3 => (byte)(v + ((a + b) >> 1)),
                    4 => (byte)(v + Paeth(a, b, c)),
                    _ => throw new InvalidDataException($"Unknown PNG filter. filter=[{filter}]")
                };
            }

            for (var x = 0; x < w; x++)
            {
                var (tx, ty) = map(x, y);
                for (var ch = 0; ch < channels; ch++)
                {
                    var sampleIndex = (x * channels) + ch;
                    var value = depth switch
                    {
                        16 => (ushort)((row[sampleIndex * 2] << 8) | row[(sampleIndex * 2) + 1]),
                        8 => row[sampleIndex],
                        _ => (ushort)((row[(sampleIndex * depth) / 8] >> (8 - depth - ((sampleIndex * depth) % 8))) & ((1 << depth) - 1))
                    };
                    samples[(((ty * fullWidth) + tx) * channels) + ch] = value;
                }
            }

            (prev, row) = (row, prev);
        }
    }

    private static int Paeth(int a, int b, int c)
    {
        var p = a + b - c;
        int pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
        return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
    }
}

internal static class BmpDecoder
{
    public static RasterImage Decode(ReadOnlySpan<byte> bmp)
    {
        var dataOffset = BitConverter.ToInt32(bmp[10..]);
        var headerSize = BitConverter.ToInt32(bmp[14..]);
        var width = BitConverter.ToInt32(bmp[18..]);
        var rawHeight = BitConverter.ToInt32(bmp[22..]);
        var bpp = BitConverter.ToUInt16(bmp[28..]);
        var compression = BitConverter.ToInt32(bmp[30..]);
        if (compression is not (0 or 3))
        {
            throw new NotSupportedException($"Compressed BMP is not supported. compression=[{compression}]");
        }

        var height = Math.Abs(rawHeight);
        var topDown = rawHeight < 0;
        var stride = (((width * bpp) + 31) / 32) * 4;
        var paletteOffset = 14 + headerSize;
        var hasAlpha = bpp == 32 && headerSize >= 56 && BitConverter.ToUInt32(bmp[(14 + 52)..]) != 0;
        var color = new byte[width * height * 3];
        var alpha = hasAlpha ? new byte[width * height] : null;
        for (var y = 0; y < height; y++)
        {
            var row = bmp.Slice(dataOffset + ((topDown ? y : height - 1 - y) * stride), stride);
            for (var x = 0; x < width; x++)
            {
                int r, g, b, a = 255;
                switch (bpp)
                {
                    case 24:
                        (b, g, r) = (row[x * 3], row[(x * 3) + 1], row[(x * 3) + 2]);
                        break;
                    case 32:
                        (b, g, r, a) = (row[x * 4], row[(x * 4) + 1], row[(x * 4) + 2], row[(x * 4) + 3]);
                        break;
                    case 8 or 4 or 1:
                        var index = bpp == 8 ? row[x] : (row[(x * bpp) / 8] >> (8 - bpp - ((x * bpp) % 8))) & ((1 << bpp) - 1);
                        var p = paletteOffset + (index * 4);
                        (b, g, r) = (bmp[p], bmp[p + 1], bmp[p + 2]);
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported BMP bit depth. bpp=[{bpp}]");
                }

                var i = (y * width) + x;
                color[i * 3] = (byte)r;
                color[(i * 3) + 1] = (byte)g;
                color[(i * 3) + 2] = (byte)b;
                alpha?[i] = (byte)a;
            }
        }

        return new RasterImage(width, height, 3, color, alpha);
    }
}

internal static class GifDecoder
{
    private const int MaxCodes = 4096;

    // First frame only (Excel pictures are static)
    public static RasterImage Decode(ReadOnlySpan<byte> gif)
    {
        using var raster = DecodePooled(gif);
        return raster.ToRasterImage();
    }

    // Planes are rented from ArrayPool; the caller disposes the result to return them
    public static PooledRaster DecodePooled(ReadOnlySpan<byte> gif)
    {
        var width = BitConverter.ToUInt16(gif[6..]);
        var height = BitConverter.ToUInt16(gif[8..]);
        var flags = gif[10];
        var pos = 13;
        var globalTable = ReadOnlySpan<byte>.Empty;
        if ((flags & 0x80) != 0)
        {
            var size = 3 << ((flags & 7) + 1);
            globalTable = gif.Slice(pos, size);
            pos += size;
        }

        var transparent = -1;
        while (pos < gif.Length)
        {
            var block = gif[pos++];
            if (block == 0x21)
            {
                var label = gif[pos++];
                if (label == 0xF9 && (gif[pos + 1] & 1) != 0)
                {
                    transparent = gif[pos + 4];
                }

                while (gif[pos] != 0)
                {
                    pos += gif[pos] + 1;
                }

                pos++;
                continue;
            }

            if (block != 0x2C)
            {
                break;
            }

            var left = BitConverter.ToUInt16(gif[pos..]);
            var top = BitConverter.ToUInt16(gif[(pos + 2)..]);
            var iw = BitConverter.ToUInt16(gif[(pos + 4)..]);
            var ih = BitConverter.ToUInt16(gif[(pos + 6)..]);
            var imageFlags = gif[pos + 8];
            pos += 9;
            var table = globalTable;
            if ((imageFlags & 0x80) != 0)
            {
                var size = 3 << ((imageFlags & 7) + 1);
                table = gif.Slice(pos, size);
                pos += size;
            }

            int minCodeSize = gif[pos++];
            var dataLength = MeasureSubBlocks(gif, pos);
            var pixelCount = iw * ih;
            var data = ArrayPool<byte>.Shared.Rent(dataLength);
            var indices = ArrayPool<byte>.Shared.Rent(pixelCount);
            try
            {
                var dataSpan = data.AsSpan(0, dataLength);
                CopySubBlocks(gif, pos, dataSpan);
                var indexSpan = indices.AsSpan(0, pixelCount);
                var written = Lzw(dataSpan, minCodeSize, indexSpan);

                indexSpan[written..].Clear();

                PooledRaster? raster = null;
                try
                {
                    raster = new PooledRaster(width, height, 3, transparent >= 0);
                    var coversCanvas = (left == 0) && (top == 0) && (iw == width) && (ih == height);
                    if (!coversCanvas)
                    {
                        raster.Color.Clear();
                        raster.Alpha.Clear();
                    }

                    var interlaced = (imageFlags & 0x40) != 0;
                    WritePixels(indexSpan, table, transparent, raster, left, top, iw, ih, interlaced);
                    var result = raster;
                    raster = null;
                    return result;
                }
                finally
                {
                    raster?.Dispose();
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(indices);
                ArrayPool<byte>.Shared.Return(data);
            }
        }

        throw new InvalidDataException("GIF has no image.");
    }

    private static int MeasureSubBlocks(ReadOnlySpan<byte> gif, int pos)
    {
        var length = 0;
        while (gif[pos] != 0)
        {
            length += gif.Slice(pos + 1, gif[pos]).Length;
            pos += gif[pos] + 1;
        }

        return length;
    }

    private static void CopySubBlocks(ReadOnlySpan<byte> gif, int pos, Span<byte> destination)
    {
        var offset = 0;
        while (gif[pos] != 0)
        {
            var block = gif.Slice(pos + 1, gif[pos]);
            block.CopyTo(destination[offset..]);
            offset += block.Length;
            pos += gif[pos] + 1;
        }
    }

    private static void WritePixels(ReadOnlySpan<byte> indices, ReadOnlySpan<byte> table, int transparent, PooledRaster raster, int left, int top, int iw, int ih, bool interlaced)
    {
        if (!interlaced)
        {
            for (var r = 0; r < ih; r++)
            {
                WriteRow(indices.Slice(r * iw, iw), table, transparent, raster, ((top + r) * raster.Width) + left);
            }

            return;
        }

        ReadOnlySpan<(int Start, int Step)> passes = [(0, 8), (4, 8), (2, 4), (1, 2)];
        var source = 0;
        foreach (var (start, step) in passes)
        {
            for (var y = start; y < ih; y += step)
            {
                WriteRow(indices.Slice(source * iw, iw), table, transparent, raster, ((top + y) * raster.Width) + left);
                source++;
            }
        }
    }

    private static void WriteRow(ReadOnlySpan<byte> row, ReadOnlySpan<byte> table, int transparent, PooledRaster raster, int start)
    {
        var color = raster.Color;
        var alpha = raster.Alpha;
        var hasAlpha = raster.HasAlpha;
        for (var x = 0; x < row.Length; x++)
        {
            var index = row[x];
            var i = start + x;
            color[i * 3] = table[index * 3];
            color[(i * 3) + 1] = table[(index * 3) + 1];
            color[(i * 3) + 2] = table[(index * 3) + 2];
            if (hasAlpha)
            {
                alpha[i] = index == transparent ? (byte)0 : (byte)255;
            }
        }
    }

    private static int Lzw(ReadOnlySpan<byte> data, int minCodeSize, Span<byte> output)
    {
        var clear = 1 << minCodeSize;
        var end = clear + 1;

        var prefixArray = ArrayPool<ushort>.Shared.Rent(MaxCodes);
        var bytesArray = ArrayPool<byte>.Shared.Rent(MaxCodes * 2);
        try
        {
            var prefix = prefixArray.AsSpan(0, MaxCodes);
            var suffix = bytesArray.AsSpan(0, MaxCodes);
            var stack = bytesArray.AsSpan(MaxCodes, MaxCodes);
            prefix.Clear();
            suffix.Clear();
            for (var i = 0; i < clear; i++)
            {
                suffix[i] = (byte)i;
            }

            int codeSize = minCodeSize + 1, next = end + 1, prev = -1, outPos = 0;

            uint accumulator = 0;
            int bits = 0, dataPos = 0;
            var pixelCount = output.Length;
            while (outPos < pixelCount && ((long)(data.Length - dataPos) * 8) + bits >= codeSize)
            {
                while (bits < codeSize)
                {
                    accumulator |= (uint)data[dataPos++] << bits;
                    bits += 8;
                }

                var code = (int)(accumulator & ((1u << codeSize) - 1));
                accumulator >>= codeSize;
                bits -= codeSize;

                if (code == clear)
                {
                    codeSize = minCodeSize + 1;
                    next = end + 1;
                    prev = -1;
                    continue;
                }

                if (code == end)
                {
                    break;
                }

                if (prev < 0)
                {
                    output[outPos++] = suffix[code];
                    prev = code;
                    continue;
                }

                var len = 0;
                var c = code < next ? code : prev;
                while (c >= clear)
                {
                    stack[len++] = suffix[c];
                    c = prefix[c];
                }

                stack[len++] = suffix[c];
                var first = suffix[c];
                for (var i = len - 1; i >= 0 && outPos < pixelCount; i--)
                {
                    output[outPos++] = stack[i];
                }

                if ((code >= next) && (outPos < pixelCount))
                {
                    output[outPos++] = first;
                }

                if (next < MaxCodes)
                {
                    prefix[next] = (ushort)prev;
                    suffix[next] = first;
                    next++;
                    if (next == (1 << codeSize) && codeSize < 12)
                    {
                        codeSize++;
                    }
                }

                prev = code;
            }

            return outPos;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(bytesArray);
            ArrayPool<ushort>.Shared.Return(prefixArray);
        }
    }
}

internal sealed class PooledRaster : IDisposable
{
    private byte[]? color;
    private byte[]? alpha;

    public PooledRaster(int width, int height, int colors, bool hasAlpha)
    {
        Width = width;
        Height = height;
        Colors = colors;
        color = ArrayPool<byte>.Shared.Rent(width * height * colors);
        alpha = hasAlpha ? ArrayPool<byte>.Shared.Rent(width * height) : null;
    }

    public int Width { get; }

    public int Height { get; }

    public int Colors { get; }

    public bool HasAlpha => alpha is not null;

    public Span<byte> Color => color.AsSpan(0, Width * Height * Colors);

    public Span<byte> Alpha => alpha is null ? [] : alpha.AsSpan(0, Width * Height);

    public RasterImage ToRasterImage() => new(Width, Height, Colors, [.. Color], HasAlpha ? [.. Alpha] : null);

    public void Dispose()
    {
        if (color is not null)
        {
            ArrayPool<byte>.Shared.Return(color);
            color = null;
        }

        if (alpha is not null)
        {
            ArrayPool<byte>.Shared.Return(alpha);
            alpha = null;
        }
    }
}
