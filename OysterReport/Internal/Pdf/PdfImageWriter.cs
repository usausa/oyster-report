// Image XObjects. Ported from PDFKit (MIT) lib/image/jpeg.js, lib/image/png.js and png-js (MIT) decodePixels.
namespace OysterReport.Internal.Pdf;

using System.Globalization;
using System.IO.Compression;
using System.Text;

internal static class PdfImageWriter
{
    public static int Write(PdfDocumentWriter doc, ReadOnlySpan<byte> bytes)
    {
        try
        {
            return WriteCore(doc, bytes);
        }
        catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentOutOfRangeException or OverflowException)
        {
            // Truncated or malformed image data
            throw new InvalidDataException("Corrupted image data.", ex);
        }
    }

    private static int WriteCore(PdfDocumentWriter doc, ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length > 3 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            return WriteJpeg(doc, bytes.ToArray());
        }

        if (bytes.Length > 8 && bytes[0] == 0x89 && bytes[1] == (byte)'P' && bytes[2] == (byte)'N' && bytes[3] == (byte)'G')
        {
            return WritePng(doc, bytes);
        }

        if (bytes.Length > 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M')
        {
            return WriteRaster(doc, BmpDecoder.Decode(bytes));
        }

        if (bytes.Length > 6 && bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F')
        {
            return WriteRaster(doc, GifDecoder.Decode(bytes));
        }

        throw new NotSupportedException("Unsupported image format (JPEG/PNG/BMP/GIF only).");
    }

    private static int WriteRaster(PdfDocumentWriter doc, RasterImage image)
    {
        var id = doc.Reserve();
        var smask = string.Empty;
        if (image.Alpha is not null)
        {
            var maskId = doc.Reserve();
            doc.WriteStream(maskId, string.Create(CultureInfo.InvariantCulture,
                $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /FlateDecode"),
                PdfDocumentWriter.Deflate(image.Alpha), alreadyEncoded: true);
            smask = string.Create(CultureInfo.InvariantCulture, $" /SMask {maskId} 0 R");
        }

        doc.WriteStream(id, string.Create(CultureInfo.InvariantCulture,
            $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /BitsPerComponent 8 /ColorSpace {(image.Colors == 3 ? "/DeviceRGB" : "/DeviceGray")}{smask} /Filter /FlateDecode"),
            PdfDocumentWriter.Deflate(image.Color), alreadyEncoded: true);
        return id;
    }

    private static int WriteJpeg(PdfDocumentWriter doc, byte[] jpeg)
    {
        var pos = 2;
        int width = 0, height = 0, components = 3, bits = 8;
        var adobe = false;
        while (pos + 4 < jpeg.Length)
        {
            while (pos < jpeg.Length && jpeg[pos] != 0xFF)
            {
                pos++;
            }

            int marker = TrueTypeFont.U16(jpeg, pos);
            pos += 2;
            if (marker is (>= 0xFFD0 and <= 0xFFD9) or 0xFF01 or 0xFFFF)
            {
                continue;
            }

            int length = TrueTypeFont.U16(jpeg, pos);
            if (marker == 0xFFEE && length >= 12 && Encoding.ASCII.GetString(jpeg, pos + 2, 5) == "Adobe")
            {
                adobe = true;
            }

            if (marker is >= 0xFFC0 and <= 0xFFCF and not 0xFFC4 and not 0xFFC8 and not 0xFFCC)
            {
                bits = jpeg[pos + 2];
                height = TrueTypeFont.U16(jpeg, pos + 3);
                width = TrueTypeFont.U16(jpeg, pos + 5);
                components = jpeg[pos + 7];
                break;
            }

            pos += length;
        }

        if (width == 0 || height == 0)
        {
            throw new InvalidOperationException("Invalid JPEG (no SOF marker).");
        }

        var cs = components switch { 1 => "/DeviceGray", 4 => "/DeviceCMYK", _ => "/DeviceRGB" };
        // Adobe CMYK JPEGs are stored inverted
        var decode = components == 4 && adobe ? " /Decode [1 0 1 0 1 0 1 0]" : string.Empty;
        var id = doc.Reserve();
        doc.WriteStream(id, string.Create(CultureInfo.InvariantCulture,
            $"/Type /XObject /Subtype /Image /Width {width} /Height {height} /BitsPerComponent {bits} /ColorSpace {cs}{decode} /Filter /DCTDecode"), jpeg, alreadyEncoded: true);
        return id;
    }

    private static int WritePng(PdfDocumentWriter doc, ReadOnlySpan<byte> png)
    {
        int width = 0, height = 0, bits = 8, colorType = 0, interlace = 0;
        byte[] palette = [];
        byte[]? transparency = null;
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
                    bits = body[8];
                    colorType = body[9];
                    interlace = body[12];
                    break;
                case "PLTE":
                    palette = body.ToArray();
                    break;
                case "tRNS":
                    transparency = body.ToArray();
                    break;
                case "IDAT":
                    idat.Write(body);
                    break;
                case "IEND":
                    pos = png.Length;
                    continue;
            }

            pos += 12 + len;
        }

        var hasAlpha = colorType is 4 or 6;
        if (hasAlpha || interlace != 0 || (colorType == 3 && transparency is not null))
        {
            // Needs pixel access (alpha split / Adam7 / palette alpha): managed decode
            return WriteRaster(doc, PngDecoder.Decode(png));
        }

        // No decode: pass the zlib IDAT stream through with the PNG predictor (PDFKit fast path)
        var id = doc.Reserve();
        var colors = colorType == 2 ? 3 : 1;
        var cs = colorType switch
        {
            0 => "/DeviceGray",
            3 => string.Create(CultureInfo.InvariantCulture, $"[/Indexed /DeviceRGB {(palette.Length / 3) - 1} <{Convert.ToHexString(palette)}>]"),
            _ => "/DeviceRGB"
        };
        var mask = string.Empty;
        if (transparency is not null)
        {
            // Color key masking (gray: 1 sample, rgb: 3 samples, 16-bit values)
            var keys = new List<int>();
            for (var i = 0; i + 1 < transparency.Length; i += 2)
            {
                var v = TrueTypeFont.U16(transparency, i);
                keys.Add(v);
                keys.Add(v);
            }

            mask = " /Mask [" + string.Join(' ', keys) + "]";
        }

        doc.WriteStream(id, string.Create(CultureInfo.InvariantCulture,
            $"/Type /XObject /Subtype /Image /Width {width} /Height {height} /BitsPerComponent {bits} /ColorSpace {cs}{mask} " +
            $"/Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors {colors} /BitsPerComponent {bits} /Columns {width} >>"),
            idat.ToArray(), alreadyEncoded: true);
        return id;
    }
}
