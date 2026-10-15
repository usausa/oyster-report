// Minimal PDF writer. Structure ported from PDFKit (MIT, Copyright (c) 2014 Devon Govett):
//   lib/document.js + lib/object.js (serialization/xref), lib/font/embedded.js (Type0/CIDFontType2/ToUnicode),
//   lib/mixins/vector.js (path ops), lib/image/png.js + lib/image/jpeg.js (image XObjects).
namespace OysterReport.Internal.Pdf;

using System.Globalization;
using System.IO.Compression;
using System.Text;

internal sealed class PdfDocumentWriter
{
    private readonly Stream output;
    private readonly bool compress;
    private readonly List<long> offsets = [0];
    private readonly List<int> pageIds = [];
    private readonly List<PdfEmbeddedFont> fonts = [];
    private readonly Dictionary<byte, int> alphaStates = [];
    private readonly int pagesId;

    public PdfDocumentWriter(Stream output, bool compress)
    {
        this.output = output;
        this.compress = compress;
        Write("%PDF-1.7\n%\u00E2\u00E3\u00CF\u00D3\n");
        pagesId = Reserve();
    }

    public string? Title { get; set; }

    public int Reserve()
    {
        offsets.Add(-1);
        return offsets.Count - 1;
    }

    public void WriteObject(int id, string body)
    {
        offsets[id] = output.Position;
        Write(string.Create(CultureInfo.InvariantCulture, $"{id} 0 obj\n{body}\nendobj\n"));
    }

    public void WriteStream(int id, string dict, byte[] data, bool alreadyEncoded = false)
    {
        var filter = string.Empty;
        if (compress && !alreadyEncoded)
        {
            data = Deflate(data);
            filter = "/Filter /FlateDecode ";
        }

        offsets[id] = output.Position;
        Write(string.Create(CultureInfo.InvariantCulture, $"{id} 0 obj\n<< {filter}{dict} /Length {data.Length} >>\nstream\n"));
        output.Write(data);
        Write("\nendstream\nendobj\n");
    }

    public static byte[] Deflate(ReadOnlySpan<byte> data)
    {
        using var ms = new MemoryStream();
        using (var z = new ZLibStream(ms, CompressionLevel.Optimal, leaveOpen: true))
        {
            z.Write(data);
        }

        return ms.ToArray();
    }

    public PdfEmbeddedFont AddFont(TrueTypeFont font)
    {
        var f = new PdfEmbeddedFont(font, string.Create(CultureInfo.InvariantCulture, $"F{fonts.Count + 1}"), this);
        fonts.Add(f);
        return f;
    }

    internal int GetAlphaState(byte alpha)
    {
        if (!alphaStates.TryGetValue(alpha, out var id))
        {
            id = Reserve();
            var a = (alpha / 255d).ToString("0.###", CultureInfo.InvariantCulture);
            WriteObject(id, $"<< /Type /ExtGState /ca {a} /CA {a} >>");
            alphaStates[alpha] = id;
        }

        return id;
    }

    public PdfPageCanvas BeginPage(double width, double height) => new(this, width, height);

    internal void EndPage(PdfPageCanvas page)
    {
        var contentId = Reserve();
        WriteStream(contentId, string.Empty, Encoding.ASCII.GetBytes(page.Content.ToString()));
        var fontRes = string.Join(' ', page.UsedFonts.Select(static f => string.Create(CultureInfo.InvariantCulture, $"/{f.Name} {f.Id} 0 R")));
        var xobjRes = string.Join(' ', page.UsedImages.Select(static i => string.Create(CultureInfo.InvariantCulture, $"/Im{i} {i} 0 R")));
        var gsRes = string.Join(' ', page.UsedStates.Select(static i => string.Create(CultureInfo.InvariantCulture, $"/GS{i} {i} 0 R")));
        var pageId = Reserve();
        WriteObject(pageId, string.Create(CultureInfo.InvariantCulture,
            $"<< /Type /Page /Parent {pagesId} 0 R /MediaBox [0 0 {page.Width:0.####} {page.Height:0.####}] /Contents {contentId} 0 R " +
            $"/Resources << /ProcSet [/PDF /Text /ImageB /ImageC /ImageI] /Font << {fontRes} >> /XObject << {xobjRes} >> /ExtGState << {gsRes} >> >> >>"));
        pageIds.Add(pageId);
    }

    public void Close()
    {
        // Fonts that were resolved but never drawn are not embedded
        foreach (var f in fonts.Where(static x => x.IsUsed))
        {
            f.Embed(this);
        }

        WriteObject(pagesId, string.Create(CultureInfo.InvariantCulture,
            $"<< /Type /Pages /Count {pageIds.Count} /Kids [{string.Join(' ', pageIds.Select(static p => string.Create(CultureInfo.InvariantCulture, $"{p} 0 R")))}] >>"));
        var catalogId = Reserve();
        WriteObject(catalogId, string.Create(CultureInfo.InvariantCulture, $"<< /Type /Catalog /Pages {pagesId} 0 R >>"));
        var infoId = Reserve();
        var title = Title is null ? string.Empty : "/Title " + TextString(Title);
        WriteObject(infoId, $"<< /Producer {TextString("OysterReport")} {title} >>");

        var xref = output.Position;
        var sb = new StringBuilder();
        sb.Append(CultureInfo.InvariantCulture, $"xref\n0 {offsets.Count}\n0000000000 65535 f \n");
        for (var i = 1; i < offsets.Count; i++)
        {
            sb.Append(CultureInfo.InvariantCulture, $"{offsets[i]:D10} 00000 n \n");
        }

        sb.Append(CultureInfo.InvariantCulture, $"trailer\n<< /Size {offsets.Count} /Root {catalogId} 0 R /Info {infoId} 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        Write(sb.ToString());
    }

    internal static string TextString(string s) =>
        "<FEFF" + Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(s)) + ">";

    private void Write(string s) => output.Write(Encoding.Latin1.GetBytes(s));
}

internal sealed class PdfPageCanvas
{
    private readonly PdfDocumentWriter document;

    public PdfPageCanvas(PdfDocumentWriter document, double width, double height)
    {
        this.document = document;
        Width = width;
        Height = height;
        // Callers use a top-left origin (XGraphics PageDirection.Downwards). Like PDFsharp, coordinates are
        // converted to PDF user space here instead of flipping the CTM, so text is never drawn mirrored.
    }

    public double Width { get; }

    public double Height { get; }

    public StringBuilder Content { get; } = new();

    public HashSet<PdfEmbeddedFont> UsedFonts { get; } = [];

    public HashSet<int> UsedImages { get; } = [];

    public HashSet<int> UsedStates { get; } = [];

    private void Op(FormattableString s) => Content.Append(FormattableString.Invariant(s)).Append('\n');

    public void Save() => Op($"q");

    public void Restore() => Op($"Q");

    private double Y(double y) => Height - y;

    public void ClipRect(double x, double y, double w, double h) => Op($"{x:0.####} {Y(y + h):0.####} {w:0.####} {h:0.####} re W n");

    private void SetFill(PdfColor c)
    {
        Op($"{c.R / 255d:0.####} {c.G / 255d:0.####} {c.B / 255d:0.####} rg");
        SetAlpha(c.A);
    }

    private void SetStroke(PdfColor c)
    {
        Op($"{c.R / 255d:0.####} {c.G / 255d:0.####} {c.B / 255d:0.####} RG");
        SetAlpha(c.A);
    }

    private void SetAlpha(byte alpha)
    {
        if (alpha == 255)
        {
            return;
        }

        var id = document.GetAlphaState(alpha);
        UsedStates.Add(id);
        Op($"/GS{id} gs");
    }

    public void FillRects(PdfColor color, IEnumerable<ReportRect> rects)
    {
        Save();
        SetFill(color);
        foreach (var r in rects)
        {
            Op($"{r.X:0.####} {Y(r.Y + r.Height):0.####} {r.Width:0.####} {r.Height:0.####} re");
        }

        Op($"f");
        Restore();
    }

    public void FillRect(PdfColor color, double x, double y, double w, double h)
    {
        Save();
        SetFill(color);
        Op($"{x:0.####} {Y(y + h):0.####} {w:0.####} {h:0.####} re f");
        Restore();
    }

    // PDFsharp XPen defaults: flat cap (0), miter join (0)
    public void StrokeLine(PdfColor color, double width, double[] dash, double x1, double y1, double x2, double y2)
    {
        Save();
        SetStroke(color);
        var pattern = string.Join(' ', dash.Select(static d => d.ToString("0.####", CultureInfo.InvariantCulture)));
        Op($"{width:0.####} w 0 J 0 j [{pattern}]0 d {x1:0.####} {Y(y1):0.####} m {x2:0.####} {Y(y2):0.####} l S");
        Restore();
    }

    public void DrawImage(int imageId, double x, double y, double w, double h)
    {
        UsedImages.Add(imageId);
        Op($"q {w:0.####} 0 0 {h:0.####} {x:0.####} {Y(y + h):0.####} cm /Im{imageId} Do Q");
    }

    // skew: 0 for upright, sin(20deg) for PDFsharp-compatible italic simulation
    public void ShowText(PdfEmbeddedFont font, double size, PdfColor color, double x, double baseline, string text, double skew)
    {
        UsedFonts.Add(font);
        var hex = font.Encode(text);
        _ = font.Id;
        Save();
        SetFill(color);
        Op($"BT /{font.Name} {size:0.####} Tf 1 0 {skew:0.####} 1 {x:0.####} {Y(baseline):0.####} Tm <{hex}> Tj ET");
        Restore();
    }
}

internal readonly record struct PdfColor(byte A, byte R, byte G, byte B)
{
    public static PdfColor Black { get; } = new(255, 0, 0, 0);

    public static PdfColor White { get; } = new(255, 255, 255, 255);
}

// Port of PDFKit lib/font/embedded.js (no shaping: plain cmap lookup, same as PDFsharp).
// TrueType: glyphs are subset and renumbered (CIDFontType2, CIDToGIDMap Identity).
// CFF (OpenType/CFF, CID-keyed): CFF table embedded as CIDFontType0C, codes are the font's CIDs.
internal sealed class PdfEmbeddedFont(TrueTypeFont font, string name, PdfDocumentWriter owner)
{
    private readonly TrueTypeSubset? subset = font.IsCff ? null : new TrueTypeSubset(font);
    private readonly SortedDictionary<int, (int Width, int CodePoint)> codes = [];
    private int? id;

    public string Name { get; } = name;

    // Object id is reserved on first use
    public int Id => id ??= owner.Reserve();

    public bool IsUsed => id.HasValue;

    public TrueTypeFont Font { get; } = font;

    public string Encode(string text)
    {
        var sb = new StringBuilder(text.Length * 4);
        foreach (var rune in text.EnumerateRunes())
        {
            var gid = Font.GetGlyphId(rune.Value);
            var code = Font.Cff is { } cff ? cff.ToCid(gid) : subset!.IncludeGlyph(gid);
            sb.Append(code.ToString("X4", CultureInfo.InvariantCulture));
            // [PDFsharp compat] integer division (truncation) like GlyphIndexToPdfWidth; viewers advance glyphs by /W
            // ToUnicode from the *input* code point (no Kangxi radical mapping issue)
            codes.TryAdd(code, (Font.GetAdvanceWidth(gid) * 1000 / Font.UnitsPerEm, rune.Value));
        }

        return sb.ToString();
    }

    public void Embed(PdfDocumentWriter doc)
    {
        var scale = 1000d / Font.UnitsPerEm;
        var descent = -Font.CellDescent * scale;
        var fontFileId = doc.Reserve();
        string fontFileKey;
        if (Font.Cff is { } cff)
        {
            if (!cff.IsCidKeyed)
            {
                throw new NotSupportedException("Name-keyed CFF fonts are not supported yet.");
            }

            doc.WriteStream(fontFileId, "/Subtype /CIDFontType0C", cff.Data);
            fontFileKey = "FontFile3";
        }
        else
        {
            var fontData = subset!.Encode();
            doc.WriteStream(fontFileId, string.Create(CultureInfo.InvariantCulture, $"/Length1 {fontData.Length}"), fontData);
            fontFileKey = "FontFile2";
        }

        var flags = 1 << 2; // symbolic
        if (Font.IsFixedPitch)
        {
            flags |= 1 << 0;
        }

        var familyClass = Font.FamilyClass >> 8;
        if (familyClass is >= 1 and <= 7)
        {
            flags |= 1 << 1;
        }

        if (familyClass == 10)
        {
            flags |= 1 << 3;
        }

        var tag = string.Concat(Enumerable.Range(0, 6).Select(i => (char)('A' + (((Id * 7) + (i * 3)) % 26))));
        var baseFont = Font.IsCff ? Font.PostScriptName.Replace(' ', '_') : $"{tag}+{Font.PostScriptName.Replace(' ', '_')}";

        var descriptorId = doc.Reserve();
        doc.WriteObject(descriptorId, string.Create(CultureInfo.InvariantCulture,
            $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {flags} " +
            $"/FontBBox [{Font.XMin * scale:0.###} {Font.YMin * scale:0.###} {Font.XMax * scale:0.###} {Font.YMax * scale:0.###}] " +
            $"/ItalicAngle {Font.ItalicAngle:0.###} /Ascent {Font.CellAscent * scale:0.###} /Descent {descent:0.###} " +
            $"/CapHeight {Font.CapHeight * scale:0.###} /XHeight {Font.XHeight * scale:0.###} /StemV 0 /{fontFileKey} {fontFileId} 0 R >>"));

        // [PDFsharp compat] .notdef (code 0) is not listed in /W, so viewers advance missing glyphs by /DW (default 1000)
        var w = string.Join(' ', codes.Where(static c => c.Key != 0).Select(static c => string.Create(CultureInfo.InvariantCulture, $"{c.Key} [{c.Value.Width}]")));
        var subtype = Font.IsCff ? "/CIDFontType0" : "/CIDFontType2 /CIDToGIDMap /Identity";
        var cidFontId = doc.Reserve();
        doc.WriteObject(cidFontId, string.Create(CultureInfo.InvariantCulture,
            $"<< /Type /Font /Subtype {subtype} /BaseFont /{baseFont} " +
            $"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {descriptorId} 0 R /W [{w}] >>"));

        var toUnicodeId = doc.Reserve();
        doc.WriteStream(toUnicodeId, string.Empty, Encoding.ASCII.GetBytes(BuildToUnicode()));

        doc.WriteObject(Id, string.Create(CultureInfo.InvariantCulture,
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H /DescendantFonts [{cidFontId} 0 R] /ToUnicode {toUnicodeId} 0 R >>"));
    }

    private string BuildToUnicode()
    {
        static string Hex(int v) => v.ToString("x4", CultureInfo.InvariantCulture);
        static string Utf16(int cp)
        {
            if (cp <= 0xFFFF)
            {
                return Hex(cp);
            }

            var v = cp - 0x10000;
            return Hex(((v >> 10) & 0x3FF) | 0xD800) + Hex(0xDC00 | (v & 0x3FF));
        }

        var entries = codes.Where(static c => c.Key != 0).Select(static c => $"<{Hex(c.Key)}> <{Utf16(c.Value.CodePoint)}>").ToList();
        var sb = new StringBuilder();
        sb.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
        sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
        sb.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000><ffff>\nendcodespacerange\n");
        for (var start = 0; start < entries.Count; start += 100)
        {
            var chunk = entries.Skip(start).Take(100).ToList();
            sb.Append(CultureInfo.InvariantCulture, $"{chunk.Count} beginbfchar\n{string.Join('\n', chunk)}\nendbfchar\n");
        }

        sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend");
        return sb.ToString();
    }
}
