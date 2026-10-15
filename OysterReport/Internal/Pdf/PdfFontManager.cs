namespace OysterReport.Internal.Pdf;

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// Per-document font resolution (replaces PDFsharp GlobalFontSettings + ReportFontResolverAdapter).
// No process-global state: each document owns its faces and subsets.
internal sealed class PdfFontManager(PdfDocumentWriter writer, IReportFontResolver? resolver)
{
    // Parsed faces are immutable, so they are shared across documents (subset state lives in PdfEmbeddedFont).
    // Keyed by the caller's byte[] so a resolver returning the same buffer never re-copies/re-parses a 6MB font.
    private static readonly ConditionalWeakTable<byte[], TrueTypeFont> ParsedFaces = [];

    private readonly Dictionary<string, PdfEmbeddedFont?> fontsByKey = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, PdfEmbeddedFont> resolverFaces = new(StringComparer.OrdinalIgnoreCase);

    public PdfFont Resolve(ReportFont font, double defaultSize)
    {
        var size = font.Size <= 0 ? defaultSize : font.Size;
        var nameToUse = font.Name;
        var simulateBold = false;
        var simulateItalic = false;
        PdfEmbeddedFont? face = null;

        var resolved = resolver?.ResolveTypeface(font.Name, font.Bold, font.Italic);
        if ((resolved is not null) && !String.IsNullOrWhiteSpace(resolved.FaceName))
        {
            nameToUse = resolved.FaceName;
            face = GetResolverFace(resolved.FaceName);
            if (face is not null)
            {
                simulateBold = resolved.MustSimulateBold;
                simulateItalic = resolved.MustSimulateItalic;
            }
        }

        if (face is null && !String.IsNullOrWhiteSpace(nameToUse))
        {
            face = GetSystemFace(nameToUse, font.Bold, font.Italic, out simulateBold, out simulateItalic);
        }

        // Falls back to the original Excel font name if the resolver-returned name fails
        if (face is null && !String.Equals(nameToUse, font.Name, StringComparison.OrdinalIgnoreCase) && !String.IsNullOrWhiteSpace(font.Name))
        {
            face = GetSystemFace(font.Name, font.Bold, font.Italic, out simulateBold, out simulateItalic);
        }

        if (face is null)
        {
            throw new InvalidOperationException($"No appropriate font found. name=[{font.Name}]");
        }

        return new PdfFont(face, size, simulateBold, simulateItalic);
    }

    public PdfFont? TryResolveByName(string name, double size)
    {
        // Same as the PDFsharp version: only faces already supplied by the resolver, then system fonts (resolver is not queried here)
        var face = resolverFaces.GetValueOrDefault(name) ?? GetSystemFace(name, false, false, out _, out _);
        return face is null ? null : new PdfFont(face, size, false, false);
    }

    private PdfEmbeddedFont? GetResolverFace(string faceName)
    {
        if (resolverFaces.TryGetValue(faceName, out var cached))
        {
            return cached;
        }

        if (resolver?.GetFont(faceName) is not { IsEmpty: false } data)
        {
            return null;
        }

        var bytes = MemoryMarshal.TryGetArray(data, out var segment) && segment.Offset == 0 && segment.Count == segment.Array!.Length
            ? segment.Array
            : data.ToArray();
        var face = writer.AddFont(ParsedFaces.GetValue(bytes, static b => TrueTypeFont.Load(b)));
        resolverFaces[faceName] = face;
        return face;
    }

    private PdfEmbeddedFont? GetSystemFace(string family, bool bold, bool italic, out bool simulateBold, out bool simulateItalic)
    {
        simulateBold = false;
        simulateItalic = false;
        if (resolverFaces.TryGetValue(family, out var registered))
        {
            return registered;
        }

        if (!SystemFontCatalog.Default.TryResolve(family, bold, italic, out var entry, out simulateBold, out simulateItalic))
        {
            return null;
        }

        var key = $"{entry.Path}|{entry.FaceIndex}";
        if (!fontsByKey.TryGetValue(key, out var face))
        {
            var file = SystemFontCatalog.Default.LoadFile(entry);
            face = writer.AddFont(entry.FaceIndex == 0
                ? ParsedFaces.GetValue(file, static b => TrueTypeFont.Load(b))
                : TrueTypeFont.Load(file, entry.FaceIndex));
            fontsByKey[key] = face;
        }

        return face;
    }
}

// A face at a size, with PDFsharp-compatible metrics (XFont equivalents)
internal sealed class PdfFont(PdfEmbeddedFont face, double size, bool simulateBold, bool simulateItalic)
{
    public const double ItalicSkew = 0.34202014332566873; // sin(20deg), PDFsharp Const.ItalicSkewAngleSinus

    public PdfEmbeddedFont Face { get; } = face;

    public double Size { get; } = size;

    public bool SimulateBold { get; } = simulateBold;

    public bool SimulateItalic { get; } = simulateItalic;

    private TrueTypeFont Font => Face.Font;

    // XFont.GetHeight()
    public double LineSpace => Font.CellSpace * Size / Font.UnitsPerEm;

    public double CellAscent => LineSpace * Font.CellAscent / Font.CellSpace;

    public double CellDescent => LineSpace * Font.CellDescent / Font.CellSpace;

    public int UnitsPerEm => Font.UnitsPerEm;

    public int Ascent => Font.CellAscent;

    public int UnderlinePosition => Font.UnderlinePosition;

    public int UnderlineThickness => Font.UnderlineThickness;

    public int StrikeoutPosition => Font.StrikeoutPosition;

    public int StrikeoutThickness => Font.StrikeoutSize;

    // FontHelper.MeasureString (integer sum of advances in font units)
    public double MeasureWidth(string text)
    {
        var width = 0;
        foreach (var rune in text.EnumerateRunes())
        {
            width += Font.GetAdvanceWidth(Font.GetGlyphId(rune.Value));
        }

        return width * Size / Font.UnitsPerEm;
    }
}
