namespace OysterReport.Internal.Pdf;

using System.Collections.Concurrent;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

internal sealed class PdfFontManager(PdfDocumentWriter writer, IReportFontResolver? resolver, FontLocator locator)
{
    private static readonly ConditionalWeakTable<byte[], ConcurrentDictionary<int, TrueTypeFont>> ParsedFaces = [];

    private readonly Dictionary<string, PdfEmbeddedFont?> fontsByKey = [with(StringComparer.OrdinalIgnoreCase)];
    private readonly Dictionary<string, PdfEmbeddedFont> resolverFaces = [with(StringComparer.OrdinalIgnoreCase)];
    private readonly Dictionary<string, PdfEmbeddedFont> resolverFacesByName = [with(StringComparer.OrdinalIgnoreCase)];

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
            face = GetResolverFace(resolved.FaceName, resolved.FaceIndex);
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
        var face = resolverFacesByName.GetValueOrDefault(name) ?? GetSystemFace(name, false, false, out _, out _);
        return face is null ? null : new PdfFont(face, size, false, false);
    }

    private static TrueTypeFont Parse(byte[] bytes, int faceIndex) =>
        ParsedFaces.GetValue(bytes, static _ => new ConcurrentDictionary<int, TrueTypeFont>())
            .GetOrAdd(faceIndex, index => TrueTypeFont.Load(bytes, index));

    private PdfEmbeddedFont? GetResolverFace(string faceName, int? faceIndex)
    {
        var key = faceIndex is null ? faceName : String.Create(CultureInfo.InvariantCulture, $"{faceName}\0{faceIndex}");
        if (resolverFaces.TryGetValue(key, out var cached))
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

        // Collections: explicit index, else the face named FaceName, else the first face
        var index = faceIndex ?? TrueTypeFont.FindFaceIndex(bytes, faceName) ?? 0;
        var face = writer.AddFont(Parse(bytes, index));
        resolverFaces[key] = face;
        resolverFacesByName.TryAdd(faceName, face);
        return face;
    }

    private PdfEmbeddedFont? GetSystemFace(string family, bool bold, bool italic, out bool simulateBold, out bool simulateItalic)
    {
        simulateBold = false;
        simulateItalic = false;
        if (resolverFacesByName.TryGetValue(family, out var registered))
        {
            return registered;
        }

        if (!locator.TryResolve(family, bold, italic, out var entry, out simulateBold, out simulateItalic))
        {
            return null;
        }

        var key = String.Create(CultureInfo.InvariantCulture, $"{entry.Path}|{entry.FaceIndex}");
        if (!fontsByKey.TryGetValue(key, out var face))
        {
            face = writer.AddFont(Parse(SystemFontCatalog.LoadFile(entry), entry.FaceIndex));
            fontsByKey[key] = face;
        }

        return face;
    }
}

internal sealed class PdfFont(PdfEmbeddedFont face, double size, bool simulateBold, bool simulateItalic)
{
    public const double ItalicSkew = 0.34202014332566873;

    public PdfEmbeddedFont Face { get; } = face;

    public double Size { get; } = size;

    public bool SimulateBold { get; } = simulateBold;

    public bool SimulateItalic { get; } = simulateItalic;

    private TrueTypeFont Font => Face.Font;

    public double LineSpace => Font.CellSpace * Size / Font.UnitsPerEm;

    public double CellAscent => LineSpace * Font.CellAscent / Font.CellSpace;

    public double CellDescent => LineSpace * Font.CellDescent / Font.CellSpace;

    public int UnitsPerEm => Font.UnitsPerEm;

    public int Ascent => Font.CellAscent;

    public int UnderlinePosition => Font.UnderlinePosition;

    public int UnderlineThickness => Font.UnderlineThickness;

    public int StrikeoutPosition => Font.StrikeoutPosition;

    public int StrikeoutThickness => Font.StrikeoutSize;

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
