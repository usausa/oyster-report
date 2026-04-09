namespace Example;

using OysterReport;

internal sealed class JapaneseFontResolver : IReportFontResolver
{
    private const string EmbeddedFontName = "IPAexGothic";
    private const double EmbeddedFontMaxDigitWidthAt10Pt = 8d;

    private static readonly Dictionary<string, string> InstalledFontMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["ＭＳ Ｐゴシック"] = "MS PGothic",
            ["MS Pゴシック"] = "MS PGothic",
            ["ＭＳ ゴシック"] = "MS Gothic",
            ["ＭＳ Ｐ明朝"] = "MS PMincho",
            ["MS P明朝"] = "MS PMincho",
            ["ＭＳ 明朝"] = "MS Mincho",
            ["HGP明朝E"] = "HG明朝E",
            ["HGPMinchoE"] = "HG明朝E",
            ["HGS明朝E"] = "HG明朝E",
            ["HGSMinchoE"] = "HG明朝E"
        };

    private readonly ReadOnlyMemory<byte>? embeddedFontData;

    private JapaneseFontResolver(ReadOnlyMemory<byte>? embeddedFontData)
    {
        this.embeddedFontData = embeddedFontData;
    }

    public static JapaneseFontResolver CreateInstalledFontResolver() => new(null);

    public static JapaneseFontResolver CreateEmbeddedFontResolver(string fontPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fontPath);
        return new JapaneseFontResolver(File.ReadAllBytes(fontPath));
    }

    public ReportFontResolveResult? ResolveFont(ReportFontRequest request)
    {
        if (!InstalledFontMap.ContainsKey(request.FontName))
        {
            return null;
        }

        if (embeddedFontData is { } fontData)
        {
            return new ReportFontResolveResult
            {
                FontName = EmbeddedFontName,
                FontData = fontData
            };
        }

        return new ReportFontResolveResult
        {
            FontName = InstalledFontMap[request.FontName]
        };
    }

    public double? ResolveMaxDigitWidth(string fontName, double fontSizePoints)
    {
        if (embeddedFontData is null || !InstalledFontMap.ContainsKey(fontName))
        {
            return null;
        }

        return EmbeddedFontMaxDigitWidthAt10Pt * (fontSizePoints / 10d);
    }
}
