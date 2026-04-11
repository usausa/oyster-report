namespace Example;

using OysterReport;

internal sealed class EmbeddedFontResolver : IReportFontResolver
{
    private const string EmbeddedFontName = "IPAexGothic";

    private static readonly HashSet<string> GothicFontNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "ＭＳ Ｐゴシック",
            "MS Pゴシック",
            "ＭＳ ゴシック",
            "メイリオ",
            "Meiryo",
            "游ゴシック",
            "Yu Gothic",
            "游ゴシック Medium",
            "Yu Gothic Medium"
        };

    private readonly ReadOnlyMemory<byte> fontData;

    public EmbeddedFontResolver()
    {
        fontData = File.ReadAllBytes("ipaexg.ttf");
    }

    public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        if (!GothicFontNames.Contains(familyName))
        {
            return null;
        }

        return new FontResolveInfo(EmbeddedFontName)
        {
            MustSimulateBold = bold,
            MustSimulateItalic = italic
        };
    }

    public ReadOnlyMemory<byte>? GetFont(string faceName) =>
        String.Equals(faceName, EmbeddedFontName, StringComparison.Ordinal)
            ? fontData
            : null;
}
