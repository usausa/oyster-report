namespace Example;

using OysterReport;

// 埋め込み IPAex ゴシックフォントを使うリゾルバー。
// ゴシック系の日本語フォントを "IPAexGothic" に統一し、TTF バイト列を返す。
// FontData を返すことで PDFSharp はシステム検索を行わず、渡されたバイト列を直接使用する。
// 明朝・HG 系フォントは null を返して既定の解決（Windows インストール済みフォント）に委ねる。
internal sealed class IpaExGothicFontResolver : IReportFontResolver
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

    public IpaExGothicFontResolver(string fontFilePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fontFilePath);
        fontData = File.ReadAllBytes(fontFilePath);
    }

    public FontInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        if (!GothicFontNames.Contains(familyName))
        {
            return null;
        }

        return new FontInfo(BuildFaceName(bold, italic))
        {
            MustSimulateBold = bold,
            MustSimulateItalic = italic
        };
    }

    public ReadOnlyMemory<byte>? GetFont(string faceName) =>
        String.Equals(ExtractBaseFaceName(faceName), EmbeddedFontName, StringComparison.Ordinal)
            ? fontData
            : null;

    private static string BuildFaceName(bool bold, bool italic)
    {
        var faceName = EmbeddedFontName;
        if (bold)
        {
            faceName += "#b";
        }

        if (italic)
        {
            faceName += "#i";
        }

        return faceName;
    }

    private static string ExtractBaseFaceName(string faceName) =>
        faceName
            .Replace("#b", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("#i", string.Empty, StringComparison.OrdinalIgnoreCase);
}
