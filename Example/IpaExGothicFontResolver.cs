namespace Example;

using OysterReport;

// 埋め込み IPAex ゴシックフォントを使うリゾルバー。
// ゴシック系の日本語フォントを "IPAexGothic" に統一し、TTF バイト列を返す。
// FontData を返すことで PDFSharp はシステム検索を行わず、渡されたバイト列を直接使用する。
// 明朝・HG 系フォントは null を返して既定の解決（Windows インストール済みフォント）に委ねる。
internal sealed class IpaExGothicFontResolver : IReportFontResolver
{
    private const string EmbeddedFontName = "IPAexGothic";

    // IPAex ゴシックの 96 DPI 参照ピクセルでの最大桁幅 (10pt 基準の実測値)
    private const double MaxDigitWidthAt10Pt = 8d;

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

    public ReportFontResolveResult? ResolveFont(ReportFontRequest request)
    {
        if (!GothicFontNames.Contains(request.FontName))
        {
            return null;
        }

        // FontData を返すことで PDFSharp はバイト列を直接使用し、システム検索しない
        return new ReportFontResolveResult
        {
            FontName = EmbeddedFontName,
            FontData = fontData
        };
    }

    public double? ResolveMaxDigitWidth(string fontName, double fontSizePoints)
    {
        if (!GothicFontNames.Contains(fontName))
        {
            return null;
        }

        return MaxDigitWidthAt10Pt * (fontSizePoints / 10d);
    }
}
