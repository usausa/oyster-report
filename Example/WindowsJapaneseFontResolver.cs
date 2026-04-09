namespace Example;

using OysterReport;

// Windows のインストール済みフォントを使うリゾルバー。
// Excel の全角フォント名や混在名を、Windows レジストリで検索可能な ASCII 名に変換する。
// FontData は返さないため PDFSharp は WindowsInstalledFontResolver を経由してフォントを取得する。
//
// レジストリ上の TTC フェイス構成:
//   msgothic.ttc  → face 0: MS Gothic / face 1: MS UI Gothic / face 2: MS PGothic
//   msmincho.ttc  → face 0: MS Mincho  / face 1: MS PMincho
//   meiryo.ttc    → face 0: Meiryo     / face 1: Meiryo Italic / face 2: Meiryo UI
internal sealed class WindowsJapaneseFontResolver : IReportFontResolver
{
    private static readonly Dictionary<string, string> FontMap =
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

    public ReportFontResolveResult? ResolveFont(ReportFontRequest request)
    {
        if (FontMap.TryGetValue(request.FontName, out var installedName))
        {
            // FontData = null → PDFSharp が WindowsInstalledFontResolver 経由でレジストリ検索する
            return new ReportFontResolveResult { FontName = installedName };
        }

        return null;
    }

    // ResolveMaxDigitWidth は実装しない → ライブラリ内蔵の推定テーブルを使用する
}
