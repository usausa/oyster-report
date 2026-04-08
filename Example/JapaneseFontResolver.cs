namespace Example;

using OysterReport;

// WindowsInstalledFontResolver は Windows フォントレジストリから TTC の
// フェイスインデックスを正しく解決し、各フェイスの TTF バイト列を抽出する。
// このマップでは Excel セルのフォント名 (全角名・混在名) を
// Windows レジストリで検索可能な ASCII 名へ変換する。
//
// レジストリの TTC 対応:
//   msgothic.ttc  → face 0: MS Gothic (等幅)
//                   face 1: MS UI Gothic (プロポーショナル・UI 用)
//                   face 2: MS PGothic  (プロポーショナル)
//   msmincho.ttc  → face 0: MS Mincho  (等幅)
//                   face 1: MS PMincho  (プロポーショナル)
//   meiryo.ttc    → face 0: Meiryo
//                   face 1: Meiryo Italic
//                   face 2: Meiryo UI
//                   face 3: Meiryo UI Italic
internal sealed class JapaneseFontResolver : IReportFontResolver
{
    private static readonly Dictionary<string, string> FontMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // MS Gothic 系 (msgothic.ttc) — 全角名・混在名 → ASCII レジストリキー名
            ["ＭＳ Ｐゴシック"] = "MS PGothic",   // face 2 (プロポーショナル)
            ["MS Pゴシック"] = "MS PGothic",        // face 2 (プロポーショナル)
            ["ＭＳ ゴシック"] = "MS Gothic",       // face 0 (等幅)

            // MS Mincho 系 (msmincho.ttc)
            ["ＭＳ Ｐ明朝"] = "MS PMincho",        // face 1 (プロポーショナル)
            ["MS P明朝"] = "MS PMincho",            // face 1 (プロポーショナル)
            ["ＭＳ 明朝"] = "MS Mincho",           // face 0 (等幅)

            // HG Mincho E 系
            ["HGP明朝E"] = "HG明朝E",
            ["HGPMinchoE"] = "HG明朝E",
            ["HGS明朝E"] = "HG明朝E",
            ["HGSMinchoE"] = "HG明朝E"
        };

    public ReportFontResolveResult Resolve(ReportFontRequest request)
    {
        if (FontMap.TryGetValue(request.FontName, out var resolved))
        {
            return new ReportFontResolveResult
            {
                IsResolved = true,
                ResolvedFontName = resolved
            };
        }

        return new ReportFontResolveResult { IsResolved = false };
    }
}
