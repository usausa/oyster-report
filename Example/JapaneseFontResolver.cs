namespace Example;

using OysterReport;

// Windows にインストールされた TTC フォントは PDFsharp 6.x では
// collectionNumber が未実装 (face 0 のみ対応) のため、
// TTC の face 0 に対応するフォント名 (レジストリキーの先頭名) へマッピングする。
//
// 各 TTC の face 0:
//   msgothic.ttc  → MS Gothic     (レジストリキー先頭: "MS Gothic & ...")
//   msmincho.ttc  → MS Mincho     (レジストリキー先頭: "MS Mincho & ...")
//   meiryo.ttc    → Meiryo        (レジストリキー先頭: "Meiryo & ...")
//   HGRME.TTC     → HG明朝E      (レジストリキー先頭: "HG明朝E & ...")
internal sealed class JapaneseFontResolver : IReportFontResolver
{
    private static readonly Dictionary<string, string> FontMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // MS Gothic 系 (msgothic.ttc: face 0 = MS Gothic)
            ["ＭＳ Ｐゴシック"] = "MS Gothic",
            ["MS PGothic"] = "MS Gothic",
            ["MS Pゴシック"] = "MS Gothic",
            ["ＭＳ ゴシック"] = "MS Gothic",
            ["MS UI Gothic"] = "MS Gothic",

            // MS Mincho 系 (msmincho.ttc: face 0 = MS Mincho)
            ["ＭＳ Ｐ明朝"] = "MS Mincho",
            ["MS PMincho"] = "MS Mincho",
            ["MS P明朝"] = "MS Mincho",
            ["ＭＳ 明朝"] = "MS Mincho",

            // Meiryo 系 (meiryo.ttc: face 0 = Meiryo)
            ["Meiryo UI"] = "Meiryo",

            // HG Mincho E 系 (HGRME.TTC: face 0 = HGMinchoE; レジストリキー: "HG明朝E & ...")
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
