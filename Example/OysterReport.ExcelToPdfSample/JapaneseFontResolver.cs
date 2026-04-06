// <copyright file="JapaneseFontResolver.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

using OysterReport.Writing.Pdf;

// Windows の TTC フォントのうち、レジストリキーが全角名や略称になっているものを
// Windows フォントレジストリで検索可能な ASCII 名へマッピングする。
//
// TTC フェイス対応 (WindowsInstalledFontResolver が自動判別):
//   msgothic.ttc  → face 0: MS Gothic / face 1: MS UI Gothic / face 2: MS PGothic
//   msmincho.ttc  → face 0: MS Mincho / face 1: MS PMincho
//   meiryo.ttc    → face 0: Meiryo
//   HGRME.TTC     → face 0: HG明朝E
internal sealed class JapaneseFontResolver : IReportFontResolver
{
    private static readonly Dictionary<string, string> FontMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // MS Gothic 系 (msgothic.ttc)
            ["ＭＳ Ｐゴシック"] = "MS PGothic",
            ["MS Pゴシック"] = "MS PGothic",
            ["ＭＳ ゴシック"] = "MS Gothic",

            // MS Mincho 系 (msmincho.ttc)
            ["ＭＳ Ｐ明朝"] = "MS PMincho",
            ["MS P明朝"] = "MS PMincho",
            ["ＭＳ 明朝"] = "MS Mincho",

            // Meiryo 系 (meiryo.ttc)
            ["Meiryo UI"] = "Meiryo",

            // HG Mincho E 系 (HGRME.TTC)
            ["HGP明朝E"] = "HG明朝E",
            ["HGPMinchoE"] = "HG明朝E",
            ["HGS明朝E"] = "HG明朝E",
            ["HGSMinchoE"] = "HG明朝E",
        };

    public ReportFontResolveResult Resolve(ReportFontRequest request)
    {
        if (FontMap.TryGetValue(request.FontName, out var resolved))
        {
            return new ReportFontResolveResult
            {
                IsResolved = true,
                ResolvedFontName = resolved,
            };
        }

        return new ReportFontResolveResult { IsResolved = false };
    }
}
