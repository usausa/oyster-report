namespace Example;

using OysterReport;

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

    public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
        FontMap.TryGetValue(familyName, out var resolvedName) ? new FontResolveInfo(resolvedName) : null;
}
