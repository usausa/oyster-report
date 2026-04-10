// <copyright file="IpaExGothicFontResolver.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests.Helpers;

using OysterReport;

/// <summary>
/// テスト用の IPAex ゴシックフォントリゾルバー。
/// ゴシック系日本語フォントを ipaexg.ttf で解決する。
/// </summary>
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
        fontData = File.ReadAllBytes(fontFilePath);
    }

    public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        if (!GothicFontNames.Contains(familyName))
        {
            return null;
        }

        return new FontResolveInfo(BuildFaceName(bold, italic))
        {
            MustSimulateBold = bold,
            MustSimulateItalic = italic
        };
    }

    public ReadOnlyMemory<byte>? GetFont(string faceName) =>
        string.Equals(ExtractBaseFaceName(faceName), EmbeddedFontName, StringComparison.Ordinal)
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
