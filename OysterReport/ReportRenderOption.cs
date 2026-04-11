namespace OysterReport;

using System.Diagnostics.CodeAnalysis;

using ClosedXML.Excel;

[ExcludeFromCodeCoverage]
public sealed record ReportRenderOption
{
    // 用紙サイズからページ幅・高さ (pt) を解決する関数。
    public Func<XLPaperSize, (double Width, double Height)> PageSizeResolver { get; set; } = ResolveDefaultPageSize;

    // セル内テキストの左右余白 (pt)。
    public double HorizontalCellTextPaddingPoints { get; set; } = 2d;

    // セル内テキストの既定フォントサイズ (pt)。
    public double DefaultCellFontSizePoints { get; set; } = 11d;

    // ヘッダー・フッターの既定フォントサイズ (pt)。
    public double HeaderFooterFontSizePoints { get; set; } = 9d;

    // ヘッダー・フッター描画時のフォールバック候補一覧。
    public IReadOnlyList<string> HeaderFooterFallbackFontNames { get; set; } =
    [
        "Arial",
        "Segoe UI",
        "Helvetica",
        "Liberation Sans",
        "DejaVu Sans"
    ];

    // Thick 罫線の描画幅 (pt)。
    public double ThickBorderWidthPoints { get; set; } = 2.25d;

    // Medium 罫線の描画幅 (pt)。
    public double MediumBorderWidthPoints { get; set; } = 1.5d;

    // 通常罫線の描画幅 (pt)。
    public double NormalBorderWidthPoints { get; set; } = 0.75d;

    // Hair 罫線の描画幅 (pt)。
    public double HairBorderWidthPoints { get; set; } = 0.25d;

    // 下線の描画幅 (pt)。フォントメトリクスの値より小さい場合のフォールバック最小値。
    // Underline drawing width (pt). Minimum fallback when font metrics suggest a smaller value.
    public double UnderlineWidthPoints { get; set; } = 0.5d;

    // 打ち消し線の描画幅 (pt)。フォントメトリクスの値より小さい場合のフォールバック最小値。
    // Strikeout drawing width (pt). Minimum fallback when font metrics suggest a smaller value.
    public double StrikeoutWidthPoints { get; set; } = 0.5d;

    // 列幅ポイント変換時の補正係数。
    public double ColumnWidthAdjustment { get; set; } = 1d;

    // 未知フォントに使用するフォールバック最大桁幅 (96 DPI 参照ピクセル)。
    public double FallbackMaxDigitWidth { get; set; } = 7d;

    private static (double Width, double Height) ResolveDefaultPageSize(XLPaperSize paperSize) =>
        paperSize switch
        {
            XLPaperSize.LetterPaper => (612d, 792d),
            XLPaperSize.LegalPaper => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
}
