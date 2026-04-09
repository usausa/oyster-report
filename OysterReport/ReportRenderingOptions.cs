namespace OysterReport;

using ClosedXML.Excel;

/// <summary>
/// PDF 描画時の調整値をまとめたオプション。
/// 既定値は Excel に近い見た目になるように調整されている。
/// </summary>
public sealed record ReportRenderingOptions
{
    /// <summary>用紙サイズからページ幅・高さ (pt) を解決する関数。</summary>
    [CLSCompliant(false)]
    public Func<XLPaperSize, (double Width, double Height)> PageSizeResolver { get; set; } = ResolveDefaultPageSize;

    /// <summary>セル内テキストの左右余白 (pt)。</summary>
    public double HorizontalCellTextPaddingPoints { get; set; } = 2d;

    /// <summary>セル内テキストの既定フォントサイズ (pt)。</summary>
    public double DefaultCellFontSizePoints { get; set; } = 11d;

    /// <summary>ヘッダー・フッターの既定フォントサイズ (pt)。</summary>
    public double HeaderFooterFontSizePoints { get; set; } = 9d;

    /// <summary>ヘッダー・フッター描画時のフォールバック候補一覧。</summary>
    public IReadOnlyList<string> HeaderFooterFallbackFontNames { get; set; } =
    [
        "Arial",
        "Segoe UI",
        "Helvetica",
        "Liberation Sans",
        "DejaVu Sans"
    ];

    /// <summary>Thick 罫線の描画幅 (pt)。</summary>
    public double ThickBorderWidthPoints { get; set; } = 2.25d;

    /// <summary>Medium 罫線の描画幅 (pt)。</summary>
    public double MediumBorderWidthPoints { get; set; } = 1.5d;

    /// <summary>通常罫線の描画幅 (pt)。</summary>
    public double NormalBorderWidthPoints { get; set; } = 0.75d;

    /// <summary>Hair 罫線の描画幅 (pt)。</summary>
    public double HairBorderWidthPoints { get; set; } = 0.25d;

    /// <summary>列幅ポイント変換時の補正係数。</summary>
    public double ColumnWidthAdjustment { get; set; } = 1d;

    /// <summary>未知フォントに使用するフォールバック最大桁幅 (96 DPI 参照ピクセル)。</summary>
    public double FallbackMaxDigitWidth { get; set; } = 7d;

    private static (double Width, double Height) ResolveDefaultPageSize(XLPaperSize paperSize) =>
        paperSize switch
        {
            XLPaperSize.LetterPaper => (612d, 792d),
            XLPaperSize.LegalPaper => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
}
