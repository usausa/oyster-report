namespace OysterReport;

/// <summary>
/// PDF 描画時の調整値をまとめたオプション。
/// 既定値は Excel に近い見た目になるように調整されている。
/// </summary>
public sealed record ReportRenderingOptions
{
    /// <summary>セル内テキストの左右余白 (pt)。</summary>
    public double HorizontalCellTextPaddingPoints { get; init; } = 2d;

    /// <summary>セル内テキストの既定フォントサイズ (pt)。</summary>
    public double DefaultCellFontSizePoints { get; init; } = 11d;

    /// <summary>ヘッダー・フッターの既定フォントサイズ (pt)。</summary>
    public double HeaderFooterFontSizePoints { get; init; } = 9d;

    /// <summary>ヘッダー・フッター描画時のフォールバック候補一覧。</summary>
    public IReadOnlyList<string> HeaderFooterFallbackFontNames { get; init; } =
    [
        "Arial",
        "Segoe UI",
        "Helvetica",
        "Liberation Sans",
        "DejaVu Sans"
    ];

    /// <summary>Thick 罫線の描画幅 (pt)。</summary>
    public double ThickBorderWidthPoints { get; init; } = 2.25d;

    /// <summary>Medium 罫線の描画幅 (pt)。</summary>
    public double MediumBorderWidthPoints { get; init; } = 1.5d;

    /// <summary>通常罫線の描画幅 (pt)。</summary>
    public double NormalBorderWidthPoints { get; init; } = 0.75d;

    /// <summary>Hair 罫線の描画幅 (pt)。</summary>
    public double HairBorderWidthPoints { get; init; } = 0.25d;

    /// <summary>列幅ポイント変換時の補正係数。</summary>
    public double ColumnWidthAdjustment { get; init; } = 1d;

    /// <summary>未知フォントに使用するフォールバック最大桁幅 (96 DPI 参照ピクセル)。</summary>
    public double FallbackMaxDigitWidth { get; init; } = 7d;
}
