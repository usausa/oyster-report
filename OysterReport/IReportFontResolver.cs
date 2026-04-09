namespace OysterReport;

public interface IReportFontResolver
{
    /// <summary>
    /// フォントを解決する。
    /// <see langword="null" /> を返した場合は既定のフォント解決にフォールバックする。
    /// </summary>
    ReportFontResolveResult? ResolveFont(ReportFontRequest request);

    /// <summary>
    /// 列幅計算用の最大桁幅を返す。
    /// 単位は 96 DPI 参照ピクセルで、<see langword="null" /> の場合は既定の推定値を使用する。
    /// </summary>
    double? ResolveMaxDigitWidth(string fontName, double fontSizePoints) => null;
}
