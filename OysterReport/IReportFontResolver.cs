namespace OysterReport;

public interface IReportFontResolver
{
    /// <summary>
    /// フォントを解決する。
    /// <see langword="null" /> を返した場合は既定のフォント解決にフォールバックする。
    /// </summary>
    ReportFontResolveResult? ResolveFont(ReportFontRequest request);
}
