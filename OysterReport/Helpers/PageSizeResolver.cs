namespace OysterReport.Helpers;

using OysterReport;

internal static class PageSizeResolver
{
    public static (double Width, double Height) GetPageSize(ReportPaperSize paperSize)
    {
        return paperSize switch
        {
            ReportPaperSize.Letter => (612d, 792d),
            ReportPaperSize.Legal => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
    }
}
