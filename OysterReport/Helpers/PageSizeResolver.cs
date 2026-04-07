namespace OysterReport.Helpers;

using ClosedXML.Excel;

internal static class PageSizeResolver
{
    public static (double Width, double Height) GetPageSize(XLPaperSize paperSize)
    {
        return paperSize switch
        {
            XLPaperSize.LetterPaper => (612d, 792d),
            XLPaperSize.LegalPaper => (612d, 1008d),
            _ => (595.28d, 841.89d)
        };
    }
}
