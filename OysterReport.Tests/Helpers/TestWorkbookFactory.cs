namespace OysterReport.Tests.Helpers;

internal static class TestWorkbookFactory
{
    public static MemoryStream CreateWorkbook(Action<IXLWorkbook> configure)
    {
        using var workbook = new XLWorkbook();
        configure(workbook);

        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }
}
