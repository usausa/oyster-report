namespace OysterReport.Tests.Helpers;

using DocumentFormat.OpenXml.Packaging;

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

    // Applies raw OpenXML edits (e.g. inline strings, the 1904 date flag) that ClosedXML cannot express
    public static MemoryStream CreateWorkbook(Action<IXLWorkbook> configure, Action<SpreadsheetDocument> postProcess)
    {
        var stream = CreateWorkbook(configure);
        using (var document = SpreadsheetDocument.Open(stream, isEditable: true))
        {
            postProcess(document);
        }

        stream.Position = 0;
        return stream;
    }

    // Rewrites an ARGB value in the stylesheet (e.g. into the 00RRGGBB form openpyxl writes for 6-digit colors)
    public static void ReplaceRgbColor(SpreadsheetDocument document, string source, string replacement)
    {
        foreach (var color in document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Descendants<ColorType>())
        {
            if (color.Rgb?.Value == source)
            {
                color.Rgb = replacement;
            }
        }
    }
}
