namespace OysterReport.Internal.OpenXml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal static class SharedStringCatalog
{
    public static string[] Load(SharedStringTablePart? part)
    {
        if (part is null)
        {
            return [];
        }

        var list = new List<string>();
        using var reader = OpenXmlReader.Create(part);

        while (reader.Read())
        {
            if (!reader.IsStartElement || (reader.ElementType != typeof(SharedStringItem)))
            {
                continue;
            }

            var item = (SharedStringItem)reader.LoadCurrentElement()!;
            list.Add(OpenXmlText.Extract(item));
        }

        return list.ToArray();
    }
}
