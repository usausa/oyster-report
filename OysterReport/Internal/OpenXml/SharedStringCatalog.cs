namespace OysterReport.Internal.OpenXml;

using System.Text;

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
            list.Add(ExtractText(item));
        }

        return list.ToArray();
    }

    private static string ExtractText(SharedStringItem item)
    {
        if (item.Text is not null)
        {
            return item.Text.Text;
        }

        var sb = new StringBuilder();
        foreach (var child in item.ChildElements)
        {
            if (child is Run run)
            {
                if (run.Text?.Text is { } runText)
                {
                    sb.Append(runText);
                }
            }
            else if (child is Text t && t.Text is { } textValue)
            {
                sb.Append(textValue);
            }
        }

        return sb.ToString();
    }
}
