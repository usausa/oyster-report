// Loads shared strings from sharedStrings.xml into an index-addressable array.
// Uses OpenXmlReader (SAX) to avoid building a full DOM for potentially large dictionaries.

namespace OysterReport.Prototype;

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
        // A shared-string item is either plain <t> text or one or more <r> runs (rich text).
        if (item.Text is not null)
        {
            return item.Text.Text ?? string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var child in item.ChildElements)
        {
            if (child is Run run)
            {
                if (run.Text is not null)
                {
                    sb.Append(run.Text.Text);
                }
            }
            else if (child is Text t)
            {
                sb.Append(t.Text);
            }
        }

        return sb.ToString();
    }
}
