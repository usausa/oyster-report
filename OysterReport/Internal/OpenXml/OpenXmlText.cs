namespace OysterReport.Internal.OpenXml;

using System.Text;

using DocumentFormat.OpenXml.Spreadsheet;

internal static class OpenXmlText
{
    // Extracts plain text from a shared string item or inline string, joining rich text runs
    public static string Extract(RstType item)
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
            else if (child is Text { Text: { } textValue })
            {
                sb.Append(textValue);
            }
        }

        return sb.ToString();
    }
}
