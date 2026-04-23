namespace OysterReport;

using OysterReport.Internal;

public sealed class TemplateRow
{
    private readonly ReportSheet sheet;

    public int RowNumber { get; private set; }

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateRow(ReportSheet sheet, int rowNumber)
    {
        this.sheet = sheet;
        RowNumber = rowNumber;
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public TemplateRow InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    public TemplateRow InsertCopyAfter(TemplateRow afterRow)
    {
        var newRowNumber = afterRow.RowNumber + 1;
        var sourceRowNum = (newRowNumber <= RowNumber) ? RowNumber + 1 : RowNumber;

        sheet.InsertEmptyRowsAt(newRowNumber, 1);

        if (newRowNumber <= RowNumber)
        {
            RowNumber++;
        }

        sheet.CopyRowContent(sourceRowNum, newRowNumber);

        return new TemplateRow(sheet, newRowNumber);
    }

    public void Delete()
    {
        sheet.DeleteRows(RowNumber, RowNumber);
    }

    public int ReplacePlaceholder(string markerName, string value)
    {
        var placeholder = "{{" + markerName + "}}";
        var count = 0;

        foreach (var cell in sheet.Cells)
        {
            if (cell.Row != RowNumber)
            {
                continue;
            }

            var text = cell.DisplayText;
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                var replaced = text.Replace(placeholder, value, StringComparison.Ordinal);
                TemplateSheet.SetCellText(cell, replaced);
                count++;
            }
        }

        return count;
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }
}
