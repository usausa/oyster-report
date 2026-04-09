namespace OysterReport;

using ClosedXML.Excel;

public sealed class TemplateSheet
{
#pragma warning disable IDE0032
    private readonly IXLWorksheet worksheet;
#pragma warning restore IDE0032

#pragma warning disable IDE0032
    internal IXLWorksheet UnderlyingWorksheet => worksheet;
#pragma warning restore IDE0032

    public string Name => worksheet.Name;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateSheet(IXLWorksheet worksheet)
    {
        this.worksheet = worksheet;
    }

    //--------------------------------------------------------------------------------
    // Row
    //--------------------------------------------------------------------------------

    public TemplateRow GetRow(int row) => new(worksheet, row);

    public TemplateRowRange GetRows(int startRow, int endRow) => new(worksheet, startRow, endRow);

    public TemplateRow FindRow(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.GetString().Contains(placeholder, StringComparison.Ordinal))
            {
                return new TemplateRow(worksheet, cell.Address.RowNumber);
            }
        }

        throw new InvalidOperationException($"Marker not found in sheet. maker=[{marker}]");
    }

    public TemplateRowRange FindRows(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        var markerRow = -1;
        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.GetString().Contains(placeholder, StringComparison.Ordinal))
            {
                markerRow = cell.Address.RowNumber;
                break;
            }
        }

        if (markerRow < 0)
        {
            throw new InvalidOperationException($"Marker not found in sheet. maker=[{marker}]");
        }

        var startRow = markerRow;
        var endRow = markerRow;

        while (RowContainsAnyPlaceholder(endRow + 1))
        {
            endRow++;
        }

        while ((startRow > 1) && RowContainsAnyPlaceholder(startRow - 1))
        {
            startRow--;
        }

        return new TemplateRowRange(worksheet, startRow, endRow);
    }

    private bool RowContainsAnyPlaceholder(int rowNum)
    {
        if (rowNum < 1)
        {
            return false;
        }

        var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (var col = 1; col <= lastCol; col++)
        {
            var text = worksheet.Cell(rowNum, col).GetString();
            if (text.Contains("{{", StringComparison.Ordinal) && text.Contains("}}", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
    }

    public void DeleteRows(int startRow, int endRow)
    {
        worksheet.Rows(startRow, endRow).Delete();
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public int ReplacePlaceholder(string marker, string value)
    {
        var placeholder = "{{" + marker + "}}";
        var count = 0;

        foreach (var cell in worksheet.CellsUsed())
        {
            var text = cell.GetString();
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                cell.Value = text.Replace(placeholder, value, StringComparison.Ordinal);
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
