namespace OysterReport;

using ClosedXML.Excel;

using OysterReport.Internal;

public sealed class TemplateSheet
{
    private readonly ReportWorkbook workbook;

    internal ReportSheet UnderlyingSheet { get; }

    internal ReportMeasurementProfile WorkbookMeasurementProfile => workbook.MeasurementProfile;

    public string Name => UnderlyingSheet.Name;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateSheet(ReportWorkbook workbook, ReportSheet sheet)
    {
        this.workbook = workbook;
        UnderlyingSheet = sheet;
    }

    //--------------------------------------------------------------------------------
    // Row
    //--------------------------------------------------------------------------------

    public TemplateRow GetRow(int row) => new(UnderlyingSheet, row);

    public TemplateRowRange GetRows(int startRow, int endRow) => new(UnderlyingSheet, startRow, endRow);

    public TemplateRow FindRow(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                return new TemplateRow(UnderlyingSheet, cell.Row);
            }
        }

        throw new InvalidOperationException($"Marker not found in sheet. maker=[{marker}]");
    }

    public TemplateRowRange FindRows(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        var markerRow = -1;
        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                markerRow = cell.Row;
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

        return new TemplateRowRange(UnderlyingSheet, startRow, endRow);
    }

    private bool RowContainsAnyPlaceholder(int rowNum)
    {
        if (rowNum < 1)
        {
            return false;
        }

        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.Row != rowNum)
            {
                continue;
            }

            var text = cell.DisplayText;
            if (text.Contains("{{", StringComparison.Ordinal) && text.Contains("}}", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
    }

    public void DeleteRows(int startRow, int endRow)
    {
        UnderlyingSheet.DeleteRows(startRow, endRow);
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public int ReplacePlaceholder(string marker, string value)
    {
        var placeholder = "{{" + marker + "}}";
        var count = 0;

        foreach (var cell in UnderlyingSheet.Cells)
        {
            var text = cell.DisplayText;
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                var replaced = text.Replace(placeholder, value, StringComparison.Ordinal);
                SetCellText(cell, replaced);
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

    //--------------------------------------------------------------------------------
    // Cell text access (for tests and diagnostics)
    //--------------------------------------------------------------------------------

    public string GetCellText(int row, int column)
    {
        var cell = UnderlyingSheet.FindCell(row, column);
        return cell is null ? string.Empty : cell.DisplayText;
    }

    internal static void SetCellText(ReportCell cell, string value)
    {
        cell.Value = new ReportCellValue
        {
            Kind = XLDataType.Text,
            RawValue = value
        };
        cell.DisplayText = value;
    }
}
