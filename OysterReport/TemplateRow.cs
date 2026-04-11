namespace OysterReport;

using ClosedXML.Excel;

public sealed class TemplateRow
{
    private readonly IXLWorksheet worksheet;

    public int RowNumber { get; }

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateRow(IXLWorksheet ws, int rowNumber)
    {
        worksheet = ws;
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

        worksheet.Row(newRowNumber).InsertRowsAbove(1);
        var insertedRow = worksheet.Row(newRowNumber);

        // this (コピー元) の行番号を再計算: 挿入位置が this より上なら +1 シフト
        var sourceRowNum = (newRowNumber <= RowNumber) ? RowNumber + 1 : RowNumber;

        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (var col = 1; col <= lastColumn; col++)
        {
            var srcCell = worksheet.Cell(sourceRowNum, col);
            var dstCell = worksheet.Cell(newRowNumber, col);
            dstCell.Value = srcCell.Value;
            dstCell.Style = srcCell.Style;
        }

        insertedRow.Height = worksheet.Row(sourceRowNum).Height;

        return new TemplateRow(worksheet, newRowNumber);
    }

    public void Delete()
    {
        worksheet.Row(RowNumber).Delete();
    }

    public int ReplacePlaceholder(string markerName, string value)
    {
        var placeholder = "{{" + markerName + "}}";
        var count = 0;
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (var col = 1; col <= lastColumn; col++)
        {
            var cell = worksheet.Cell(RowNumber, col);
            var text = cell.GetString();
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                cell.Value = text.Replace(placeholder, value, StringComparison.Ordinal);
                count++;
            }
        }

        return count;
    }

    // この行内のプレースホルダを辞書で一括置換する。
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
