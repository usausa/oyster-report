namespace OysterReport;

using ClosedXML.Excel;

public sealed class TemplateRowRange
{
    private readonly IXLWorksheet worksheet;

    public int StartRow { get; }

    public int EndRow { get; }

    public int RowCount => EndRow - StartRow + 1;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateRowRange(IXLWorksheet ws, int startRow, int endRow)
    {
        worksheet = ws;
        StartRow = startRow;
        EndRow = endRow;
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public TemplateRowRange InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    public TemplateRowRange InsertCopyAfter(TemplateRowRange afterRange)
    {
        var newStartRow = afterRange.EndRow + 1;
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        worksheet.Row(newStartRow).InsertRowsAbove(RowCount);

        // this (コピー元) の行番号を再計算: 挿入位置が this より上なら RowCount 分シフト
        var sourceStartRow = (newStartRow <= StartRow) ? StartRow + RowCount : StartRow;

        for (var offset = 0; offset < RowCount; offset++)
        {
            var srcRowNum = sourceStartRow + offset;
            var dstRowNum = newStartRow + offset;
            worksheet.Row(dstRowNum).Height = worksheet.Row(srcRowNum).Height;

            for (var col = 1; col <= lastColumn; col++)
            {
                var srcCell = worksheet.Cell(srcRowNum, col);
                var dstCell = worksheet.Cell(dstRowNum, col);
                dstCell.Value = srcCell.Value;
                dstCell.Style = srcCell.Style;
            }
        }

        return new TemplateRowRange(worksheet, newStartRow, newStartRow + RowCount - 1);
    }

    public void Delete()
    {
        worksheet.Rows(StartRow, EndRow).Delete();
    }

    public int ReplacePlaceholder(string markerName, string value)
    {
        var placeholder = "{{" + markerName + "}}";
        var count = 0;
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (var row = StartRow; row <= EndRow; row++)
        {
            for (var col = 1; col <= lastColumn; col++)
            {
                var cell = worksheet.Cell(row, col);
                var text = cell.GetString();
                if (text.Contains(placeholder, StringComparison.Ordinal))
                {
                    cell.Value = text.Replace(placeholder, value, StringComparison.Ordinal);
                    count++;
                }
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
