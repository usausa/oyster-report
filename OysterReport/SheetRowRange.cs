namespace OysterReport;

using ClosedXML.Excel;

// シート上の連続行範囲を表すハンドル。1 明細が複数行にまたがるテンプレートで使用する。
public sealed class SheetRowRange
{
    private readonly IXLWorksheet worksheet;

    internal SheetRowRange(IXLWorksheet ws, int startRow, int endRow)
    {
        worksheet = ws;
        StartRow = startRow;
        EndRow = endRow;
    }

    // 開始行番号 (1-based)。
    public int StartRow { get; }

    // 終了行番号 (1-based, inclusive)。
    public int EndRow { get; }

    // 行数。
    public int RowCount => EndRow - StartRow + 1;

    // この行範囲のコピーを直下に挿入し、挿入された新しい行範囲を返す。
    // フロー B で使用する。
    public SheetRowRange InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    // この行範囲の内容をコピーし、afterRange の直下に挿入する。挿入された新しい行範囲を返す。
    // コピー元は this、挿入位置は afterRange の直下。フロー A で使用する。
    public SheetRowRange InsertCopyAfter(SheetRowRange afterRange)
    {
        ArgumentNullException.ThrowIfNull(afterRange);
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

        return new SheetRowRange(worksheet, newStartRow, newStartRow + RowCount - 1);
    }

    // この行範囲内のプレースホルダを置換する。
    public int ReplacePlaceholder(string markerName, string value)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        ArgumentNullException.ThrowIfNull(value);
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

    // この行範囲内のプレースホルダを辞書で一括置換する。
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        ArgumentNullException.ThrowIfNull(values);
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }

    // この行範囲を削除する。後続行は自動的に上にシフトされる。
    public void Delete()
    {
        worksheet.Rows(StartRow, EndRow).Delete();
    }
}
