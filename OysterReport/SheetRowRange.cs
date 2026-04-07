namespace OysterReport;

using ClosedXML.Excel;

/// <summary>
/// シート上の連続行範囲を表すハンドル。1 明細が複数行にまたがるテンプレートで使用する。
/// </summary>
public sealed class SheetRowRange
{
    private readonly IXLWorksheet _worksheet;

    internal SheetRowRange(IXLWorksheet worksheet, int startRow, int endRow)
    {
        _worksheet = worksheet;
        StartRow = startRow;
        EndRow = endRow;
    }

    /// <summary>開始行番号 (1-based)。</summary>
    public int StartRow { get; }

    /// <summary>終了行番号 (1-based, inclusive)。</summary>
    public int EndRow { get; }

    /// <summary>行数。</summary>
    public int RowCount => EndRow - StartRow + 1;

    /// <summary>
    /// この行範囲のコピーを直下に挿入し、挿入された新しい行範囲を返す。
    /// フロー B で使用する。
    /// </summary>
    public SheetRowRange InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    /// <summary>
    /// この行範囲の内容をコピーし、afterRange の直下に挿入する。挿入された新しい行範囲を返す。
    /// コピー元は this、挿入位置は afterRange の直下。フロー A で使用する。
    /// </summary>
    public SheetRowRange InsertCopyAfter(SheetRowRange afterRange)
    {
        ArgumentNullException.ThrowIfNull(afterRange);
        var newStartRow = afterRange.EndRow + 1;
        var lastColumn = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        _worksheet.Row(newStartRow).InsertRowsAbove(RowCount);

        // this (コピー元) の行番号を再計算: 挿入位置が this より上なら RowCount 分シフト
        var sourceStartRow = (newStartRow <= StartRow) ? StartRow + RowCount : StartRow;

        for (var offset = 0; offset < RowCount; offset++)
        {
            var srcRowNum = sourceStartRow + offset;
            var dstRowNum = newStartRow + offset;
            _worksheet.Row(dstRowNum).Height = _worksheet.Row(srcRowNum).Height;

            for (var col = 1; col <= lastColumn; col++)
            {
                var srcCell = _worksheet.Cell(srcRowNum, col);
                var dstCell = _worksheet.Cell(dstRowNum, col);
                dstCell.Value = srcCell.Value;
                dstCell.Style = srcCell.Style;
            }
        }

        return new SheetRowRange(_worksheet, newStartRow, newStartRow + RowCount - 1);
    }

    /// <summary>この行範囲内のプレースホルダを置換する。</summary>
    public int ReplacePlaceholder(string markerName, string value)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        ArgumentNullException.ThrowIfNull(value);
        var placeholder = "{{" + markerName + "}}";
        var count = 0;
        var lastColumn = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (var row = StartRow; row <= EndRow; row++)
        {
            for (var col = 1; col <= lastColumn; col++)
            {
                var cell = _worksheet.Cell(row, col);
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

    /// <summary>この行範囲内のプレースホルダを辞書で一括置換する。</summary>
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

    /// <summary>この行範囲を削除する。後続行は自動的に上にシフトされる。</summary>
    public void Delete()
    {
        _worksheet.Rows(StartRow, EndRow).Delete();
    }
}
