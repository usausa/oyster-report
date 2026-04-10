namespace OysterReport;

using ClosedXML.Excel;

// シート上の 1 行を表す軽量ハンドル。コピー挿入・プレースホルダ置換・削除を提供する。
public sealed class TemplateRow
{
    private readonly IXLWorksheet worksheet;

    internal TemplateRow(IXLWorksheet ws, int rowNumber)
    {
        worksheet = ws;
        RowNumber = rowNumber;
    }

    // この行の行番号 (1-based)。
    public int RowNumber { get; }

    // この行のコピーを直下に挿入し、挿入された新しい行を返す。
    // フロー B（行番号を進めながら処理する方式）で使用する。
    public TemplateRow InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    // この行の内容をコピーし、afterRow の直下に挿入する。挿入された新しい行を返す。
    // コピー元は this、挿入位置は afterRow の直下。フロー A（テンプレートのコピーを追加していく方式）で使用する。
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

    // この行内のプレースホルダを置換する。
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

    // この行を削除する。後続行は自動的に上にシフトされる。
    public void Delete()
    {
        worksheet.Row(RowNumber).Delete();
    }
}
