namespace OysterReport;

using OysterReport.Internal;

public sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty; // テンプレート名

    public string? SourceFilePath { get; init; } // 読み込み元ファイルパス

    public DateTimeOffset? SourceLastWriteTime { get; init; } // 読み込み元の最終更新日時

    public string? Author { get; init; } // テンプレート作成者
}

public sealed record ReportMeasurementProfile
{
    public double Dpi { get; init; } = 96d; // 計測時に前提とする DPI

    public double MaxDigitWidth { get; init; } = 7d; // 既定フォントでの最大数字幅

    public string DefaultFontName { get; init; } = "Arial"; // 既定フォント名

    public double DefaultFontSize { get; init; } = 11d; // 既定フォントサイズ

    public double ColumnWidthAdjustment { get; init; } = 1d; // 列幅換算の補正係数
}

public sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];
    private readonly List<ReportDiagnostic> diagnostics = [];

    public ReportWorkbook(ReportMetadata? metadata = null, ReportMeasurementProfile? measurementProfile = null)
    {
        Metadata = metadata ?? new ReportMetadata();
        MeasurementProfile = measurementProfile ?? new ReportMeasurementProfile();
    }

    public IReadOnlyList<ReportSheet> Sheets => sheets; // ワークブックに含まれるシート一覧

    public ReportMetadata Metadata { get; } // 帳票全体に関するメタデータ

    public ReportMeasurementProfile MeasurementProfile { get; } // 計測条件と環境差吸収設定

    public IReadOnlyList<ReportDiagnostic> Diagnostics => diagnostics; // 読み込み時点で収集した診断情報

    public ReportSheet AddSheet(string name)
    {
        var sheet = new ReportSheet(name);
        AddSheet(sheet);
        return sheet;
    }

    public void AddSheet(ReportSheet sheet)
    {
        sheets.Add(sheet);
    }

    internal void AddDiagnostic(ReportDiagnostic diagnostic)
    {
        diagnostics.Add(diagnostic);
    }
}

public sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public ReportSheet(string name)
    {
        Name = name;
        UsedRange = new ReportRange(1, 1, 1, 1);
    }

    public string Name { get; } // シート名

    public ReportRange UsedRange { get; private set; } // 使用範囲

    public IReadOnlyList<ReportRow> Rows => rows; // 行定義一覧

    public IReadOnlyList<ReportColumn> Columns => columns; // 列定義一覧

    public IReadOnlyList<ReportCell> Cells => cells; // 使用範囲内のセル一覧

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges; // 結合セル範囲一覧

    public IReadOnlyList<ReportImage> Images => images; // シート上の画像一覧

    public ReportPageSetup PageSetup { get; private set; } = new(); // 印刷時のページ設定

    public ReportHeaderFooter HeaderFooter { get; private set; } = new(); // ヘッダ、フッタ定義

    public ReportPrintArea? PrintArea { get; private set; } // 明示的な印刷範囲

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks; // 水平手動改ページ一覧

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks; // 垂直手動改ページ一覧

    public bool ShowGridLines { get; private set; } // グリッド線表示フラグ

    public int ReplacePlaceholder(string markerName, string value)
    {
        var replaceCount = 0;
        foreach (var cell in cells.Where(static x => x.Placeholder is not null))
        {
            if (!string.Equals(cell.Placeholder!.MarkerName, markerName, StringComparison.Ordinal))
            {
                continue;
            }

            cell.SetDisplayText(value);
            cell.Placeholder.SetResolvedText(value);
            replaceCount++;
        }

        return replaceCount;
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var replaceCount = 0;
        foreach (var (key, value) in values)
        {
            replaceCount += ReplacePlaceholder(key, value ?? string.Empty);
        }

        return replaceCount;
    }

    public void AddRows(RowExpansionRequest request)
    {
        var repeatCount = request.GetRepeatCount();
        var templateRows = rows
            .Where(row => row.Index >= request.TemplateStartRowIndex && row.Index <= request.TemplateEndRowIndex)
            .OrderBy(row => row.Index)
            .ToList();

        if (templateRows.Count == 0)
        {
            throw new InvalidOperationException("Template rows were not found.");
        }

        var blockSize = request.TemplateEndRowIndex - request.TemplateStartRowIndex + 1;
        var additionalRows = repeatCount * blockSize;

        foreach (var row in rows.Where(row => row.Index > request.TemplateEndRowIndex))
        {
            row.SetIndex(row.Index + additionalRows);
        }

        foreach (var cell in cells.Where(cell => cell.Row > request.TemplateEndRowIndex))
        {
            cell.SetRowColumn(cell.Row + additionalRows, cell.Column);
        }

        foreach (var range in mergedRanges.Where(range => range.Range.StartRow > request.TemplateEndRowIndex))
        {
            range.SetRange(range.Range.ShiftRows(additionalRows));
        }

        foreach (var image in images.Where(image => image.FromRow > request.TemplateEndRowIndex))
        {
            image.ShiftRows(additionalRows);
        }

        var insertIndex = rows.FindLastIndex(row => row.Index <= request.TemplateEndRowIndex) + 1;
        var templateCells = cells
            .Where(cell => cell.Row >= request.TemplateStartRowIndex && cell.Row <= request.TemplateEndRowIndex)
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .ToList();

        for (var iteration = 0; iteration < repeatCount; iteration++)
        {
            var rowOffset = blockSize * (iteration + 1);
            foreach (var templateRow in templateRows)
            {
                rows.Insert(insertIndex++, templateRow.CloneWithIndex(templateRow.Index + rowOffset));
            }

            foreach (var templateCell in templateCells)
            {
                var clone = templateCell.CloneWithPosition(templateCell.Row + rowOffset, templateCell.Column);
                var placeholderValues = request.GetPlaceholderValues(iteration);
                if (clone.Placeholder is not null &&
                    placeholderValues.TryGetValue(clone.Placeholder.MarkerName, out var replacement))
                {
                    var resolvedText = replacement ?? string.Empty;
                    clone.SetDisplayText(resolvedText);
                    clone.Placeholder.SetResolvedText(resolvedText);
                }

                cells.Add(clone);
            }

            foreach (var templateRange in mergedRanges.Where(range => range.Range.StartRow >= request.TemplateStartRowIndex && range.Range.EndRow <= request.TemplateEndRowIndex).ToList())
            {
                mergedRanges.Add(templateRange.CloneShifted(rowOffset));
            }

            foreach (var templateImage in images.Where(image => image.FromRow >= request.TemplateStartRowIndex && image.FromRow <= request.TemplateEndRowIndex).ToList())
            {
                images.Add(templateImage.CloneShifted(rowOffset));
            }
        }

        rows.Sort(static (left, right) => left.Index.CompareTo(right.Index));
        cells.Sort(static (left, right) =>
        {
            var rowCompare = left.Row.CompareTo(right.Row);
            return rowCompare != 0 ? rowCompare : left.Column.CompareTo(right.Column);
        });
        mergedRanges.Sort(static (left, right) => left.Range.StartRow.CompareTo(right.Range.StartRow));

        UpdateUsedRange();
        RecalculateLayout();
    }

    internal void AddRowDefinition(ReportRow row) => rows.Add(row);

    internal void AddColumnDefinition(ReportColumn column) => columns.Add(column);

    internal void AddCell(ReportCell cell) => cells.Add(cell);

    internal void AddMergedRange(ReportMergedRange range) => mergedRanges.Add(range);

    internal void AddImage(ReportImage image) => images.Add(image);

    internal void AddHorizontalPageBreak(ReportPageBreak pageBreak) => horizontalPageBreaks.Add(pageBreak);

    internal void AddVerticalPageBreak(ReportPageBreak pageBreak) => verticalPageBreaks.Add(pageBreak);

    internal void SetPageSetup(ReportPageSetup pageSetup) => PageSetup = pageSetup;

    internal void SetHeaderFooter(ReportHeaderFooter headerFooter) => HeaderFooter = headerFooter;

    internal void SetPrintArea(ReportPrintArea? printArea) => PrintArea = printArea;

    internal void SetShowGridLines(bool showGridLines) => ShowGridLines = showGridLines;

    internal void SetUsedRange(ReportRange usedRange) => UsedRange = usedRange;

    internal void RecalculateLayout()
    {
        var top = 0d;
        foreach (var row in rows.OrderBy(static row => row.Index))
        {
            row.SetTop(top);
            top += row.HeightPoint;
        }

        var left = 0d;
        foreach (var column in columns.OrderBy(static column => column.Index))
        {
            column.SetLeft(left);
            left += column.WidthPoint;
        }

        foreach (var cell in cells)
        {
            var row = rows.FirstOrDefault(item => item.Index == cell.Row);
            var column = columns.FirstOrDefault(item => item.Index == cell.Column);
            if (row is null || column is null)
            {
                continue;
            }

            cell.SetBounds(new ReportRect
            {
                X = column.LeftPoint,
                Y = row.TopPoint,
                Width = column.WidthPoint,
                Height = row.HeightPoint
            });
        }
    }

    private void UpdateUsedRange()
    {
        if (cells.Count == 0)
        {
            UsedRange = new ReportRange(1, 1, 1, 1);
            return;
        }

        UsedRange = new ReportRange(
            cells.Min(static cell => cell.Row),
            cells.Min(static cell => cell.Column),
            cells.Max(static cell => cell.Row),
            cells.Max(static cell => cell.Column));
    }
}

public sealed class ReportRow
{
    public ReportRow(int index, double heightPoint, bool isHidden = false, int outlineLevel = 0)
    {
        Index = index;
        HeightPoint = heightPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
    }

    public int Index { get; private set; } // 1 始まりの行番号

    public double HeightPoint { get; } // 行高(point)

    public double TopPoint { get; private set; } // シート先頭からの上端位置(point)

    public bool IsHidden { get; } // 非表示行かどうか

    public int OutlineLevel { get; } // アウトラインレベル

    internal ReportRow CloneWithIndex(int index) => new(index, HeightPoint, IsHidden, OutlineLevel);

    internal void SetIndex(int index) => Index = index;

    internal void SetTop(double topPoint) => TopPoint = topPoint;
}

public sealed class ReportColumn
{
    public ReportColumn(int index, double widthPoint, bool isHidden = false, int outlineLevel = 0, double originalExcelWidth = 0)
    {
        Index = index;
        WidthPoint = widthPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
        OriginalExcelWidth = originalExcelWidth;
    }

    public int Index { get; } // 1 始まりの列番号

    public double WidthPoint { get; } // 列幅(point)

    public double LeftPoint { get; private set; } // シート左端からの左端位置(point)

    public bool IsHidden { get; } // 非表示列かどうか

    public int OutlineLevel { get; } // アウトラインレベル

    public double OriginalExcelWidth { get; } // Excel 上の元列幅値

    internal void SetLeft(double leftPoint) => LeftPoint = leftPoint;
}

public sealed record ReportCellValue
{
    public ReportCellValueKind Kind { get; init; } // 元データの種別

    public object? RawValue { get; init; } // Excel から取得した元値
}

public sealed class ReportCell
{
    public ReportCell(
        int row,
        int column,
        ReportCellValue value,
        string sourceText,
        string displayText,
        ReportCellStyle style,
        ReportPlaceholderText? placeholder = null)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
        Value = value;
        SourceText = sourceText;
        DisplayText = displayText;
        Placeholder = placeholder;
        Style = style;
    }

    public int Row { get; private set; } // 1 始まりの行番号

    public int Column { get; private set; } // 1 始まりの列番号

    public string Address { get; private set; } // A1 形式のセル番地

    public ReportCellValue Value { get; } // 元データ値

    public string SourceText { get; } // Excel から読んだ元の表示文字列

    public string DisplayText { get; private set; } // 現在の表示文字列

    public ReportPlaceholderText? Placeholder { get; } // プレースホルダ情報

    public ReportCellStyle Style { get; private set; } // セルスタイル

    public ReportRect Bounds { get; private set; } // セル外枠の物理矩形

    public ReportMergeInfo? Merge { get; private set; } // 結合セル参加情報

    internal void SetDisplayText(string displayText) => DisplayText = displayText;

    internal void SetBounds(ReportRect bounds) => Bounds = bounds;

    internal void SetMerge(ReportMergeInfo? merge) => Merge = merge;

    internal void SetStyle(ReportCellStyle style) => Style = style;

    internal void SetRowColumn(int row, int column)
    {
        Row = row;
        Column = column;
        Address = AddressHelper.ToAddress(row, column);
    }

    internal ReportCell CloneWithPosition(int row, int column)
    {
        var placeholder = Placeholder?.Clone();
        return new ReportCell(row, column, Value, SourceText, DisplayText, Style, placeholder)
        {
            Bounds = Bounds,
            Merge = Merge
        };
    }
}

public sealed class ReportPlaceholderText
{
    public ReportPlaceholderText(string markerText, string markerName)
    {
        MarkerText = markerText;
        MarkerName = markerName;
    }

    public string MarkerText { get; } // Excel 上の特殊値そのもの

    public string MarkerName { get; } // アプリケーションから指定する識別子

    public string? ResolvedText { get; private set; } // 置換後の表示文字列

    internal ReportPlaceholderText Clone() =>
        new(MarkerText, MarkerName)
        {
            ResolvedText = ResolvedText
        };

    internal void SetResolvedText(string? text) => ResolvedText = text;
}

public sealed record ReportMergeInfo
{
    public string OwnerCellAddress { get; init; } = string.Empty; // 結合セルの代表セル番地

    public bool IsOwner { get; init; } // 結合セルの代表セルかどうか

    public ReportRange Range { get; init; } // 結合セル範囲
}

public sealed class ReportMergedRange
{
    public ReportMergedRange(ReportRange range)
    {
        Range = range;
        OwnerCellAddress = AddressHelper.ToAddress(range.StartRow, range.StartColumn);
    }

    public ReportRange Range { get; private set; } // 結合セル範囲

    public string OwnerCellAddress { get; } // 代表セル番地

    internal ReportMergedRange CloneShifted(int rowOffset) => new(Range.ShiftRows(rowOffset));

    internal void SetRange(ReportRange range) => Range = range;
}

public sealed record RowExpansionRequest
{
    public int TemplateStartRowIndex { get; init; } // 繰り返し元の開始行番号

    public int TemplateEndRowIndex { get; init; } // 繰り返し元の終了行番号

    public int RepeatCount { get; init; } // 追加する繰り返し回数

    public IReadOnlyList<IReadOnlyDictionary<string, string?>> PlaceholderValuesByIteration { get; init; } =
        Array.Empty<IReadOnlyDictionary<string, string?>>(); // 各繰り返し行に適用するプレースホルダ値

    internal int GetRepeatCount()
    {
        if (RepeatCount > 0)
        {
            return RepeatCount;
        }

        if (PlaceholderValuesByIteration.Count > 0)
        {
            return PlaceholderValuesByIteration.Count;
        }

        throw new InvalidOperationException("RepeatCount or PlaceholderValuesByIteration must be specified.");
    }

    internal IReadOnlyDictionary<string, string?> GetPlaceholderValues(int iteration) =>
        iteration < PlaceholderValuesByIteration.Count
            ? PlaceholderValuesByIteration[iteration]
            : new Dictionary<string, string?>();
}
