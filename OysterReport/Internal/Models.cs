namespace OysterReport.Internal;

using System.Diagnostics.CodeAnalysis;

using ClosedXML.Excel;

// ReportWorkbook
// ├── ReportMetadata                 テンプレート名
// ├── ReportMeasurementProfile       列幅計算用フォントメトリクス
// └── ReportSheet[]
//     ├── ReportRow[]                行高・表示/非表示・アウトラインレベル
//     ├── ReportColumn[]             列幅・表示/非表示
//     ├── ReportCell[]
//     │   ├── ReportCellValue       型別の値
//     │   ├── ReportCellStyle
//     │   │   ├── ReportFont
//     │   │   ├── ReportFill
//     │   │   ├── ReportBorders
//     │   │   └── ReportAlignment
//     │   └── ReportMergeInfo?      マージ先情報
//     ├── ReportMergedRange[]        マージセル範囲
//     ├── ReportImage[]              埋め込み画像
//     ├── ReportPageSetup            用紙・余白・中央揃え
//     ├── ReportHeaderFooter         ヘッダー/フッターテキスト
//     ├── ReportPrintArea?           印刷範囲
//     └── ReportPageBreak[]          水平/垂直改ページ

//--------------------------------------------------------------------------------
// Metadata
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty;
}

//--------------------------------------------------------------------------------
// Measurement profile
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportMeasurementProfile
{
    public double MaxDigitWidth { get; init; } = 7d;

    public double ColumnWidthAdjustment { get; init; } = 1d;
}

//--------------------------------------------------------------------------------
// Cell value and style
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportCellValue
{
    public XLDataType Kind { get; init; } = XLDataType.Blank;

    // TODO: Use the typed source value when adding value-aware formatting or placeholder features.
    public object? RawValue { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportFont
{
    public string Name { get; init; } = "Arial";

    public double Size { get; init; } = 11d;

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    // TODO: Honor Excel underline when PDF text decoration support is added.
    public bool Underline { get; init; }

    // TODO: Honor Excel strikeout when PDF text decoration support is added.
    public bool Strikeout { get; init; }

    public string ColorHex { get; init; } = "#FF000000";
}

[ExcludeFromCodeCoverage]
internal sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000";
}

[ExcludeFromCodeCoverage]
internal sealed record ReportBorder
{
    public XLBorderStyleValues Style { get; init; } = XLBorderStyleValues.None;

    public string ColorHex { get; init; } = "#FF000000";

    public double Width { get; init; } = 0.5d;
}

[ExcludeFromCodeCoverage]
internal sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new();

    public ReportBorder Top { get; init; } = new();

    public ReportBorder Right { get; init; } = new();

    public ReportBorder Bottom { get; init; } = new();
}

[ExcludeFromCodeCoverage]
internal sealed record ReportAlignment
{
    public XLAlignmentHorizontalValues Horizontal { get; init; } = XLAlignmentHorizontalValues.General;

    public XLAlignmentVerticalValues Vertical { get; init; } = XLAlignmentVerticalValues.Top;
}

[ExcludeFromCodeCoverage]
internal sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new();

    public ReportFill Fill { get; init; } = new();

    public ReportBorders Borders { get; init; } = new();

    public ReportAlignment Alignment { get; init; } = new();

    public bool WrapText { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportMergeInfo
{
    // ExcelReader で設定されるが、レンダリングパイプラインは cell.Address を直接使用するため
    // TODO 現時点では参照されない。デバッグや非オーナーセルからオーナーを特定する
    // 将来機能のために保持している。
    public string OwnerCellAddress { get; init; } = string.Empty;

    public ReportRange Range { get; init; }
}

//--------------------------------------------------------------------------------
// Page setup
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportPageBreak
{
    // 改ページが発生する行または列のインデックス。
    // TODO 現在の単一ページレンダリングでは参照されないが、
    // 複数ページ対応を実装する際に HorizontalPageBreaks / VerticalPageBreaks と合わせて使用する。
    public int Index { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportPrintArea
{
    public ReportRange Range { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportHeaderFooter
{
    // TODO: Apply Excel's header/footer margin alignment rules during PDF rendering.
    public bool AlignWithMargins { get; init; } = true;

    public bool DifferentFirst { get; init; }

    public bool DifferentOddEven { get; init; }

    // TODO: Apply Excel's header/footer scaling rule during PDF rendering.
    public bool ScaleWithDocument { get; init; } = true;

    public string? OddHeader { get; init; }

    public string? OddFooter { get; init; }

    public string? EvenHeader { get; init; }

    public string? EvenFooter { get; init; }

    public string? FirstHeader { get; init; }

    public string? FirstFooter { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportPageSetup
{
    public XLPaperSize PaperSize { get; init; } = XLPaperSize.A4Paper;

    public XLPageOrientation Orientation { get; init; } = XLPageOrientation.Default;

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d };

    public double HeaderMarginPoint { get; init; } = 18d;

    public double FooterMarginPoint { get; init; } = 18d;

    // TODO: Apply Excel print scaling when multipage fit/scaling support is implemented.
    public int ScalePercent { get; init; } = 100;

    // TODO: Apply Excel fit-to-page width scaling when multipage fit/scaling support is implemented.
    public int? FitToPagesWide { get; init; }

    // TODO: Apply Excel fit-to-page height scaling when multipage fit/scaling support is implemented.
    public int? FitToPagesTall { get; init; }

    public bool CenterHorizontally { get; init; }

    public bool CenterVertically { get; init; }
}

//--------------------------------------------------------------------------------
// Sheet structure
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed class ReportRow
{
    public int Index { get; init; }

    public double HeightPoint { get; init; }

    public double TopPoint { get; set; }

    public bool IsHidden { get; init; }

    public int OutlineLevel { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportColumn
{
    public int Index { get; init; }

    public double WidthPoint { get; init; }

    public double LeftPoint { get; set; }

    public bool IsHidden { get; init; }

    public int OutlineLevel { get; init; }

    public double OriginalExcelWidth { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportMergedRange
{
    public ReportRange Range { get; init; }

    public string OwnerCellAddress => AddressHelper.ToAddress(Range.StartRow, Range.StartColumn);
}

[ExcludeFromCodeCoverage]
internal sealed class ReportImage
{
    public string Name { get; init; } = string.Empty;

    public string FromCellAddress { get; init; } = string.Empty;

    public string? ToCellAddress { get; init; }

    public ReportOffset Offset { get; init; }

    public double WidthPoint { get; init; }

    public double HeightPoint { get; init; }

    public ReadOnlyMemory<byte> ImageBytes { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportCell
{
    public int Row { get; init; }

    public int Column { get; init; }

    public string Address => AddressHelper.ToAddress(Row, Column);

    public ReportCellValue Value { get; init; } = new();

    public string DisplayText { get; init; } = string.Empty;

    public ReportCellStyle Style { get; set; } = new();

    public ReportMergeInfo? Merge { get; set; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public string Name { get; init; } = string.Empty;

    public ReportRange UsedRange { get; set; } = new() { StartRow = 1, StartColumn = 1, EndRow = 1, EndColumn = 1 };

    public IReadOnlyList<ReportRow> Rows => rows;

    public IReadOnlyList<ReportColumn> Columns => columns;

    public IReadOnlyList<ReportCell> Cells => cells;

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges;

    public IReadOnlyList<ReportImage> Images => images;

    public ReportPageSetup PageSetup { get; set; } = new();

    public ReportHeaderFooter HeaderFooter { get; set; } = new();

    public ReportPrintArea? PrintArea { get; set; }

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks;

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks;

    public bool ShowGridLines { get; set; }

    public void AddRowDefinition(ReportRow row) => rows.Add(row);

    public void AddColumnDefinition(ReportColumn column) => columns.Add(column);

    public void AddCell(ReportCell cell) => cells.Add(cell);

    public void AddMergedRange(ReportMergedRange range) => mergedRanges.Add(range);

    public void AddImage(ReportImage image) => images.Add(image);

    public void AddHorizontalPageBreak(ReportPageBreak pageBreak) => horizontalPageBreaks.Add(pageBreak);

    public void AddVerticalPageBreak(ReportPageBreak pageBreak) => verticalPageBreaks.Add(pageBreak);

    public void RecalculateLayout()
    {
        var top = 0d;
        foreach (var row in rows.OrderBy(static x => x.Index))
        {
            row.TopPoint = top;
            top += row.HeightPoint;
        }

        var left = 0d;
        foreach (var column in columns.OrderBy(static x => x.Index))
        {
            column.LeftPoint = left;
            left += column.WidthPoint;
        }
    }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];

    public IReadOnlyList<ReportSheet> Sheets => sheets;

    public ReportMetadata Metadata { get; init; } = new();

    public ReportMeasurementProfile MeasurementProfile { get; init; } = new();

    public void AddSheet(ReportSheet sheet) => sheets.Add(sheet);
}
