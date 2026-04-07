namespace OysterReport.Internal;

using OysterReport;

internal sealed record PdfRenderPlan
{
    public IReadOnlyList<PdfRenderSheetPlan> Sheets { get; init; } = Array.Empty<PdfRenderSheetPlan>(); // 解決済みシートレンダリング情報一覧
}

internal sealed record PdfRenderSheetPlan
{
    public string SheetName { get; init; } = string.Empty; // 対象シート名

    public IReadOnlyList<PdfRenderPagePlan> Pages { get; init; } = Array.Empty<PdfRenderPagePlan>(); // ページ分割後のページ一覧

    public IReadOnlyList<PdfBorderRenderInfo> Borders { get; init; } = Array.Empty<PdfBorderRenderInfo>(); // 罫線競合解決後の罫線一覧

    public IReadOnlyList<PdfImageRenderInfo> Images { get; init; } = Array.Empty<PdfImageRenderInfo>(); // 画像の最終配置一覧
}

internal sealed record PdfRenderPagePlan
{
    public int PageNumber { get; init; } // 1 始まりのページ番号

    public ReportRect PageBounds { get; init; } // 用紙全体の矩形

    public ReportRect PrintableBounds { get; init; } // 余白を除いた印字可能領域

    public PdfHeaderFooterRenderInfo HeaderFooter { get; init; } = new(); // 当該ページのヘッダ、フッタ描画情報

    public IReadOnlyList<PdfCellRenderInfo> Cells { get; init; } = Array.Empty<PdfCellRenderInfo>(); // 当該ページに描画するセル一覧
}

internal sealed record PdfCellRenderInfo
{
    public string CellAddress { get; init; } = string.Empty; // 対象セルの番地

    public ReportRect OuterBounds { get; init; } // セル外枠の最終矩形

    public ReportRect ContentBounds { get; init; } // 内容描画領域の最終矩形

    public ReportRect TextBounds { get; init; } // テキスト描画に使う最終矩形

    public bool IsMergedOwner { get; init; } // 結合セルの代表セルかどうか

    public bool IsClipped { get; init; } // 描画時にクリップが必要かどうか
}

internal sealed record PdfBorderRenderInfo
{
    public ReportLine Line { get; init; } // 描画する線分

    public ReportBorderStyle Style { get; init; } // 線分に適用する罫線スタイル

    public double Width { get; init; } // 線分に適用する罫線幅

    public string ColorHex { get; init; } = "#FF000000"; // 線分に適用する罫線色

    public string OwnerCellAddress { get; init; } = string.Empty; // この線分の由来となる代表セル番地
}

internal sealed record PdfImageRenderInfo
{
    public string Name { get; init; } = string.Empty; // 画像識別名

    public ReportRect Bounds { get; init; } // 描画に使う最終矩形

    public ReadOnlyMemory<byte> ImageBytes { get; init; } // 画像バイトデータ
}

internal sealed record PdfHeaderFooterRenderInfo
{
    public string? HeaderText { get; init; } // 当該ページに描画するヘッダ文字列

    public string? FooterText { get; init; } // 当該ページに描画するフッタ文字列

    public ReportRect HeaderBounds { get; init; } // ヘッダ描画領域

    public ReportRect FooterBounds { get; init; } // フッタ描画領域
}
