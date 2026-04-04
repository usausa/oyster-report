namespace OysterReport.Model;

using OysterReport.Common;
using OysterReport.Common.Geometry;

public sealed record ReportPageSetup
{
    public ReportPaperSize PaperSize { get; init; } = ReportPaperSize.A4; // 用紙サイズ

    public ReportPageOrientation Orientation { get; init; } = ReportPageOrientation.Portrait; // 用紙向き

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d }; // 本文余白

    public double HeaderMarginPoint { get; init; } = 18d; // ヘッダ余白(point)

    public double FooterMarginPoint { get; init; } = 18d; // フッタ余白(point)

    public int ScalePercent { get; init; } = 100; // 印刷倍率(%)

    public int? FitToPagesWide { get; init; } // 横方向の目標ページ数

    public int? FitToPagesTall { get; init; } // 縦方向の目標ページ数

    public bool CenterHorizontally { get; init; } // 水平中央寄せフラグ

    public bool CenterVertically { get; init; } // 垂直中央寄せフラグ
}

public sealed record ReportHeaderFooter
{
    public bool AlignWithMargins { get; init; } = true; // 余白に合わせるか

    public bool DifferentFirst { get; init; } // 先頭ページを別定義にするか

    public bool DifferentOddEven { get; init; } // 奇数偶数ページを別定義にするか

    public bool ScaleWithDocument { get; init; } = true; // 本文の拡大縮小に追従するか

    public string? OddHeader { get; init; } // 通常ページのヘッダ原文

    public string? OddFooter { get; init; } // 通常ページのフッタ原文

    public string? EvenHeader { get; init; } // 偶数ページのヘッダ原文

    public string? EvenFooter { get; init; } // 偶数ページのフッタ原文

    public string? FirstHeader { get; init; } // 先頭ページのヘッダ原文

    public string? FirstFooter { get; init; } // 先頭ページのフッタ原文
}

public sealed record ReportPrintArea
{
    public ReportRange Range { get; init; } // 印刷範囲
}

public sealed record ReportPageBreak
{
    public int Index { get; init; } // 改ページ位置の行番号または列番号

    public bool IsHorizontal { get; init; } // 水平方向改ページかどうか
}

public sealed class ReportImage
{
    public ReportImage(
        string name,
        ReportAnchorType anchorType,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        AnchorType = anchorType;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;

        var (row, _) = OysterReport.Common.AddressHelper.ParseAddress(fromCellAddress);
        FromRow = row;
    }

    public string Name { get; } // 画像識別名

    public ReportAnchorType AnchorType { get; } // アンカー種別

    public string FromCellAddress { get; private set; } // 開始セル番地

    public string? ToCellAddress { get; private set; } // 終了セル番地

    public ReportOffset Offset { get; } // セル内オフセット

    public double WidthPoint { get; } // 画像幅(point)

    public double HeightPoint { get; } // 画像高さ(point)

    public ReadOnlyMemory<byte> ImageBytes { get; } // 元画像データ

    internal int FromRow { get; private set; } // 開始行番号

    internal ReportImage CloneShifted(int rowOffset)
    {
        var (_, fromColumn) = OysterReport.Common.AddressHelper.ParseAddress(FromCellAddress);
        var shiftedFrom = OysterReport.Common.AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        string? shiftedTo = null;
        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = OysterReport.Common.AddressHelper.ParseAddress(ToCellAddress);
            shiftedTo = OysterReport.Common.AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }

        return new ReportImage(Name, AnchorType, shiftedFrom, shiftedTo, Offset, WidthPoint, HeightPoint, ImageBytes);
    }

    internal void ShiftRows(int rowOffset)
    {
        var (_, fromColumn) = OysterReport.Common.AddressHelper.ParseAddress(FromCellAddress);
        FromCellAddress = OysterReport.Common.AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        FromRow += rowOffset;

        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = OysterReport.Common.AddressHelper.ParseAddress(ToCellAddress);
            ToCellAddress = OysterReport.Common.AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }
    }
}
