namespace OysterReport.Model;

using OysterReport.Common;

public sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new(); // 文字フォント設定

    public ReportFill Fill { get; init; } = new(); // 背景塗りつぶし設定

    public ReportBorders Borders { get; init; } = new(); // 四辺の罫線設定

    public ReportAlignment Alignment { get; init; } = new(); // 水平、垂直配置設定

    public string? NumberFormat { get; init; } // Excel の表示書式

    public bool WrapText { get; init; } // 折り返し表示フラグ

    public double Rotation { get; init; } // 文字回転角度

    public bool ShrinkToFit { get; init; } // 縮小して全体表示するか
}

public sealed record ReportFont
{
    public string Name { get; init; } = "Arial"; // フォント名

    public double Size { get; init; } = 11d; // フォントサイズ(point)

    public bool Bold { get; init; } // 太字フラグ

    public bool Italic { get; init; } // 斜体フラグ

    public bool Underline { get; init; } // 下線フラグ

    public bool Strikeout { get; init; } // 打消し線フラグ

    public string ColorHex { get; init; } = "#FF000000"; // 文字色
}

public sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000"; // 背景色
}

public sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new(); // 左罫線

    public ReportBorder Top { get; init; } = new(); // 上罫線

    public ReportBorder Right { get; init; } = new(); // 右罫線

    public ReportBorder Bottom { get; init; } = new(); // 下罫線
}

public sealed record ReportBorder
{
    public ReportBorderStyle Style { get; init; } // 罫線種別

    public string ColorHex { get; init; } = "#FF000000"; // 罫線色

    public double Width { get; init; } = 0.5d; // 罫線幅(point)
}

public sealed record ReportAlignment
{
    public ReportHorizontalAlignment Horizontal { get; init; } = ReportHorizontalAlignment.General; // 水平配置

    public ReportVerticalAlignment Vertical { get; init; } = ReportVerticalAlignment.Top; // 垂直配置
}
