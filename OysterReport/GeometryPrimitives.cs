namespace OysterReport;

public readonly record struct ReportRect
{
    public double X { get; init; } // 左上 X 座標(point)

    public double Y { get; init; } // 左上 Y 座標(point)

    public double Width { get; init; } // 幅(point)

    public double Height { get; init; } // 高さ(point)

    public double Right => X + Width; // 右端 X 座標(point)

    public double Bottom => Y + Height; // 下端 Y 座標(point)

    public ReportRect Deflate(ReportThickness thickness) =>
        new()
        {
            X = X + thickness.Left,
            Y = Y + thickness.Top,
            Width = Math.Max(0, Width - thickness.Left - thickness.Right),
            Height = Math.Max(0, Height - thickness.Top - thickness.Bottom)
        };

    public static ReportRect Union(ReportRect first, ReportRect second)
    {
        var x = Math.Min(first.X, second.X);
        var y = Math.Min(first.Y, second.Y);
        var right = Math.Max(first.Right, second.Right);
        var bottom = Math.Max(first.Bottom, second.Bottom);
        return new ReportRect
        {
            X = x,
            Y = y,
            Width = right - x,
            Height = bottom - y
        };
    }
}

public readonly record struct ReportThickness
{
    public double Left { get; init; } // 左余白(point)

    public double Top { get; init; } // 上余白(point)

    public double Right { get; init; } // 右余白(point)

    public double Bottom { get; init; } // 下余白(point)

    public static ReportThickness Uniform(double value) =>
        new()
        {
            Left = value,
            Top = value,
            Right = value,
            Bottom = value
        };
}

public readonly record struct ReportLine
{
    public double X1 { get; init; } // 始点 X 座標(point)

    public double Y1 { get; init; } // 始点 Y 座標(point)

    public double X2 { get; init; } // 終点 X 座標(point)

    public double Y2 { get; init; } // 終点 Y 座標(point)
}
