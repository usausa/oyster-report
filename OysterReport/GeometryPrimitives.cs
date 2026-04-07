namespace OysterReport;

public readonly record struct ReportRect
{
    public double X { get; init; } // Upper-left X coordinate (points)

    public double Y { get; init; } // Upper-left Y coordinate (points)

    public double Width { get; init; } // Width (points)

    public double Height { get; init; } // Height (points)

    public double Right => X + Width; // Right edge X coordinate (points)

    public double Bottom => Y + Height; // Bottom edge Y coordinate (points)

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
    public double Left { get; init; } // Left margin (points)

    public double Top { get; init; } // Top margin (points)

    public double Right { get; init; } // Right margin (points)

    public double Bottom { get; init; } // Bottom margin (points)

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
    public double X1 { get; init; } // Start point X coordinate (points)

    public double Y1 { get; init; } // Start point Y coordinate (points)

    public double X2 { get; init; } // End point X coordinate (points)

    public double Y2 { get; init; } // End point Y coordinate (points)
}
