namespace OysterReport.Generator.Models;

internal readonly record struct ReportRect
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
}
