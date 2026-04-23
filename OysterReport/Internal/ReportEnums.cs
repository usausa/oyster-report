namespace OysterReport.Internal;

internal enum CellValueKind
{
    Blank = 0,
    Boolean,
    Number,
    Text,
    Error,
    DateTime,
    TimeSpan
}

internal enum BorderLineStyle
{
    None = 0,
    DashDot,
    DashDotDot,
    Dashed,
    Dotted,
    Double,
    Hair,
    Medium,
    MediumDashDot,
    MediumDashDotDot,
    MediumDashed,
    SlantDashDot,
    Thick,
    Thin
}

internal enum HorizontalAlignment
{
    General = 0,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed
}

internal enum VerticalAlignment
{
    Top = 0,
    Center,
    Bottom,
    Justify,
    Distributed
}

internal enum PageOrientation
{
    Default = 0,
    Portrait,
    Landscape
}

internal enum FillPattern
{
    None = 0,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625
}
