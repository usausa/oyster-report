namespace OysterReport.Generator;

using ClosedXML.Excel;

internal static class PdfRenderingConstants
{
    // Excel に近い見た目になるように、セル内テキストの左右余白だけを調整する。
    public const double HorizontalCellTextPaddingPoints = 2d;

    public const double DefaultCellFontSizePoints = 11d;

    public const double HeaderFooterFontSizePoints = 9d;

    public const double ThickBorderWidthPoints = 2.25d;

    public const double MediumBorderWidthPoints = 1.5d;

    public const double NormalBorderWidthPoints = 0.75d;

    public const double HairBorderWidthPoints = 0.25d;

    public const double MinimumDoubleBorderGapPoints = 1.5d;

    public const double DoubleBorderGapWidthMultiplier = 1.5d;

    public const double StraightLineTolerancePoints = 0.01d;

    public static double ResolveBorderWidth(XLBorderStyleValues style) =>
        style switch
        {
            XLBorderStyleValues.Thick => ThickBorderWidthPoints,
            XLBorderStyleValues.Medium => MediumBorderWidthPoints,
            XLBorderStyleValues.Double => NormalBorderWidthPoints,
            XLBorderStyleValues.Hair => HairBorderWidthPoints,
            _ => NormalBorderWidthPoints
        };
}
