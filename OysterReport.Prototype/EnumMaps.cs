// Mapping helpers between OpenXML SDK enums/strings and the existing ClosedXML-based enums used in Internal/Models.cs.

namespace OysterReport.Prototype;

using ClosedXML.Excel;

using DocumentFormat.OpenXml.Spreadsheet;

internal static class EnumMaps
{
    public static XLBorderStyleValues ToBorderStyle(BorderStyleValues? style)
    {
        if (style is null)
        {
            return XLBorderStyleValues.None;
        }

        var raw = style.Value;
        return raw switch
        {
            var v when v == BorderStyleValues.Thin => XLBorderStyleValues.Thin,
            var v when v == BorderStyleValues.Medium => XLBorderStyleValues.Medium,
            var v when v == BorderStyleValues.Thick => XLBorderStyleValues.Thick,
            var v when v == BorderStyleValues.Double => XLBorderStyleValues.Double,
            var v when v == BorderStyleValues.Hair => XLBorderStyleValues.Hair,
            var v when v == BorderStyleValues.Dotted => XLBorderStyleValues.Dotted,
            var v when v == BorderStyleValues.Dashed => XLBorderStyleValues.Dashed,
            var v when v == BorderStyleValues.DashDot => XLBorderStyleValues.DashDot,
            var v when v == BorderStyleValues.DashDotDot => XLBorderStyleValues.DashDotDot,
            var v when v == BorderStyleValues.MediumDashed => XLBorderStyleValues.MediumDashed,
            var v when v == BorderStyleValues.MediumDashDot => XLBorderStyleValues.MediumDashDot,
            var v when v == BorderStyleValues.MediumDashDotDot => XLBorderStyleValues.MediumDashDotDot,
            var v when v == BorderStyleValues.SlantDashDot => XLBorderStyleValues.SlantDashDot,
            _ => XLBorderStyleValues.None
        };
    }

    public static XLAlignmentHorizontalValues ToHorizontalAlignment(HorizontalAlignmentValues? value)
    {
        if (value is null)
        {
            return XLAlignmentHorizontalValues.General;
        }

        var raw = value.Value;
        return raw switch
        {
            var v when v == HorizontalAlignmentValues.Left => XLAlignmentHorizontalValues.Left,
            var v when v == HorizontalAlignmentValues.Center => XLAlignmentHorizontalValues.Center,
            var v when v == HorizontalAlignmentValues.Right => XLAlignmentHorizontalValues.Right,
            var v when v == HorizontalAlignmentValues.Justify => XLAlignmentHorizontalValues.Justify,
            var v when v == HorizontalAlignmentValues.CenterContinuous => XLAlignmentHorizontalValues.CenterContinuous,
            var v when v == HorizontalAlignmentValues.Distributed => XLAlignmentHorizontalValues.Distributed,
            var v when v == HorizontalAlignmentValues.Fill => XLAlignmentHorizontalValues.Fill,
            _ => XLAlignmentHorizontalValues.General
        };
    }

    public static XLAlignmentVerticalValues ToVerticalAlignment(VerticalAlignmentValues? value)
    {
        // Excel's default when vertical alignment is unspecified is Bottom (matches ClosedXML's reported default).
        if (value is null)
        {
            return XLAlignmentVerticalValues.Bottom;
        }

        var raw = value.Value;
        return raw switch
        {
            var v when v == VerticalAlignmentValues.Top => XLAlignmentVerticalValues.Top,
            var v when v == VerticalAlignmentValues.Center => XLAlignmentVerticalValues.Center,
            var v when v == VerticalAlignmentValues.Justify => XLAlignmentVerticalValues.Justify,
            var v when v == VerticalAlignmentValues.Distributed => XLAlignmentVerticalValues.Distributed,
            _ => XLAlignmentVerticalValues.Bottom
        };
    }

    public static XLPaperSize ToPaperSize(uint? code) =>
        code is null ? XLPaperSize.A4Paper : (XLPaperSize)code.Value;

    public static XLPageOrientation ToPageOrientation(OrientationValues? value)
    {
        if (value is null)
        {
            return XLPageOrientation.Default;
        }

        var raw = value.Value;
        return raw switch
        {
            var v when v == OrientationValues.Portrait => XLPageOrientation.Portrait,
            var v when v == OrientationValues.Landscape => XLPageOrientation.Landscape,
            _ => XLPageOrientation.Default
        };
    }

    public static XLFillPatternValues ToFillPattern(PatternValues? value)
    {
        if (value is null)
        {
            return XLFillPatternValues.None;
        }

        var raw = value.Value;
        return raw switch
        {
            var v when v == PatternValues.None => XLFillPatternValues.None,
            var v when v == PatternValues.Solid => XLFillPatternValues.Solid,
            var v when v == PatternValues.Gray125 => XLFillPatternValues.Gray125,
            var v when v == PatternValues.Gray0625 => XLFillPatternValues.Gray0625,
            var v when v == PatternValues.DarkGray => XLFillPatternValues.DarkGray,
            var v when v == PatternValues.MediumGray => XLFillPatternValues.MediumGray,
            var v when v == PatternValues.LightGray => XLFillPatternValues.LightGray,
            var v when v == PatternValues.DarkHorizontal => XLFillPatternValues.DarkHorizontal,
            var v when v == PatternValues.DarkVertical => XLFillPatternValues.DarkVertical,
            var v when v == PatternValues.DarkDown => XLFillPatternValues.DarkDown,
            var v when v == PatternValues.DarkUp => XLFillPatternValues.DarkUp,
            var v when v == PatternValues.DarkGrid => XLFillPatternValues.DarkGrid,
            var v when v == PatternValues.DarkTrellis => XLFillPatternValues.DarkTrellis,
            var v when v == PatternValues.LightHorizontal => XLFillPatternValues.LightHorizontal,
            var v when v == PatternValues.LightVertical => XLFillPatternValues.LightVertical,
            var v when v == PatternValues.LightDown => XLFillPatternValues.LightDown,
            var v when v == PatternValues.LightUp => XLFillPatternValues.LightUp,
            var v when v == PatternValues.LightGrid => XLFillPatternValues.LightGrid,
            var v when v == PatternValues.LightTrellis => XLFillPatternValues.LightTrellis,
            _ => XLFillPatternValues.None
        };
    }
}
