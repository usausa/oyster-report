namespace OysterReport.Writing.Pdf;

using System.Globalization;
using System.Reflection;
using System.Text;
using OysterReport.Common;
using OysterReport.Common.Geometry;
using OysterReport.Helpers;
using OysterReport.Internal.Rendering;
using OysterReport.Model;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

public sealed class PdfGenerator
{
    public IReportFontResolver? DefaultFontResolver { get; set; } // 既定のフォントリゾルバ

    private static int fontPlatformConfigured;

    private static readonly string[] FallbackFontNames =
    [
        "Arial",
        "Helvetica",
        "Segoe UI",
        "Liberation Sans",
        "DejaVu Sans",
        "Noto Sans",
        "Times New Roman",
        "Courier New",
    ];

    public void Generate(
        ReportWorkbook workbook,
        Stream output,
        PdfGenerateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(output);

        EnsurePdfSharpFontConfiguration();

        var effectiveOptions = options ?? new PdfGenerateOptions();
        effectiveOptions.FontResolver ??= DefaultFontResolver;

        var renderPlan = BuildRenderPlan(workbook, effectiveOptions);
        WritePdf(workbook, renderPlan, output, effectiveOptions);
    }

    internal static PdfRenderPlan BuildRenderPlan(
        ReportWorkbook workbook,
        PdfGenerateOptions options)
    {
        return PdfRenderPlanner.BuildPlan(workbook, options);
    }

    internal static void WritePdf(
        ReportWorkbook workbook,
        PdfRenderPlan renderPlan,
        Stream output,
        PdfGenerateOptions options)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(renderPlan);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(options);

        using var document = new PdfDocument
        {
            Options =
            {
                CompressContentStreams = options.CompressContentStreams,
            },
        };

        if (options.EmbedDocumentMetadata)
        {
            document.Info.Title = workbook.Metadata.TemplateName;
        }

        for (var sheetIndex = 0; sheetIndex < renderPlan.Sheets.Count; sheetIndex++)
        {
            var sheetPlan = renderPlan.Sheets[sheetIndex];
            var sourceSheet = workbook.Sheets[sheetIndex];
            foreach (var pagePlan in sheetPlan.Pages)
            {
                var page = document.AddPage();
                page.Width = XUnit.FromPoint(pagePlan.PageBounds.Width);
                page.Height = XUnit.FromPoint(pagePlan.PageBounds.Height);
                using var graphics = XGraphics.FromPdfPage(page);
                DrawPageBackground(graphics, pagePlan.PageBounds);
                DrawCells(graphics, sourceSheet, pagePlan.Cells, options);
                DrawBorders(graphics, sheetPlan.Borders);
                DrawImages(graphics, sheetPlan.Images);
                DrawHeaderFooter(graphics, pagePlan.HeaderFooter, pagePlan.PageNumber, sheetPlan.Pages.Count);
            }
        }

        document.Save(output, closeStream: false);
    }

    private static void EnsurePdfSharpFontConfiguration()
    {
        if (Interlocked.Exchange(ref fontPlatformConfigured, 1) == 1)
        {
            return;
        }

        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var globalFontSettingsType = typeof(XFont).Assembly.GetType("PdfSharp.Fonts.GlobalFontSettings");
        if (globalFontSettingsType is null)
        {
            return;
        }

        var fontResolverProperty = globalFontSettingsType.GetProperty("FontResolver", BindingFlags.Public | BindingFlags.Static);
        var fallbackFontResolverProperty = globalFontSettingsType.GetProperty("FallbackFontResolver", BindingFlags.Public | BindingFlags.Static);
        var useWindowsFontsProperty = globalFontSettingsType.GetProperty("UseWindowsFontsUnderWindows", BindingFlags.Public | BindingFlags.Static);

        if (fontResolverProperty?.GetValue(null) is null && fallbackFontResolverProperty?.GetValue(null) is null)
        {
            useWindowsFontsProperty?.SetValue(null, true);
        }
    }

    private static void DrawPageBackground(XGraphics graphics, ReportRect pageBounds)
    {
        graphics.DrawRectangle(XBrushes.White, pageBounds.X, pageBounds.Y, pageBounds.Width, pageBounds.Height);
    }

    private static void DrawCells(
        XGraphics graphics,
        ReportSheet sourceSheet,
        IReadOnlyList<PdfCellRenderInfo> cells,
        PdfGenerateOptions options)
    {
        foreach (var renderCell in cells)
        {
            var sourceCell = sourceSheet.Cells.First(cell => cell.Address == renderCell.CellAddress);
            if (!string.Equals(sourceCell.Style.Fill.BackgroundColorHex, "#00000000", StringComparison.Ordinal))
            {
                var backgroundBrush = new XSolidBrush(ToColor(sourceCell.Style.Fill.BackgroundColorHex));
                graphics.DrawRectangle(
                    backgroundBrush,
                    renderCell.OuterBounds.X,
                    renderCell.OuterBounds.Y,
                    renderCell.OuterBounds.Width,
                    renderCell.OuterBounds.Height);
            }

            if (!renderCell.IsMergedOwner && sourceCell.Merge is not null)
            {
                continue;
            }

            if (string.IsNullOrEmpty(sourceCell.DisplayText))
            {
                continue;
            }

            var font = ResolveFont(sourceCell.Style.Font, options);
            var textBrush = new XSolidBrush(ToColor(sourceCell.Style.Font.ColorHex));
            graphics.DrawString(
                sourceCell.DisplayText,
                font,
                textBrush,
                new XRect(
                    renderCell.TextBounds.X,
                    renderCell.TextBounds.Y,
                    Math.Max(0, renderCell.ContentBounds.Width),
                    Math.Max(0, renderCell.ContentBounds.Height)),
                XStringFormats.TopLeft);
        }
    }

    private static void DrawBorders(XGraphics graphics, IReadOnlyList<PdfBorderRenderInfo> borders)
    {
        foreach (var border in borders)
        {
            var pen = new XPen(ToColor(border.ColorHex), ResolveBorderWidth(border.Style));
            graphics.DrawLine(pen, border.Line.X1, border.Line.Y1, border.Line.X2, border.Line.Y2);
        }
    }

    private static void DrawImages(XGraphics graphics, IReadOnlyList<PdfImageRenderInfo> images)
    {
        foreach (var image in images)
        {
            graphics.DrawRectangle(XPens.LightGray, image.Bounds.X, image.Bounds.Y, image.Bounds.Width, image.Bounds.Height);
            graphics.DrawString(image.Name, CreateFallbackFont(8d), XBrushes.Gray, image.Bounds.X + 2d, image.Bounds.Y + 10d);
        }
    }

    private static void DrawHeaderFooter(
        XGraphics graphics,
        PdfHeaderFooterRenderInfo headerFooter,
        int pageNumber,
        int totalPages)
    {
        var headerSections = ResolveHeaderFooterSections(headerFooter.HeaderText, pageNumber, totalPages);
        var footerSections = ResolveHeaderFooterSections(headerFooter.FooterText, pageNumber, totalPages);
        var font = CreateFallbackFont(9d);

        DrawHeaderFooterSections(graphics, headerSections, headerFooter.HeaderBounds, font);
        DrawHeaderFooterSections(graphics, footerSections, headerFooter.FooterBounds, font);
    }

    private static void DrawHeaderFooterSections(
        XGraphics graphics,
        HeaderFooterSections sections,
        ReportRect bounds,
        XFont font)
    {
        if (!string.IsNullOrWhiteSpace(sections.Left))
        {
            graphics.DrawString(
                sections.Left,
                font,
                XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height),
                XStringFormats.TopLeft);
        }

        if (!string.IsNullOrWhiteSpace(sections.Center))
        {
            graphics.DrawString(
                sections.Center,
                font,
                XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height),
                XStringFormats.TopCenter);
        }

        if (!string.IsNullOrWhiteSpace(sections.Right))
        {
            graphics.DrawString(
                sections.Right,
                font,
                XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height),
                XStringFormats.TopRight);
        }
    }

    private static HeaderFooterSections ResolveHeaderFooterSections(string? text, int pageNumber, int totalPages)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return HeaderFooterSections.Empty;
        }

        var left = new StringBuilder();
        var center = new StringBuilder();
        var right = new StringBuilder();
        var current = center;

        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character != '&' || index == text.Length - 1)
            {
                current.Append(character);
                continue;
            }

            index++;
            switch (char.ToUpperInvariant(text[index]))
            {
                case 'L':
                    current = left;
                    break;
                case 'C':
                    current = center;
                    break;
                case 'R':
                    current = right;
                    break;
                case 'P':
                    current.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
                    break;
                case 'N':
                    current.Append(totalPages.ToString(CultureInfo.InvariantCulture));
                    break;
                case '&':
                    current.Append('&');
                    break;
                default:
                    break;
            }
        }

        return new HeaderFooterSections(
            left.ToString().Trim(),
            center.ToString().Trim(),
            right.ToString().Trim());
    }

    private static XFont ResolveFont(ReportFont font, PdfGenerateOptions options)
    {
        var fontSize = font.Size <= 0 ? 11d : font.Size;
        var style = XFontStyleEx.Regular;
        if (font.Bold)
        {
            style |= XFontStyleEx.Bold;
        }

        if (font.Italic)
        {
            style |= XFontStyleEx.Italic;
        }

        foreach (var fontName in EnumerateCandidateFontNames(font, options))
        {
            if (TryCreateFont(fontName, fontSize, style, out var resolvedFont))
            {
                return resolvedFont;
            }
        }

        throw new InvalidOperationException($"No appropriate font found for family name '{font.Name}' and known fallbacks.");
    }

    private static IEnumerable<string> EnumerateCandidateFontNames(ReportFont font, PdfGenerateOptions options)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (options.FontResolver is not null)
        {
            var resolution = options.FontResolver.Resolve(new ReportFontRequest
            {
                FontName = font.Name,
                Bold = font.Bold,
                Italic = font.Italic,
            });

            if (resolution.IsResolved && !string.IsNullOrWhiteSpace(resolution.ResolvedFontName) && seen.Add(resolution.ResolvedFontName))
            {
                yield return resolution.ResolvedFontName;
            }
        }

        if (!string.IsNullOrWhiteSpace(font.Name) && seen.Add(font.Name))
        {
            yield return font.Name;
        }

        if (string.Equals(font.Name, "Calibri", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var preferredFallback in new[] { "Arial", "Segoe UI", "Helvetica" })
            {
                if (seen.Add(preferredFallback))
                {
                    yield return preferredFallback;
                }
            }
        }

        foreach (var fallbackFontName in FallbackFontNames)
        {
            if (seen.Add(fallbackFontName))
            {
                yield return fallbackFontName;
            }
        }
    }

    private static XFont CreateFallbackFont(double size)
    {
        foreach (var fontName in FallbackFontNames)
        {
            if (TryCreateFont(fontName, size, XFontStyleEx.Regular, out var font))
            {
                return font;
            }
        }

        throw new InvalidOperationException("No appropriate fallback font found for header or image drawing.");
    }

    private static bool TryCreateFont(string fontName, double size, XFontStyleEx style, out XFont font)
    {
        try
        {
            font = new XFont(fontName, size, style);
            return true;
        }
        catch (InvalidOperationException)
        {
        }
        catch (ArgumentException)
        {
        }

        font = null!;
        return false;
    }

    private static XColor ToColor(string colorHex)
    {
        var normalized = ColorHelper.NormalizeHex(colorHex).TrimStart('#');
        if (normalized.Length == 8)
        {
            return XColor.FromArgb(
                Convert.ToByte(normalized[..2], 16),
                Convert.ToByte(normalized.Substring(2, 2), 16),
                Convert.ToByte(normalized.Substring(4, 2), 16),
                Convert.ToByte(normalized.Substring(6, 2), 16));
        }

        return XColors.Black;
    }

    private static double ResolveBorderWidth(ReportBorderStyle style) =>
        style switch
        {
            ReportBorderStyle.Thick => 2d,
            ReportBorderStyle.Medium => 1d,
            ReportBorderStyle.DoubleLine => 1.5d,
            _ => 0.5d,
        };

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);
    }
}
