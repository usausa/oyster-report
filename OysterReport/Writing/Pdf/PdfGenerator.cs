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
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;

public sealed class PdfGenerator
{
    public IReportFontResolver? DefaultFontResolver { get; set; } // 既定のフォントリゾルバ

    private static int fontPlatformConfigured;

    private static readonly string[] FallbackFontNames =
    [
        "MS PGothic",
        "ＭＳ Ｐゴシック",
        "MS Gothic",
        "ＭＳ ゴシック",
        "MS UI Gothic",
        "Yu Gothic",
        "Yu Gothic UI",
        "Meiryo",
        "Meiryo UI",
        "Yu Mincho",
        "MS PMincho",
        "ＭＳ Ｐ明朝",
        "HGPMinchoE",
        "HGP明朝E",
        "Arial",
        "Helvetica",
        "Segoe UI",
        "Liberation Sans",
        "DejaVu Sans",
        "Noto Sans",
        "Times New Roman",
        "Courier New",
    ];

    private static readonly Dictionary<string, string[]> KnownFontAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["ＭＳ Ｐゴシック"] = ["MS PGothic", "MS UI Gothic", "Yu Gothic", "Meiryo"],
            ["MS Pゴシック"] = ["MS PGothic", "MS UI Gothic", "Yu Gothic", "Meiryo"],
            ["MS PGothic"] = ["ＭＳ Ｐゴシック", "MS UI Gothic", "Yu Gothic", "Meiryo"],
            ["ＭＳ ゴシック"] = ["MS Gothic", "MS PGothic", "MS UI Gothic"],
            ["Meiryo UI"] = ["Meiryo", "Yu Gothic UI", "Yu Gothic"],
            ["HGP明朝E"] = ["HGPMinchoE", "Yu Mincho", "MS PMincho", "ＭＳ Ｐ明朝"],
        };

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
                DrawBorders(graphics, sourceSheet, pagePlan.Cells);
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

        if (fontResolverProperty?.GetValue(null) is null && fallbackFontResolverProperty?.GetValue(null) is null)
        {
            fontResolverProperty?.SetValue(
                null,
                new WindowsInstalledFontResolver("Yu Gothic UI", "Meiryo UI", "Yu Gothic", "Meiryo", "MS UI Gothic", "Segoe UI"));
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
            if (!IsTransparentColor(sourceCell.Style.Fill.BackgroundColorHex))
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
            var textRect = new XRect(
                renderCell.ContentBounds.X,
                renderCell.ContentBounds.Y,
                Math.Max(0, renderCell.ContentBounds.Width),
                Math.Max(0, renderCell.ContentBounds.Height));

            if (sourceCell.Style.WrapText || sourceCell.DisplayText.Contains('\n', StringComparison.Ordinal))
            {
                var formatter = new XTextFormatter(graphics)
                {
                    Alignment = ResolveParagraphAlignment(sourceCell),
                };

                formatter.DrawString(
                    sourceCell.DisplayText,
                    font,
                    textBrush,
                    textRect,
                    ResolveStringFormat(sourceCell));
                continue;
            }

            graphics.DrawString(
                sourceCell.DisplayText,
                font,
                textBrush,
                textRect,
                ResolveStringFormat(sourceCell));
        }
    }

    private static void DrawBorders(XGraphics graphics, ReportSheet sourceSheet, IReadOnlyList<PdfCellRenderInfo> cells)
    {
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(cell => cell.Address, StringComparer.Ordinal);
        var drawnLines = new HashSet<string>(StringComparer.Ordinal);

        foreach (var renderCell in cells)
        {
            if (!sourceCellsByAddress.TryGetValue(renderCell.CellAddress, out var sourceCell))
            {
                continue;
            }

            if (sourceCell.Merge is not null && !renderCell.IsMergedOwner)
            {
                continue;
            }

            var borders = sourceCell.Merge is null
                ? sourceCell.Style.Borders
                : ResolveMergedBorders(sourceSheet, sourceCell);
            var cellBounds = renderCell.OuterBounds;
            DrawBorderSide(
                graphics,
                borders.Top,
                new ReportLine
                {
                    X1 = cellBounds.X,
                    Y1 = cellBounds.Y,
                    X2 = cellBounds.Right,
                    Y2 = cellBounds.Y,
                },
                drawnLines);
            DrawBorderSide(
                graphics,
                borders.Right,
                new ReportLine
                {
                    X1 = cellBounds.Right,
                    Y1 = cellBounds.Y,
                    X2 = cellBounds.Right,
                    Y2 = cellBounds.Bottom,
                },
                drawnLines);
            DrawBorderSide(
                graphics,
                borders.Bottom,
                new ReportLine
                {
                    X1 = cellBounds.Right,
                    Y1 = cellBounds.Bottom,
                    X2 = cellBounds.X,
                    Y2 = cellBounds.Bottom,
                },
                drawnLines);
            DrawBorderSide(
                graphics,
                borders.Left,
                new ReportLine
                {
                    X1 = cellBounds.X,
                    Y1 = cellBounds.Bottom,
                    X2 = cellBounds.X,
                    Y2 = cellBounds.Y,
                },
                drawnLines);
        }
    }

    private static void DrawImages(XGraphics graphics, IReadOnlyList<PdfImageRenderInfo> images)
    {
        foreach (var image in images)
        {
            if (image.ImageBytes.IsEmpty)
            {
                continue;
            }

            try
            {
                using var stream = new MemoryStream(image.ImageBytes.ToArray());
                var xImage = XImage.FromStream(stream);
                graphics.DrawImage(xImage, image.Bounds.X, image.Bounds.Y, image.Bounds.Width, image.Bounds.Height);
            }
            catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException)
            {
            }
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

        if (!string.IsNullOrWhiteSpace(font.Name) &&
            KnownFontAliases.TryGetValue(font.Name, out var aliases))
        {
            foreach (var alias in aliases)
            {
                if (seen.Add(alias))
                {
                    yield return alias;
                }
            }
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
            font = new XFont(fontName, size, style, new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.EmbedCompleteFontFile));
            return true;
        }
        catch (InvalidOperationException)
        {
        }
        catch (ArgumentException)
        {
        }
        catch (NullReferenceException)
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

    private static bool IsTransparentColor(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);

    private static double ResolveBorderWidth(ReportBorderStyle style) =>
        style switch
        {
            ReportBorderStyle.Thick => 2.25d,
            ReportBorderStyle.Medium => 1.5d,
            ReportBorderStyle.DoubleLine => 0.75d,
            ReportBorderStyle.Hair => 0.25d,
            _ => 0.75d,
        };

    private static void ApplyBorderStyle(XPen pen, ReportBorderStyle style)
    {
        switch (style)
        {
            case ReportBorderStyle.Dashed:
                pen.DashStyle = XDashStyle.Dash;
                break;
            case ReportBorderStyle.Dotted:
                pen.DashStyle = XDashStyle.Dot;
                break;
            case ReportBorderStyle.DashDot:
                pen.DashStyle = XDashStyle.DashDot;
                break;
        }
    }

    private static void DrawDoubleBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        var gap = Math.Max(1.5d, width * 1.5d);
        if (Math.Abs(line.Y1 - line.Y2) < 0.01d)
        {
            DrawSolidBorder(
                graphics,
                color,
                width,
                line with
                {
                    Y1 = line.Y1 - (gap / 2d),
                    Y2 = line.Y2 - (gap / 2d),
                });
            DrawSolidBorder(
                graphics,
                color,
                width,
                line with
                {
                    Y1 = line.Y1 + (gap / 2d),
                    Y2 = line.Y2 + (gap / 2d),
                });
            return;
        }

        DrawSolidBorder(
            graphics,
            color,
            width,
            line with
            {
                X1 = line.X1 - (gap / 2d),
                X2 = line.X2 - (gap / 2d),
            });
        DrawSolidBorder(
            graphics,
            color,
            width,
            line with
            {
                X1 = line.X1 + (gap / 2d),
                X2 = line.X2 + (gap / 2d),
            });
    }

    private static XStringFormat ResolveStringFormat(ReportCell cell)
    {
        var horizontalAlignment = ResolveHorizontalAlignment(cell);
        var verticalAlignment = cell.Style.Alignment.Vertical;
        return new XStringFormat
        {
            Alignment = horizontalAlignment switch
            {
                ReportHorizontalAlignment.Center => XStringAlignment.Center,
                ReportHorizontalAlignment.Right => XStringAlignment.Far,
                _ => XStringAlignment.Near,
            },
            LineAlignment = verticalAlignment switch
            {
                ReportVerticalAlignment.Center => XLineAlignment.Center,
                ReportVerticalAlignment.Bottom => XLineAlignment.Far,
                _ => XLineAlignment.Near,
            },
        };
    }

    private static XParagraphAlignment ResolveParagraphAlignment(ReportCell cell) =>
        ResolveHorizontalAlignment(cell) switch
        {
            ReportHorizontalAlignment.Center => XParagraphAlignment.Center,
            ReportHorizontalAlignment.Right => XParagraphAlignment.Right,
            ReportHorizontalAlignment.Justify => XParagraphAlignment.Justify,
            _ => XParagraphAlignment.Left,
        };

    private static ReportHorizontalAlignment ResolveHorizontalAlignment(ReportCell cell)
    {
        if (cell.Style.Alignment.Horizontal != ReportHorizontalAlignment.General)
        {
            return cell.Style.Alignment.Horizontal;
        }

        return cell.Value.Kind switch
        {
            ReportCellValueKind.Number => ReportHorizontalAlignment.Right,
            ReportCellValueKind.DateTime => ReportHorizontalAlignment.Right,
            _ => ReportHorizontalAlignment.Left,
        };
    }

    private static void DrawBorderSide(
        XGraphics graphics,
        ReportBorder border,
        ReportLine line,
        HashSet<string> drawnLines)
    {
        if (border.Style == ReportBorderStyle.None)
        {
            return;
        }

        var lineKey = BuildLineKey(line);
        if (!drawnLines.Add(lineKey))
        {
            return;
        }

        var borderColor = ToColor(border.ColorHex);
        var borderWidth = border.Width > 0d ? border.Width : ResolveBorderWidth(border.Style);
        var pen = new XPen(borderColor, borderWidth);
        ApplyBorderStyle(pen, border.Style);
        if (border.Style == ReportBorderStyle.DoubleLine)
        {
            DrawDoubleBorder(graphics, borderColor, borderWidth, line);
            return;
        }

        if (IsSolidBorder(border.Style))
        {
            DrawSolidBorder(graphics, borderColor, borderWidth, line);
            return;
        }

        graphics.DrawLine(pen, line.X1, line.Y1, line.X2, line.Y2);
    }

    private static void DrawSolidBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        var brush = new XSolidBrush(color);
        if (Math.Abs(line.Y1 - line.Y2) < 0.01d)
        {
            var left = Math.Min(line.X1, line.X2);
            var top = line.Y1 - (width / 2d);
            graphics.DrawRectangle(brush, left, top, Math.Abs(line.X2 - line.X1), width);
            return;
        }

        var leftEdge = line.X1 - (width / 2d);
        var topEdge = Math.Min(line.Y1, line.Y2);
        graphics.DrawRectangle(brush, leftEdge, topEdge, width, Math.Abs(line.Y2 - line.Y1));
    }

    private static bool IsSolidBorder(ReportBorderStyle style) =>
        style is ReportBorderStyle.Hair or ReportBorderStyle.Thin or ReportBorderStyle.Medium or ReportBorderStyle.Thick;

    private static ReportBorders ResolveMergedBorders(ReportSheet sourceSheet, ReportCell ownerCell)
    {
        ArgumentNullException.ThrowIfNull(sourceSheet);
        ArgumentNullException.ThrowIfNull(ownerCell);

        var mergeInfo = ownerCell.Merge ?? throw new InvalidOperationException("Merged border resolution requires merge info.");
        var mergedCells = sourceSheet.Cells
            .Where(cell => mergeInfo.Range.Contains(cell.Row, cell.Column))
            .ToList();

        var mergedBorders = new ReportBorders
        {
            Top = ResolveMergedBorder(
                mergedCells,
                mergeInfo.Range.StartRow,
                mergeInfo.Range.StartColumn,
                mergeInfo.Range.EndColumn,
                static cell => cell.Row,
                static cell => cell.Column,
                static cell => cell.Style.Borders.Top),
            Right = ResolveMergedBorder(
                mergedCells,
                mergeInfo.Range.EndColumn,
                mergeInfo.Range.StartRow,
                mergeInfo.Range.EndRow,
                static cell => cell.Column,
                static cell => cell.Row,
                static cell => cell.Style.Borders.Right),
            Bottom = ResolveMergedBorder(
                mergedCells,
                mergeInfo.Range.EndRow,
                mergeInfo.Range.StartColumn,
                mergeInfo.Range.EndColumn,
                static cell => cell.Row,
                static cell => cell.Column,
                static cell => cell.Style.Borders.Bottom),
            Left = ResolveMergedBorder(
                mergedCells,
                mergeInfo.Range.StartColumn,
                mergeInfo.Range.StartRow,
                mergeInfo.Range.EndRow,
                static cell => cell.Column,
                static cell => cell.Row,
                static cell => cell.Style.Borders.Left),
        };

        return mergedBorders;
    }

    private static ReportBorder ResolveMergedBorder(
        IReadOnlyList<ReportCell> mergedCells,
        int fixedIndex,
        int rangeStart,
        int rangeEnd,
        Func<ReportCell, int> fixedSelector,
        Func<ReportCell, int> rangeSelector,
        Func<ReportCell, ReportBorder> borderSelector)
    {
        ReportBorder? bestBorder = null;
        foreach (var cell in mergedCells.Where(cell => fixedSelector(cell) == fixedIndex && rangeSelector(cell) >= rangeStart && rangeSelector(cell) <= rangeEnd))
        {
            var border = borderSelector(cell);
            if (bestBorder is null || GetBorderPriority(border.Style) > GetBorderPriority(bestBorder.Style))
            {
                bestBorder = border;
            }
        }

        return bestBorder ?? new ReportBorder();
    }

    private static int GetBorderPriority(ReportBorderStyle style) =>
        style switch
        {
            ReportBorderStyle.DoubleLine => 6,
            ReportBorderStyle.Thick => 5,
            ReportBorderStyle.Medium => 4,
            ReportBorderStyle.Thin => 3,
            ReportBorderStyle.DashDot => 2,
            ReportBorderStyle.Dashed => 2,
            ReportBorderStyle.Dotted => 2,
            ReportBorderStyle.Hair => 1,
            _ => 0,
        };

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);
    }

    private static string BuildLineKey(ReportLine line)
    {
        return string.Create(
            CultureInfo.InvariantCulture,
            $"{Math.Round(Math.Min(line.X1, line.X2), 4)}:{Math.Round(Math.Min(line.Y1, line.Y2), 4)}:{Math.Round(Math.Max(line.X1, line.X2), 4)}:{Math.Round(Math.Max(line.Y1, line.Y2), 4)}");
    }
}
