namespace OysterReport.Generator;

using System.Globalization;
using System.Reflection;
using System.Text;

using ClosedXML.Excel;

using OysterReport.Helpers;

using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;

internal sealed class PdfGenerator
{
    public IReportFontResolver? DefaultFontResolver { get; set; }

    private static int fontPlatformConfigured;

    private static readonly string[] HeaderFooterFallbackFontNames =
    [
        "Arial",
        "Segoe UI",
        "Helvetica",
        "Liberation Sans",
        "DejaVu Sans"
    ];

    public void Generate(
        ReportWorkbook workbook,
        Stream output,
        PdfGeneratorOption? option = null)
    {
        EnsurePdfSharpFontConfiguration();

        var effectiveOptions = option ?? new PdfGeneratorOption();
        effectiveOptions.FontResolver ??= DefaultFontResolver;

        var renderPlan = BuildRenderPlan(workbook);
        WritePdf(workbook, renderPlan, output, effectiveOptions);
    }

    internal static IReadOnlyList<PdfRenderSheetPlan> BuildRenderPlan(ReportWorkbook workbook)
    {
        return PdfRenderPlanner.BuildPlan(workbook);
    }

    internal static void WritePdf(
        ReportWorkbook workbook,
        IReadOnlyList<PdfRenderSheetPlan> sheetPlans,
        Stream output,
        PdfGeneratorOption option)
    {
        using var document = new PdfDocument();
        document.Options.CompressContentStreams = option.CompressContentStreams;

        if (option.EmbedDocumentMetadata)
        {
            document.Info.Title = workbook.Metadata.TemplateName;
        }

        for (var sheetIndex = 0; sheetIndex < sheetPlans.Count; sheetIndex++)
        {
            var sheetPlan = sheetPlans[sheetIndex];
            var sourceSheet = workbook.Sheets[sheetIndex];
            foreach (var pagePlan in sheetPlan.Pages)
            {
                var page = document.AddPage();
                page.Width = XUnit.FromPoint(pagePlan.PageBounds.Width);
                page.Height = XUnit.FromPoint(pagePlan.PageBounds.Height);
                using var graphics = XGraphics.FromPdfPage(page);
                DrawPageBackground(graphics, pagePlan.PageBounds);
                DrawCells(graphics, sourceSheet, pagePlan.Cells, option);
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
            fontResolverProperty?.SetValue(null, new WindowsInstalledFontResolver());
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
        PdfGeneratorOption option)
    {
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(cell => cell.Address, StringComparer.Ordinal);

        var backgroundGroups = new Dictionary<string, List<ReportRect>>(StringComparer.Ordinal);
        foreach (var renderCell in cells)
        {
            if (!sourceCellsByAddress.TryGetValue(renderCell.CellAddress, out var cell))
            {
                continue;
            }

            var colorHex = cell.Style.Fill.BackgroundColorHex;
            if (IsTransparentColor(colorHex))
            {
                continue;
            }

            if (!backgroundGroups.TryGetValue(colorHex, out var rects))
            {
                rects = [];
                backgroundGroups[colorHex] = rects;
            }

            rects.Add(renderCell.OuterBounds);
        }

        foreach (var (colorHex, rects) in backgroundGroups)
        {
            var brush = new XSolidBrush(ToColor(colorHex));
            var path = new XGraphicsPath();
            foreach (var rect in rects)
            {
                path.AddRectangle(new XRect(rect.X, rect.Y, rect.Width, rect.Height));
            }

            graphics.DrawPath(brush, path);
        }

        foreach (var renderCell in cells)
        {
            if (!sourceCellsByAddress.TryGetValue(renderCell.CellAddress, out var sourceCell))
            {
                continue;
            }

            if (!renderCell.IsMergedOwner && sourceCell.Merge is not null)
            {
                continue;
            }

            if (string.IsNullOrEmpty(sourceCell.DisplayText))
            {
                continue;
            }

            var font = ResolveFont(sourceCell.Style.Font, option);
            var textBrush = new XSolidBrush(ToColor(sourceCell.Style.Font.ColorHex));
            var textRect = new XRect(
                renderCell.ContentBounds.X,
                renderCell.ContentBounds.Y,
                Math.Max(0, renderCell.ContentBounds.Width),
                Math.Max(0, renderCell.ContentBounds.Height));

            var clipRect = new XRect(
                renderCell.TextBounds.X,
                renderCell.TextBounds.Y,
                Math.Max(0, renderCell.TextBounds.Width),
                Math.Max(0, renderCell.TextBounds.Height));

            var clipState = graphics.Save();
            try
            {
                graphics.IntersectClip(clipRect);

                if (sourceCell.Style.WrapText || sourceCell.DisplayText.Contains('\n', StringComparison.Ordinal))
                {
                    var formatter = new XTextFormatter(graphics)
                    {
                        Alignment = ResolveParagraphAlignment(sourceCell)
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
            finally
            {
                graphics.Restore(clipState);
            }
        }
    }

    private static void DrawBorders(XGraphics graphics, ReportSheet sourceSheet, IEnumerable<PdfCellRenderInfo> cells)
    {
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(cell => cell.Address, StringComparer.Ordinal);
        var collectedLines = new Dictionary<string, (ReportLine Line, ReportBorder Border)>(StringComparer.Ordinal);

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
            CollectBorderSide(borders.Top, new ReportLine { X1 = cellBounds.X, Y1 = cellBounds.Y, X2 = cellBounds.Right, Y2 = cellBounds.Y }, collectedLines);
            CollectBorderSide(borders.Right, new ReportLine { X1 = cellBounds.Right, Y1 = cellBounds.Y, X2 = cellBounds.Right, Y2 = cellBounds.Bottom }, collectedLines);
            CollectBorderSide(borders.Bottom, new ReportLine { X1 = cellBounds.Right, Y1 = cellBounds.Bottom, X2 = cellBounds.X, Y2 = cellBounds.Bottom }, collectedLines);
            CollectBorderSide(borders.Left, new ReportLine { X1 = cellBounds.X, Y1 = cellBounds.Bottom, X2 = cellBounds.X, Y2 = cellBounds.Y }, collectedLines);
        }

        foreach (var (line, border) in collectedLines.Values.OrderBy(entry => GetBorderPriority(entry.Border.Style)))
        {
            DrawBorderLine(graphics, border, line);
        }
    }

    private static void CollectBorderSide(
        ReportBorder border,
        ReportLine line,
        Dictionary<string, (ReportLine Line, ReportBorder Border)> collectedLines)
    {
        if (border.Style == XLBorderStyleValues.None)
        {
            return;
        }

        var lineKey = BuildLineKey(line);
        if (collectedLines.TryGetValue(lineKey, out var existing) &&
            GetBorderPriority(existing.Border.Style) >= GetBorderPriority(border.Style))
        {
            return;
        }

        collectedLines[lineKey] = (line, border);
    }

    private static void DrawBorderLine(XGraphics graphics, ReportBorder border, ReportLine line)
    {
        var borderColor = ToColor(border.ColorHex);
        var borderWidth = border.Width > 0d ? border.Width : ResolveBorderWidth(border.Style);
        var pen = new XPen(borderColor, borderWidth);
        ApplyBorderStyle(pen, border.Style);
        if (border.Style == XLBorderStyleValues.Double)
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

    private static void DrawImages(XGraphics graphics, IEnumerable<PdfImageRenderInfo> images)
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
        var font = CreateFallbackFont(PdfRenderingConstants.HeaderFooterFontSizePoints);

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
            graphics.DrawString(sections.Left, font, XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopLeft);
        }

        if (!string.IsNullOrWhiteSpace(sections.Center))
        {
            graphics.DrawString(sections.Center, font, XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopCenter);
        }

        if (!string.IsNullOrWhiteSpace(sections.Right))
        {
            graphics.DrawString(sections.Right, font, XBrushes.Black,
                new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopRight);
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
                case 'L': current = left; break;
                case 'C': current = center; break;
                case 'R': current = right; break;
                case 'P': current.Append(pageNumber.ToString(CultureInfo.InvariantCulture)); break;
                case 'N': current.Append(totalPages.ToString(CultureInfo.InvariantCulture)); break;
                case '&': current.Append('&'); break;
            }
        }

        return new HeaderFooterSections(
            left.ToString().Trim(),
            center.ToString().Trim(),
            right.ToString().Trim());
    }

    private static XFont ResolveFont(ReportFont font, PdfGeneratorOption option)
    {
        var fontSize = font.Size <= 0 ? PdfRenderingConstants.DefaultCellFontSizePoints : font.Size;
        var style = XFontStyleEx.Regular;
        if (font.Bold) style |= XFontStyleEx.Bold;
        if (font.Italic) style |= XFontStyleEx.Italic;

        foreach (var fontName in EnumerateCandidateFontNames(font, option))
        {
            if (TryCreateFont(fontName, fontSize, style, out var resolvedFont))
            {
                return resolvedFont;
            }
        }

        throw new InvalidOperationException($"No appropriate font found for family name '{font.Name}' and known fallbacks.");
    }

    private static IEnumerable<string> EnumerateCandidateFontNames(ReportFont font, PdfGeneratorOption option)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (option.FontResolver is not null)
        {
            var resolution = option.FontResolver.Resolve(new ReportFontRequest
            {
                FontName = font.Name,
                Bold = font.Bold,
                Italic = font.Italic
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
    }

    private static XFont CreateFallbackFont(double size)
    {
        foreach (var fontName in HeaderFooterFallbackFontNames)
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
            font = new XFont(fontName, size, style, new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.TryComputeSubset));
            return true;
        }
        catch (InvalidOperationException) { }
        catch (ArgumentException) { }
        catch (NullReferenceException) { }

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

    private static double ResolveBorderWidth(XLBorderStyleValues style) =>
        PdfRenderingConstants.ResolveBorderWidth(style);

    private static void ApplyBorderStyle(XPen pen, XLBorderStyleValues style)
    {
        pen.DashStyle = style switch
        {
            XLBorderStyleValues.Dashed => XDashStyle.Dash,
            XLBorderStyleValues.Dotted => XDashStyle.Dot,
            XLBorderStyleValues.DashDot => XDashStyle.DashDot,
            _ => pen.DashStyle
        };
    }

    private static void DrawDoubleBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        var gap = Math.Max(PdfRenderingConstants.MinimumDoubleBorderGapPoints, width * PdfRenderingConstants.DoubleBorderGapWidthMultiplier);
        if (Math.Abs(line.Y1 - line.Y2) < PdfRenderingConstants.StraightLineTolerancePoints)
        {
            DrawSolidBorder(graphics, color, width, line with { Y1 = line.Y1 - (gap / 2d), Y2 = line.Y2 - (gap / 2d) });
            DrawSolidBorder(graphics, color, width, line with { Y1 = line.Y1 + (gap / 2d), Y2 = line.Y2 + (gap / 2d) });
            return;
        }

        DrawSolidBorder(graphics, color, width, line with { X1 = line.X1 - (gap / 2d), X2 = line.X2 - (gap / 2d) });
        DrawSolidBorder(graphics, color, width, line with { X1 = line.X1 + (gap / 2d), X2 = line.X2 + (gap / 2d) });
    }

    private static XStringFormat ResolveStringFormat(ReportCell cell)
    {
        var horizontalAlignment = ResolveHorizontalAlignment(cell);
        var verticalAlignment = cell.Style.Alignment.Vertical;
        return new XStringFormat
        {
            Alignment = horizontalAlignment switch
            {
                XLAlignmentHorizontalValues.Center => XStringAlignment.Center,
                XLAlignmentHorizontalValues.Right => XStringAlignment.Far,
                _ => XStringAlignment.Near
            },
            LineAlignment = verticalAlignment switch
            {
                XLAlignmentVerticalValues.Center => XLineAlignment.Center,
                XLAlignmentVerticalValues.Bottom => XLineAlignment.Far,
                _ => XLineAlignment.Near
            }
        };
    }

    private static XParagraphAlignment ResolveParagraphAlignment(ReportCell cell) =>
        ResolveHorizontalAlignment(cell) switch
        {
            XLAlignmentHorizontalValues.Center => XParagraphAlignment.Center,
            XLAlignmentHorizontalValues.Right => XParagraphAlignment.Right,
            XLAlignmentHorizontalValues.Justify => XParagraphAlignment.Justify,
            _ => XParagraphAlignment.Left
        };

    private static XLAlignmentHorizontalValues ResolveHorizontalAlignment(ReportCell cell)
    {
        if (cell.Style.Alignment.Horizontal != XLAlignmentHorizontalValues.General)
        {
            return cell.Style.Alignment.Horizontal;
        }

        return cell.Value.Kind switch
        {
            XLDataType.Number => XLAlignmentHorizontalValues.Right,
            XLDataType.DateTime => XLAlignmentHorizontalValues.Right,
            _ => XLAlignmentHorizontalValues.Left
        };
    }

    private static void DrawSolidBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        var brush = new XSolidBrush(color);
        if (Math.Abs(line.Y1 - line.Y2) < PdfRenderingConstants.StraightLineTolerancePoints)
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

    private static bool IsSolidBorder(XLBorderStyleValues style) =>
        style is XLBorderStyleValues.Hair or XLBorderStyleValues.Thin or XLBorderStyleValues.Medium or XLBorderStyleValues.Thick;

    private static ReportBorders ResolveMergedBorders(ReportSheet sourceSheet, ReportCell ownerCell)
    {
        var mergeInfo = ownerCell.Merge ?? throw new InvalidOperationException("Merged border resolution requires merge info.");
        var mergedCells = sourceSheet.Cells
            .Where(cell => mergeInfo.Range.Contains(cell.Row, cell.Column))
            .ToList();

        return new ReportBorders
        {
            Top = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn,
                static cell => cell.Row, static cell => cell.Column, static cell => cell.Style.Borders.Top),
            Right = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow,
                static cell => cell.Column, static cell => cell.Row, static cell => cell.Style.Borders.Right),
            Bottom = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn,
                static cell => cell.Row, static cell => cell.Column, static cell => cell.Style.Borders.Bottom),
            Left = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow,
                static cell => cell.Column, static cell => cell.Row, static cell => cell.Style.Borders.Left)
        };
    }

    private static ReportBorder ResolveMergedBorder(
        IEnumerable<ReportCell> mergedCells,
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

    private static int GetBorderPriority(XLBorderStyleValues style) =>
        style switch
        {
            XLBorderStyleValues.Double => 6,
            XLBorderStyleValues.Thick => 5,
            XLBorderStyleValues.Medium => 4,
            XLBorderStyleValues.Thin => 3,
            XLBorderStyleValues.DashDot => 2,
            XLBorderStyleValues.Dashed => 2,
            XLBorderStyleValues.Dotted => 2,
            XLBorderStyleValues.Hair => 1,
            _ => 0
        };

    private static string BuildLineKey(ReportLine line)
    {
        return string.Create(
            CultureInfo.InvariantCulture,
            $"{Math.Round(Math.Min(line.X1, line.X2), 4)}:{Math.Round(Math.Min(line.Y1, line.Y2), 4)}:{Math.Round(Math.Max(line.X1, line.X2), 4)}:{Math.Round(Math.Max(line.Y1, line.Y2), 4)}");
    }

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);
    }
}
