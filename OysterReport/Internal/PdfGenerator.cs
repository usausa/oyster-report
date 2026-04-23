namespace OysterReport.Internal;

using System.Collections.Concurrent;
using System.Globalization;
using System.Text;

using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

internal static class PdfGenerator
{
    private const double BoldSimulationOffset = 0.35d;

    private static int fontPlatformConfigured;

    private static readonly ConcurrentDictionary<string, XColor> ColorCache = new(StringComparer.Ordinal);

    private sealed record ResolvedFontRenderInfo
    {
        public XFont Font { get; init; } = null!;

        public bool SimulateBold { get; init; }
    }

    //--------------------------------------------------------------------------------
    // Write
    //--------------------------------------------------------------------------------

    internal static void WritePdf(ReportRenderContext context, Stream output)
    {
        EnsurePdfSharpFontConfiguration();

        using var document = new PdfDocument();
        document.Options.CompressContentStreams = context.CompressContentStreams;

        if (context.EmbedDocumentMetadata)
        {
            document.Info.Title = context.Workbook.Metadata.TemplateName;
        }

        for (var sheetIndex = 0; sheetIndex < context.SheetPlans.Count; sheetIndex++)
        {
            var sheetPlan = context.SheetPlans[sheetIndex];
            var sourceSheet = context.Workbook.Sheets[sheetIndex];
            foreach (var pagePlan in sheetPlan.Pages)
            {
                var page = document.AddPage();
                page.Width = XUnit.FromPoint(pagePlan.PageBounds.Width);
                page.Height = XUnit.FromPoint(pagePlan.PageBounds.Height);
                using var graphics = XGraphics.FromPdfPage(page);
                DrawPageBackground(graphics, pagePlan.PageBounds);
                DrawCells(graphics, sourceSheet, pagePlan.Cells, context);
                DrawBorders(graphics, sourceSheet, pagePlan.Cells, context.RenderingOptions);
                DrawImages(graphics, sheetPlan.Images);
                DrawHeaderFooter(graphics, pagePlan.HeaderFooter, pagePlan.PageNumber, sheetPlan.Pages.Count, context.RenderingOptions);
            }
        }

        document.Save(output, closeStream: false);
    }

    //--------------------------------------------------------------------------------
    // Setup
    //--------------------------------------------------------------------------------

    private static void EnsurePdfSharpFontConfiguration()
    {
        if (Interlocked.Exchange(ref fontPlatformConfigured, 1) == 1)
        {
            return;
        }

        if (GlobalFontSettings.FontResolver is null && GlobalFontSettings.FallbackFontResolver is null)
        {
            GlobalFontSettings.FontResolver = new ReportFontResolverAdapter();
        }
    }

    //--------------------------------------------------------------------------------
    // Page background
    //--------------------------------------------------------------------------------

    private static void DrawPageBackground(XGraphics graphics, ReportRect pageBounds)
    {
        // Fills the entire page with a white background rectangle
        graphics.DrawRectangle(XBrushes.White, pageBounds.X, pageBounds.Y, pageBounds.Width, pageBounds.Height);
    }

    //--------------------------------------------------------------------------------
    // Cell
    //--------------------------------------------------------------------------------

    private static void DrawCells(
        XGraphics graphics,
        ReportSheet sourceSheet,
        IReadOnlyList<PdfCellRenderInfo> cells,
        ReportRenderContext context)
    {
        // Draws cell backgrounds and text
        // Backgrounds are batched by color; text is placed according to wrap, alignment, and other settings
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(static x => x.Address, StringComparer.Ordinal);

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

            if (String.IsNullOrEmpty(sourceCell.DisplayText))
            {
                continue;
            }

            var resolvedFont = ResolveFont(sourceCell.Style.Font, context);
            var fontColor = ToColor(sourceCell.Style.Font.ColorHex);
            var textBrush = new XSolidBrush(fontColor);
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

                DrawCellText(
                    graphics,
                    sourceCell.DisplayText,
                    resolvedFont,
                    textBrush,
                    textRect,
                    sourceCell);

                DrawTextDecorations(
                    graphics,
                    resolvedFont.Font,
                    fontColor,
                    textRect,
                    sourceCell.DisplayText,
                    sourceCell,
                    context.RenderingOptions);
            }
            finally
            {
                graphics.Restore(clipState);
            }
        }
    }

    //--------------------------------------------------------------------------------
    // Border
    //--------------------------------------------------------------------------------

    private static void DrawBorders(XGraphics graphics, ReportSheet sourceSheet, IEnumerable<PdfCellRenderInfo> cells, ReportRenderOption renderOption)
    {
        // Draws cell borders
        // Duplicate edges adopt the higher-priority border style and are drawn only once
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(static x => x.Address, StringComparer.Ordinal);
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

        foreach (var (line, border) in collectedLines.Values.OrderBy(static x => GetBorderPriority(x.Border.Style)))
        {
            DrawBorderLine(graphics, border, line, renderOption);
        }
    }

    //--------------------------------------------------------------------------------
    // Border helpers
    //--------------------------------------------------------------------------------

    private static void CollectBorderSide(
        ReportBorder border,
        ReportLine line,
        Dictionary<string, (ReportLine Line, ReportBorder Border)> collectedLines)
    {
        // Collects border info for one edge. When multiple styles exist on the same edge, keeps the higher-priority one
        if (border.Style == BorderLineStyle.None)
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

    private static void DrawBorderLine(XGraphics graphics, ReportBorder border, ReportLine line, ReportRenderOption renderOption)
    {
        // Draws a solid, dashed, or double border-line according to the border style.
        var borderColor = ToColor(border.ColorHex);
        var borderWidth = ResolveBorderWidth(border.Style, renderOption);
        var pen = new XPen(borderColor, borderWidth);
        ApplyBorderStyle(pen, border.Style);
        if (border.Style == BorderLineStyle.Double)
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

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    private static void DrawImages(XGraphics graphics, IEnumerable<PdfImageRenderInfo> images)
    {
        // Draws images embedded in the sheet onto the page. Skips images that fail to decode
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

    //--------------------------------------------------------------------------------
    // Header / Footer
    //--------------------------------------------------------------------------------

    private static void DrawHeaderFooter(
        XGraphics graphics,
        PdfHeaderFooterRenderInfo headerFooter,
        int pageNumber,
        int totalPages,
        ReportRenderOption renderOption)
    {
        // Draws header/footer text in the left, center, and right areas
        var headerSections = ResolveHeaderFooterSections(headerFooter.HeaderText, pageNumber, totalPages);
        var footerSections = ResolveHeaderFooterSections(headerFooter.FooterText, pageNumber, totalPages);
        var font = CreateFallbackFont(renderOption.HeaderFooterFontSize, renderOption.HeaderFooterFallbackFonts);

        DrawHeaderFooterSections(graphics, headerSections, headerFooter.HeaderBounds, font);
        DrawHeaderFooterSections(graphics, footerSections, headerFooter.FooterBounds, font);
    }

    private static void DrawHeaderFooterSections(
        XGraphics graphics,
        HeaderFooterSections sections,
        ReportRect bounds,
        XFont font)
    {
        // Draws each section of a header/footer string that has been split into left, center, and right parts
        if (!String.IsNullOrWhiteSpace(sections.Left))
        {
            graphics.DrawString(sections.Left, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopLeft);
        }

        if (!String.IsNullOrWhiteSpace(sections.Center))
        {
            graphics.DrawString(sections.Center, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopCenter);
        }

        if (!String.IsNullOrWhiteSpace(sections.Right))
        {
            graphics.DrawString(sections.Right, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopRight);
        }
    }

    private static HeaderFooterSections ResolveHeaderFooterSections(string? text, int pageNumber, int totalPages)
    {
        // Parses the Excel header/footer format string (&L, &C, &R, &P, &N) and splits it into left, center, right, and page number parts
        if (String.IsNullOrWhiteSpace(text))
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
            switch (Char.ToUpperInvariant(text[index]))
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

    //--------------------------------------------------------------------------------
    // Font
    //--------------------------------------------------------------------------------

    private static ResolvedFontRenderInfo ResolveFont(ReportFont font, ReportRenderContext context)
    {
        // Generates an XFont for PDFSharp from ReportFont attributes and the font resolver.
        var fontSize = font.Size <= 0 ? context.RenderingOptions.DefaultCellFontSize : font.Size;
        var nameToUse = font.Name;
        var simulateBold = false;

        var resolvedTypeface = context.FontResolver?.ResolveTypeface(font.Name, font.Bold, font.Italic);
        if (resolvedTypeface != null && !String.IsNullOrWhiteSpace(resolvedTypeface.FaceName))
        {
            ReportFontResolverAdapter.RegisterResolvedTypeface(resolvedTypeface);
            var embeddedFontData = context.FontResolver?.GetFont(resolvedTypeface.FaceName);
            if (embeddedFontData is { } fontData)
            {
                // Pre-registers the embedded font with the adapter.
                ReportFontResolverAdapter.RegisterEmbeddedFont(resolvedTypeface.FaceName, fontData);

                // When a single embedded font resource is returned, Bold is simulated at draw time
                // Italic is delegated to PDFSharp's official simulation by ReportFontResolverAdapter
                simulateBold = resolvedTypeface.MustSimulateBold;
            }

            nameToUse = resolvedTypeface.FaceName;
        }

        if (!simulateBold && font.Bold && ReportFontResolverAdapter.IsBoldSimulationRequired(nameToUse, font.Italic))
        {
            simulateBold = true;
        }

        var style = BuildActualFontStyle(font, simulateBold);

        if (!String.IsNullOrWhiteSpace(nameToUse) && TryCreateFont(nameToUse, fontSize, style, out var resolvedFont))
        {
            return new ResolvedFontRenderInfo
            {
                Font = resolvedFont,
                SimulateBold = simulateBold
            };
        }

        // Falls back to the original Excel font name if the resolver-returned name fails
        if (!String.Equals(nameToUse, font.Name, StringComparison.OrdinalIgnoreCase) &&
            !String.IsNullOrWhiteSpace(font.Name) &&
            TryCreateFont(font.Name, fontSize, BuildActualFontStyle(font, simulateBold: false), out var fallbackFont))
        {
            return new ResolvedFontRenderInfo
            {
                Font = fallbackFont
            };
        }

        throw new InvalidOperationException($"No appropriate font found. name=[{font.Name}]");
    }

    private static XFontStyleEx BuildActualFontStyle(ReportFont font, bool simulateBold)
    {
        var style = XFontStyleEx.Regular;
        if (font.Bold && !simulateBold)
        {
            style |= XFontStyleEx.Bold;
        }

        if (font.Italic)
        {
            style |= XFontStyleEx.Italic;
        }

        return style;
    }

    private static void DrawCellText(
        XGraphics graphics,
        string text,
        ResolvedFontRenderInfo resolvedFont,
        XBrush brush,
        XRect textRect,
        ReportCell sourceCell)
    {
        var drawCount = resolvedFont.SimulateBold ? 2 : 1;
        for (var pass = 0; pass < drawCount; pass++)
        {
            var passRect = pass == 0
                ? textRect
                : new XRect(textRect.X + BoldSimulationOffset, textRect.Y, textRect.Width, textRect.Height);

            var state = graphics.Save();
            try
            {
                if (sourceCell.Style.WrapText || text.Contains('\n', StringComparison.Ordinal))
                {
                    var formatter = new XTextFormatter(graphics)
                    {
                        Alignment = ResolveParagraphAlignment(sourceCell)
                    };

                    formatter.DrawString(
                        text,
                        resolvedFont.Font,
                        brush,
                        passRect,
                        XStringFormats.TopLeft);
                    continue;
                }

                graphics.DrawString(
                    text,
                    resolvedFont.Font,
                    brush,
                    passRect,
                    ResolveStringFormat(sourceCell));
            }
            finally
            {
                graphics.Restore(state);
            }
        }
    }

    private static void DrawTextDecorations(
        XGraphics graphics,
        XFont font,
        XColor color,
        XRect textRect,
        string text,
        ReportCell sourceCell,
        ReportRenderOption renderOption)
    {
        // Draws underline and/or strikeout for cell text based on font metrics
        if ((!sourceCell.Style.Font.Underline) && (!sourceCell.Style.Font.Strikeout))
        {
            return;
        }

        if ((textRect.Width <= 0) || (textRect.Height <= 0))
        {
            return;
        }

        var isWrap = sourceCell.Style.WrapText || text.Contains('\n', StringComparison.Ordinal);

        // Compute the baseline offset (pt) from font metrics
        var metrics = font.Metrics;
        var scale = font.Size / metrics.UnitsPerEm;
        var ascentPt = metrics.Ascent * scale;

        // Determine the vertical text start position from alignment (wrap always uses top)
        double verticalOffset;
        if (isWrap)
        {
            verticalOffset = 0;
        }
        else
        {
            var lineHeight = font.GetHeight();
            verticalOffset = sourceCell.Style.Alignment.Vertical switch
            {
                VerticalAlignment.Center => Math.Max(0, (textRect.Height - lineHeight) / 2),
                VerticalAlignment.Bottom => Math.Max(0, textRect.Height - lineHeight),
                _ => 0
            };
        }

        var textTopY = textRect.Y + verticalOffset;

        // Determine decoration width and X start from horizontal alignment and measured text width
        double decorationWidth;
        double decorationX;
        if (isWrap)
        {
            // For wrapped text, span the full content width
            decorationWidth = textRect.Width;
            decorationX = textRect.X;
        }
        else
        {
            decorationWidth = graphics.MeasureString(text, font).Width;
            var horizontalAlignment = ResolveHorizontalAlignment(sourceCell);
            decorationX = horizontalAlignment switch
            {
                HorizontalAlignment.Center => textRect.X + ((textRect.Width - decorationWidth) / 2),
                HorizontalAlignment.Right => textRect.X + textRect.Width - decorationWidth,
                _ => textRect.X
            };
        }

        if (decorationWidth <= 0)
        {
            return;
        }

        if (sourceCell.Style.Font.Underline)
        {
            // UnderlinePosition is the offset from baseline in font coordinates (Y up)
            // Typically negative (below baseline); negating it moves the line downward in screen space
            var lineThickness = Math.Max(renderOption.UnderlineWidth, Math.Abs(metrics.UnderlineThickness * scale));
            var lineY = textTopY + ascentPt - (metrics.UnderlinePosition * scale);
            DrawSolidBorder(graphics, color, lineThickness, new ReportLine
            {
                X1 = decorationX,
                Y1 = lineY,
                X2 = decorationX + decorationWidth,
                Y2 = lineY
            });
        }

        if (sourceCell.Style.Font.Strikeout)
        {
            // StrikethroughPosition is above the baseline in font coordinates (positive value)
            var lineThickness = Math.Max(renderOption.StrikeoutWidth, Math.Abs(metrics.StrikethroughThickness * scale));
            var lineY = textTopY + ascentPt - (metrics.StrikethroughPosition * scale);
            DrawSolidBorder(graphics, color, lineThickness, new ReportLine
            {
                X1 = decorationX,
                Y1 = lineY,
                X2 = decorationX + decorationWidth,
                Y2 = lineY
            });
        }
    }

    private static XFont CreateFallbackFont(double size, IEnumerable<string> fontNames)
    {
        // Creates a fallback font for header/footer rendering from the candidate list
        foreach (var fontName in fontNames)
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
        // Attempts to create an XFont by the given name; returns false on failure
        try
        {
            font = new XFont(fontName, size, style, new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.TryComputeSubset));
            return true;
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NullReferenceException)
        {
        }

        font = null!;
        return false;
    }

    //--------------------------------------------------------------------------------
    // Color / Style helpers
    //--------------------------------------------------------------------------------

    private static XColor ToColor(string colorHex)
    {
        // Converts an ARGB hex color string to an XColor (cached by normalized hex)
        return ColorCache.GetOrAdd(ColorHelper.NormalizeHex(colorHex), static normalized =>
        {
            var hex = normalized.AsSpan().TrimStart('#');
            if (hex.Length == 8)
            {
                return XColor.FromArgb(
                    Byte.Parse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                    Byte.Parse(hex.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                    Byte.Parse(hex.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                    Byte.Parse(hex.Slice(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
            }

            return XColors.Black;
        });
    }

    private static bool IsTransparentColor(string colorHex)
    {
        // Determines whether the color is fully transparent (alpha = 0)
        return ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);
    }

    private static double ResolveBorderWidth(BorderLineStyle style, ReportRenderOption renderOption)
    {
        // Determines the drawing width (pt) from the border style
        return style switch
        {
            BorderLineStyle.Thick => renderOption.ThickBorderWidth,
            BorderLineStyle.Medium => renderOption.MediumBorderWidth,
            BorderLineStyle.Double => renderOption.NormalBorderWidth,
            BorderLineStyle.Hair => renderOption.HairBorderWidth,
            _ => renderOption.NormalBorderWidth
        };
    }

    private static void ApplyBorderStyle(XPen pen, BorderLineStyle style)
    {
        // Applies the dash pattern corresponding to the border style to the XPen
        pen.DashStyle = style switch
        {
            BorderLineStyle.Dashed => XDashStyle.Dash,
            BorderLineStyle.Dotted => XDashStyle.Dot,
            BorderLineStyle.DashDot => XDashStyle.DashDot,
            _ => pen.DashStyle
        };
    }

    private static void DrawDoubleBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        // Draws a double-line border as two parallel solid lines
        var gap = Math.Max(RenderConstants.MinimumDoubleBorderGap, width * RenderConstants.DoubleBorderGapWidthMultiplier);
        if (Math.Abs(line.Y1 - line.Y2) < RenderConstants.StraightLineTolerance)
        {
            DrawSolidBorder(graphics, color, width, line with { Y1 = line.Y1 - (gap / 2d), Y2 = line.Y2 - (gap / 2d) });
            DrawSolidBorder(graphics, color, width, line with { Y1 = line.Y1 + (gap / 2d), Y2 = line.Y2 + (gap / 2d) });
            return;
        }

        DrawSolidBorder(graphics, color, width, line with { X1 = line.X1 - (gap / 2d), X2 = line.X2 - (gap / 2d) });
        DrawSolidBorder(graphics, color, width, line with { X1 = line.X1 + (gap / 2d), X2 = line.X2 + (gap / 2d) });
    }

    //--------------------------------------------------------------------------------
    // Cell layout
    //--------------------------------------------------------------------------------

    private static XStringFormat ResolveStringFormat(ReportCell cell)
    {
        // Generates an XStringFormat from the cell's horizontal and vertical alignment
        var horizontalAlignment = ResolveHorizontalAlignment(cell);
        var verticalAlignment = cell.Style.Alignment.Vertical;
        return new XStringFormat
        {
            Alignment = horizontalAlignment switch
            {
                HorizontalAlignment.Center => XStringAlignment.Center,
                HorizontalAlignment.Right => XStringAlignment.Far,
                _ => XStringAlignment.Near
            },
            LineAlignment = verticalAlignment switch
            {
                VerticalAlignment.Center => XLineAlignment.Center,
                VerticalAlignment.Bottom => XLineAlignment.Far,
                _ => XLineAlignment.Near
            }
        };
    }

    private static XParagraphAlignment ResolveParagraphAlignment(ReportCell cell)
    {
        // Returns the paragraph alignment enum from the cell's horizontal alignment
        return ResolveHorizontalAlignment(cell) switch
        {
            HorizontalAlignment.Center => XParagraphAlignment.Center,
            HorizontalAlignment.Right => XParagraphAlignment.Right,
            HorizontalAlignment.Justify => XParagraphAlignment.Justify,
            _ => XParagraphAlignment.Left
        };
    }

    private static HorizontalAlignment ResolveHorizontalAlignment(ReportCell cell)
    {
        // General uses the default for the value type
        if (cell.Style.Alignment.Horizontal != HorizontalAlignment.General)
        {
            return cell.Style.Alignment.Horizontal;
        }

        // Determines horizontal alignment from the cell's alignment setting
        return cell.Value.Kind switch
        {
            CellValueKind.Number => HorizontalAlignment.Right,
            CellValueKind.DateTime => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
    }

    private static void DrawSolidBorder(XGraphics graphics, XColor color, double width, ReportLine line)
    {
        // Draws a solid border-line as a filled rectangle
        var brush = new XSolidBrush(color);
        if (Math.Abs(line.Y1 - line.Y2) < RenderConstants.StraightLineTolerance)
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

    private static bool IsSolidBorder(BorderLineStyle style)
    {
        // Determines whether the border style should be drawn as a solid filled rectangle
        return style is BorderLineStyle.Hair or BorderLineStyle.Thin or BorderLineStyle.Medium or BorderLineStyle.Thick;
    }

    //--------------------------------------------------------------------------------
    // Merged borders
    //--------------------------------------------------------------------------------

    private static ReportBorders ResolveMergedBorders(ReportSheet sourceSheet, ReportCell ownerCell)
    {
        // Resolves the outer borders of a merged cell to the highest-priority border from all cells within the merge
        var mergeInfo = ownerCell.Merge ?? throw new InvalidOperationException("Merged border resolution requires merge info.");
        var mergedCells = sourceSheet.Cells
            .Where(cell => mergeInfo.Range.Contains(cell.Row, cell.Column))
            .ToList();

        return new ReportBorders
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
                static cell => cell.Style.Borders.Left)
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
        foreach (var cell in mergedCells.Where(cell => (fixedSelector(cell) == fixedIndex) && (rangeSelector(cell) >= rangeStart) && (rangeSelector(cell) <= rangeEnd)))
        {
            var border = borderSelector(cell);
            if (bestBorder is null || GetBorderPriority(border.Style) > GetBorderPriority(bestBorder.Style))
            {
                bestBorder = border;
            }
        }

        return bestBorder ?? new ReportBorder();
    }

    private static int GetBorderPriority(BorderLineStyle style) =>
        style switch
        {
            BorderLineStyle.Double => 6,
            BorderLineStyle.Thick => 5,
            BorderLineStyle.Medium => 4,
            BorderLineStyle.Thin => 3,
            BorderLineStyle.DashDot => 2,
            BorderLineStyle.Dashed => 2,
            BorderLineStyle.Dotted => 2,
            BorderLineStyle.Hair => 1,
            _ => 0
        };

    private static string BuildLineKey(ReportLine line)
    {
        return String.Create(
            CultureInfo.InvariantCulture,
            $"{Math.Round(Math.Min(line.X1, line.X2), 4)}:{Math.Round(Math.Min(line.Y1, line.Y2), 4)}:{Math.Round(Math.Max(line.X1, line.X2), 4)}:{Math.Round(Math.Max(line.Y1, line.Y2), 4)}");
    }

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);
    }
}
