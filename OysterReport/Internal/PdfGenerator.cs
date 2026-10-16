namespace OysterReport.Internal;

using System.Globalization;
using System.Text;

using OysterReport.Internal.Pdf;

internal static class PdfGenerator
{
    private const double BoldSimulationOffset = 0.35d;

    private const int MissingGlyphListLimit = 20;

    //--------------------------------------------------------------------------------
    // Write
    //--------------------------------------------------------------------------------

    internal static void WritePdf(ReportRenderContext context, Stream output)
    {
        var writer = new PdfDocumentWriter(output, context.CompressContentStreams);
        if (context.EmbedDocumentMetadata)
        {
            writer.Title = context.Workbook.Metadata.TemplateName;
            writer.CreationDate = DateTimeOffset.Now;
        }

        var options = context.RenderingOptions;
        var fonts = new PdfFontManager(writer, context.FontResolver, FontLocator.Create(options));

        for (var sheetIndex = 0; sheetIndex < context.SheetPlans.Count; sheetIndex++)
        {
            var sheetPlan = context.SheetPlans[sheetIndex];
            var sourceSheet = context.Workbook.Sheets[sheetIndex];
            var missingGlyphs = new Dictionary<PdfEmbeddedFont, SortedSet<int>>();
            foreach (var pagePlan in sheetPlan.Pages)
            {
                var page = writer.BeginPage(pagePlan.PageBounds.Width, pagePlan.PageBounds.Height);
                page.FillRect(PdfColor.White, pagePlan.PageBounds.X, pagePlan.PageBounds.Y, pagePlan.PageBounds.Width, pagePlan.PageBounds.Height);
                DrawCells(page, fonts, missingGlyphs, sourceSheet, pagePlan.Cells, context);
                DrawBorders(page, sourceSheet, pagePlan.Cells, options);
                DrawImages(page, writer, sheetPlan.SheetName, sheetPlan.Images, options);
                DrawHeaderFooter(page, fonts, missingGlyphs, pagePlan.HeaderFooter, pagePlan.PageNumber, sheetPlan.Pages.Count, options);
                writer.EndPage(page);
            }

            ReportMissingGlyphs(sheetPlan.SheetName, missingGlyphs, options);
        }

        writer.Close();
        ReportEmbeddingRestrictions(writer.Fonts, options);
    }

    //--------------------------------------------------------------------------------
    // Warnings
    //--------------------------------------------------------------------------------

    private static void CollectMissingGlyphs(PdfFont font, string text, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs)
    {
        foreach (var rune in text.EnumerateRunes())
        {
            if ((rune.Value < 0x20) || font.Face.Font.HasGlyph(rune.Value))
            {
                continue;
            }

            if (!missingGlyphs.TryGetValue(font.Face, out var codePoints))
            {
                codePoints = [];
                missingGlyphs[font.Face] = codePoints;
            }

            codePoints.Add(rune.Value);
        }
    }

    private static void ReportMissingGlyphs(string sheetName, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs, ReportRenderOption renderOption)
    {
        foreach (var (face, codePoints) in missingGlyphs)
        {
            var name = face.Font.PostScriptName;
            var characters = String.Concat(codePoints.Take(MissingGlyphListLimit).Select(Char.ConvertFromUtf32));
            var more = codePoints.Count > MissingGlyphListLimit ? String.Create(CultureInfo.InvariantCulture, $" (+{codePoints.Count - MissingGlyphListLimit})") : string.Empty;
            renderOption.OnRenderWarning?.Invoke(new ReportRenderWarning
            {
                Kind = ReportRenderWarningKind.MissingGlyph,
                SheetName = sheetName,
                Source = name,
                Message = $"Glyphs not found in font. font=[{name}], characters=[{characters}]{more}"
            });
        }
    }

    private static void ReportEmbeddingRestrictions(IEnumerable<PdfEmbeddedFont> fonts, ReportRenderOption renderOption)
    {
        foreach (var face in fonts.Where(static x => x.IsUsed))
        {
            // OS/2 fsType: usage permission in bits 0-3, 0x0100 no subsetting, 0x0200 bitmap embedding only
            var fsType = face.Font.FsType;
            var reason = (fsType & 0x000F) == 0x0002 ? "restricted license"
                : (fsType & 0x0200) != 0 ? "bitmap embedding only"
                : (fsType & 0x0100) != 0 ? "no subsetting"
                : null;
            if (reason is null)
            {
                continue;
            }

            var name = face.Font.PostScriptName;
            renderOption.OnRenderWarning?.Invoke(new ReportRenderWarning
            {
                Kind = ReportRenderWarningKind.FontEmbeddingRestricted,
                Source = name,
                Message = $"Font license restricts embedding ({reason}); the font is embedded anyway. font=[{name}]"
            });
        }
    }

    //--------------------------------------------------------------------------------
    // Cell
    //--------------------------------------------------------------------------------

    private static void DrawCells(PdfPageCanvas page, PdfFontManager fonts, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs, ReportSheet sourceSheet, IReadOnlyList<PdfCellRenderInfo> cells, ReportRenderContext context)
    {
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(static x => x.Address, StringComparer.Ordinal);

        // Backgrounds are batched by color (one path per color)
        var backgroundGroups = new Dictionary<string, List<ReportRect>>(StringComparer.Ordinal);
        foreach (var renderCell in cells)
        {
            var colorHex = renderCell.BackgroundColorHex;
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
            page.FillRects(ToColor(colorHex), rects);
        }

        foreach (var renderCell in cells)
        {
            if (!sourceCellsByAddress.TryGetValue(renderCell.CellAddress, out var sourceCell))
            {
                continue;
            }

            if (!renderCell.IsMergedOwner && (sourceCell.Merge is not null))
            {
                continue;
            }

            if (String.IsNullOrEmpty(sourceCell.DisplayText))
            {
                continue;
            }

            var font = fonts.Resolve(sourceCell.Style.Font, context.RenderingOptions.DefaultCellFontSize);
            CollectMissingGlyphs(font, sourceCell.DisplayText, missingGlyphs);
            var fontColor = ToColor(sourceCell.Style.Font.ColorHex);
            var content = renderCell.ContentBounds;
            var text = renderCell.TextBounds;

            page.Save();
            page.ClipRect(text.X, text.Y, Math.Max(0, text.Width), Math.Max(0, text.Height));
            DrawCellText(page, sourceCell.DisplayText, font, fontColor, content.X, content.Y, Math.Max(0, content.Width), Math.Max(0, content.Height), sourceCell);
            DrawTextDecorations(page, font, fontColor, content.X, content.Y, Math.Max(0, content.Width), Math.Max(0, content.Height), sourceCell.DisplayText, sourceCell, context.RenderingOptions);
            page.Restore();
        }
    }

    private static void DrawCellText(PdfPageCanvas page, string text, PdfFont font, PdfColor color, double x, double y, double width, double height, ReportCell sourceCell)
    {
        var drawCount = font.SimulateBold ? 2 : 1;
        for (var pass = 0; pass < drawCount; pass++)
        {
            var passX = pass == 0 ? x : x + BoldSimulationOffset;
            if (sourceCell.Style.WrapText || text.Contains('\n', StringComparison.Ordinal))
            {
                PdfTextLayout.DrawWrapped(page, text, font, color, passX, y, width, height, ResolveParagraphAlignment(sourceCell));
                continue;
            }

            var (horizontal, vertical) = ResolveStringFormat(sourceCell);
            PdfTextLayout.DrawString(page, text, font, color, passX, y, width, height, horizontal, vertical);
        }
    }

    private static void DrawTextDecorations(PdfPageCanvas page, PdfFont font, PdfColor color, double x, double y, double width, double height, string text, ReportCell sourceCell, ReportRenderOption renderOption)
    {
        if (sourceCell.Style.Font is { Underline: false, Strikeout: false })
        {
            return;
        }

        if ((width <= 0) || (height <= 0))
        {
            return;
        }

        var isWrap = sourceCell.Style.WrapText || text.Contains('\n', StringComparison.Ordinal);
        var scale = font.Size / font.UnitsPerEm;
        var ascentPt = font.Ascent * scale;

        double verticalOffset = 0;
        if (!isWrap)
        {
            var lineHeight = font.LineSpace;
            verticalOffset = sourceCell.Style.Alignment.Vertical switch
            {
                VerticalAlignment.Center => Math.Max(0, (height - lineHeight) / 2),
                VerticalAlignment.Bottom => Math.Max(0, height - lineHeight),
                _ => 0
            };
        }

        var textTopY = y + verticalOffset;
        double decorationWidth;
        double decorationX;
        if (isWrap)
        {
            decorationWidth = width;
            decorationX = x;
        }
        else
        {
            decorationWidth = font.MeasureWidth(text);
            decorationX = ResolveHorizontalAlignment(sourceCell) switch
            {
                HorizontalAlignment.Center => x + ((width - decorationWidth) / 2),
                HorizontalAlignment.Right => x + width - decorationWidth,
                _ => x
            };
        }

        if (decorationWidth <= 0)
        {
            return;
        }

        if (sourceCell.Style.Font.Underline)
        {
            var thickness = Math.Max(renderOption.UnderlineWidth, Math.Abs(font.UnderlineThickness * scale));
            var lineY = textTopY + ascentPt - (font.UnderlinePosition * scale);
            DrawSolidBorder(page, color, thickness, new ReportLine { X1 = decorationX, Y1 = lineY, X2 = decorationX + decorationWidth, Y2 = lineY });
        }

        if (sourceCell.Style.Font.Strikeout)
        {
            var thickness = Math.Max(renderOption.StrikeoutWidth, Math.Abs(font.StrikeoutThickness * scale));
            var lineY = textTopY + ascentPt - (font.StrikeoutPosition * scale);
            DrawSolidBorder(page, color, thickness, new ReportLine { X1 = decorationX, Y1 = lineY, X2 = decorationX + decorationWidth, Y2 = lineY });
        }
    }

    //--------------------------------------------------------------------------------
    // Border
    //--------------------------------------------------------------------------------

    private static void DrawBorders(PdfPageCanvas page, ReportSheet sourceSheet, IEnumerable<PdfCellRenderInfo> cells, ReportRenderOption renderOption)
    {
        var sourceCellsByAddress = sourceSheet.Cells.ToDictionary(static x => x.Address, StringComparer.Ordinal);
        var collectedLines = new Dictionary<string, (ReportLine Line, ReportBorder Border)>(StringComparer.Ordinal);

        foreach (var renderCell in cells)
        {
            if (!sourceCellsByAddress.TryGetValue(renderCell.CellAddress, out var sourceCell))
            {
                continue;
            }

            if ((sourceCell.Merge is not null) && !renderCell.IsMergedOwner)
            {
                continue;
            }

            var borders = sourceCell.Merge is null ? sourceCell.Style.Borders : ResolveMergedBorders(sourceSheet, sourceCell);
            var b = renderCell.OuterBounds;
            CollectBorderSide(borders.Top, new ReportLine { X1 = b.X, Y1 = b.Y, X2 = b.Right, Y2 = b.Y }, collectedLines);
            CollectBorderSide(borders.Right, new ReportLine { X1 = b.Right, Y1 = b.Y, X2 = b.Right, Y2 = b.Bottom }, collectedLines);
            CollectBorderSide(borders.Bottom, new ReportLine { X1 = b.Right, Y1 = b.Bottom, X2 = b.X, Y2 = b.Bottom }, collectedLines);
            CollectBorderSide(borders.Left, new ReportLine { X1 = b.X, Y1 = b.Bottom, X2 = b.X, Y2 = b.Y }, collectedLines);
        }

        foreach (var (line, border) in collectedLines.Values.OrderBy(static x => GetBorderPriority(x.Border.Style)))
        {
            DrawBorderLine(page, border, line, renderOption);
        }
    }

    private static void CollectBorderSide(ReportBorder border, ReportLine line, Dictionary<string, (ReportLine Line, ReportBorder Border)> collectedLines)
    {
        if (border.Style == BorderLineStyle.None)
        {
            return;
        }

        var lineKey = BuildLineKey(line);
        if (collectedLines.TryGetValue(lineKey, out var existing) &&
            (GetBorderPriority(existing.Border.Style) >= GetBorderPriority(border.Style)))
        {
            return;
        }

        collectedLines[lineKey] = (line, border);
    }

    private static void DrawBorderLine(PdfPageCanvas page, ReportBorder border, ReportLine line, ReportRenderOption renderOption)
    {
        var color = ToColor(border.ColorHex);
        var width = ResolveBorderWidth(border.Style, renderOption);
        if (border.Style == BorderLineStyle.Double)
        {
            DrawDoubleBorder(page, color, width, line);
            return;
        }

        if (IsSolidBorder(border.Style))
        {
            DrawSolidBorder(page, color, width, line);
            return;
        }

        // PDFsharp XDashStyle patterns: dot = width, dash = 3 * width
        var dash = border.Style switch
        {
            BorderLineStyle.Dashed => new[] { 3 * width, width },
            BorderLineStyle.Dotted => [width],
            BorderLineStyle.DashDot => [3 * width, width, width, width],
            _ => []
        };
        page.StrokeLine(color, width, dash, line.X1, line.Y1, line.X2, line.Y2);
    }

    private static void DrawDoubleBorder(PdfPageCanvas page, PdfColor color, double width, ReportLine line)
    {
        var gap = Math.Max(RenderConstants.MinimumDoubleBorderGap, width * RenderConstants.DoubleBorderGapWidthMultiplier);
        if (Math.Abs(line.Y1 - line.Y2) < RenderConstants.StraightLineTolerance)
        {
            DrawSolidBorder(page, color, width, line with { Y1 = line.Y1 - (gap / 2d), Y2 = line.Y2 - (gap / 2d) });
            DrawSolidBorder(page, color, width, line with { Y1 = line.Y1 + (gap / 2d), Y2 = line.Y2 + (gap / 2d) });
            return;
        }

        DrawSolidBorder(page, color, width, line with { X1 = line.X1 - (gap / 2d), X2 = line.X2 - (gap / 2d) });
        DrawSolidBorder(page, color, width, line with { X1 = line.X1 + (gap / 2d), X2 = line.X2 + (gap / 2d) });
    }

    private static void DrawSolidBorder(PdfPageCanvas page, PdfColor color, double width, ReportLine line)
    {
        if (Math.Abs(line.Y1 - line.Y2) < RenderConstants.StraightLineTolerance)
        {
            page.FillRect(color, Math.Min(line.X1, line.X2), line.Y1 - (width / 2d), Math.Abs(line.X2 - line.X1), width);
            return;
        }

        page.FillRect(color, line.X1 - (width / 2d), Math.Min(line.Y1, line.Y2), width, Math.Abs(line.Y2 - line.Y1));
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    private static void DrawImages(PdfPageCanvas page, PdfDocumentWriter writer, string sheetName, IEnumerable<PdfImageRenderInfo> images, ReportRenderOption renderOption)
    {
        foreach (var image in images)
        {
            if (image.ImageBytes.IsEmpty)
            {
                continue;
            }

            try
            {
                var id = AddImage(writer, image.ImageBytes, renderOption);
                page.DrawImage(id, image.Bounds.X, image.Bounds.Y, image.Bounds.Width, image.Bounds.Height);
            }
            catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException or InvalidDataException)
            {
                renderOption.OnRenderWarning?.Invoke(new ReportRenderWarning
                {
                    Kind = ReportRenderWarningKind.ImageDecodeFailed,
                    SheetName = sheetName,
                    Source = image.Name,
                    Message = $"Failed to decode image. name=[{image.Name}]",
                    Exception = ex
                });
            }
        }
    }

    private static int AddImage(PdfDocumentWriter writer, ReadOnlyMemory<byte> bytes, ReportRenderOption renderOption)
    {
        try
        {
            return PdfImageWriter.Write(writer, bytes.Span);
        }
        catch (Exception ex) when ((ex is NotSupportedException or InvalidDataException) && (renderOption.ImageConverter is not null))
        {
            // Formats outside the built-in decoders (e.g. WebP) can be converted by the caller
            var converted = renderOption.ImageConverter(bytes) ?? throw new NotSupportedException("Image converter returned no image.", ex);
            return PdfImageWriter.Write(writer, converted.Span);
        }
    }

    //--------------------------------------------------------------------------------
    // Header / Footer
    //--------------------------------------------------------------------------------

    private static void DrawHeaderFooter(PdfPageCanvas page, PdfFontManager fonts, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs, PdfHeaderFooterRenderInfo headerFooter, int pageNumber, int totalPages, ReportRenderOption renderOption)
    {
        var headerSections = ResolveHeaderFooterSections(headerFooter.HeaderText, pageNumber, totalPages);
        var footerSections = ResolveHeaderFooterSections(headerFooter.FooterText, pageNumber, totalPages);

        // Resolved even when there is nothing to draw, so a missing fallback font is reported as before
        var font = CreateFallbackFont(fonts, renderOption.HeaderFooterFontSize, renderOption.HeaderFooterFallbackFonts);
        if (headerSections.IsEmpty && footerSections.IsEmpty)
        {
            return;
        }

        foreach (var sections in new[] { headerSections, footerSections })
        {
            CollectMissingGlyphs(font, sections.Left + sections.Center + sections.Right, missingGlyphs);
        }

        DrawHeaderFooterSections(page, headerSections, headerFooter.HeaderBounds, font);
        DrawHeaderFooterSections(page, footerSections, headerFooter.FooterBounds, font);
    }

    private static void DrawHeaderFooterSections(PdfPageCanvas page, HeaderFooterSections sections, ReportRect bounds, PdfFont font)
    {
        if (!String.IsNullOrWhiteSpace(sections.Left))
        {
            PdfTextLayout.DrawString(page, sections.Left, font, PdfColor.Black, bounds.X, bounds.Y, bounds.Width, bounds.Height, TextHorizontal.Near, TextVertical.Near);
        }

        if (!String.IsNullOrWhiteSpace(sections.Center))
        {
            PdfTextLayout.DrawString(page, sections.Center, font, PdfColor.Black, bounds.X, bounds.Y, bounds.Width, bounds.Height, TextHorizontal.Center, TextVertical.Near);
        }

        if (!String.IsNullOrWhiteSpace(sections.Right))
        {
            PdfTextLayout.DrawString(page, sections.Right, font, PdfColor.Black, bounds.X, bounds.Y, bounds.Width, bounds.Height, TextHorizontal.Far, TextVertical.Near);
        }
    }

    private static PdfFont CreateFallbackFont(PdfFontManager fonts, double size, IEnumerable<string> fontNames)
    {
        foreach (var fontName in fontNames)
        {
            if (fonts.TryResolveByName(fontName, size) is { } font)
            {
                return font;
            }
        }

        throw new InvalidOperationException("No appropriate fallback font found for header or image drawing.");
    }

    private static HeaderFooterSections ResolveHeaderFooterSections(string? text, int pageNumber, int totalPages)
    {
        if (String.IsNullOrWhiteSpace(text))
        {
            return HeaderFooterSections.Empty;
        }

        var left = new StringBuilder();
        var center = new StringBuilder();
        var right = new StringBuilder();
        var current = center;
        Span<char> scratch = stackalloc char[16];

        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if ((character != '&') || (index == text.Length - 1))
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
                case 'P':
                    if (pageNumber.TryFormat(scratch, out var pn, default, CultureInfo.InvariantCulture))
                    {
                        current.Append(scratch[..pn]);
                    }
                    break;
                case 'N':
                    if (totalPages.TryFormat(scratch, out var tp, default, CultureInfo.InvariantCulture))
                    {
                        current.Append(scratch[..tp]);
                    }
                    break;
                case '&':
                    current.Append('&');
                    break;
            }
        }

        return new HeaderFooterSections(left.ToString().Trim(), center.ToString().Trim(), right.ToString().Trim());
    }

    //--------------------------------------------------------------------------------
    // Color / Style helpers
    //--------------------------------------------------------------------------------

    private static PdfColor ToColor(string colorHex)
    {
        var hex = ColorHelper.NormalizeHex(colorHex).AsSpan().TrimStart('#');
        if (hex.Length == 8)
        {
            return new PdfColor(
                Byte.Parse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                Byte.Parse(hex.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                Byte.Parse(hex.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                Byte.Parse(hex.Slice(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }

        return PdfColor.Black;
    }

    private static bool IsTransparentColor(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);

    private static double ResolveBorderWidth(BorderLineStyle style, ReportRenderOption renderOption) =>
        style switch
        {
            BorderLineStyle.Thick => renderOption.ThickBorderWidth,
            BorderLineStyle.Medium => renderOption.MediumBorderWidth,
            BorderLineStyle.Double => renderOption.NormalBorderWidth,
            BorderLineStyle.Hair => renderOption.HairBorderWidth,
            _ => renderOption.NormalBorderWidth
        };

    private static bool IsSolidBorder(BorderLineStyle style) =>
        style is BorderLineStyle.Hair or BorderLineStyle.Thin or BorderLineStyle.Medium or BorderLineStyle.Thick;

    private static (TextHorizontal Horizontal, TextVertical Vertical) ResolveStringFormat(ReportCell cell) =>
        (ResolveHorizontalAlignment(cell) switch
        {
            HorizontalAlignment.Center => TextHorizontal.Center,
            HorizontalAlignment.Right => TextHorizontal.Far,
            _ => TextHorizontal.Near
        },
        cell.Style.Alignment.Vertical switch
        {
            VerticalAlignment.Center => TextVertical.Center,
            VerticalAlignment.Bottom => TextVertical.Far,
            _ => TextVertical.Near
        });

    private static TextParagraph ResolveParagraphAlignment(ReportCell cell) =>
        ResolveHorizontalAlignment(cell) switch
        {
            HorizontalAlignment.Center => TextParagraph.Center,
            HorizontalAlignment.Right => TextParagraph.Right,
            HorizontalAlignment.Justify => TextParagraph.Justify,
            _ => TextParagraph.Left
        };

    private static HorizontalAlignment ResolveHorizontalAlignment(ReportCell cell)
    {
        if (cell.Style.Alignment.Horizontal != HorizontalAlignment.General)
        {
            return cell.Style.Alignment.Horizontal;
        }

        return cell.Value.Kind switch
        {
            CellValueKind.Number => HorizontalAlignment.Right,
            CellValueKind.DateTime => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
    }

    //--------------------------------------------------------------------------------
    // Merged borders
    //--------------------------------------------------------------------------------

    private static ReportBorders ResolveMergedBorders(ReportSheet sourceSheet, ReportCell ownerCell)
    {
        var mergeInfo = ownerCell.Merge ?? throw new InvalidOperationException("Merged border resolution requires merge info.");
        var mergedCells = sourceSheet.Cells.Where(cell => mergeInfo.Range.Contains(cell.Row, cell.Column)).ToList();

        return new ReportBorders
        {
            Top = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn, static c => c.Row, static c => c.Column, static c => c.Style.Borders.Top),
            Right = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow, static c => c.Column, static c => c.Row, static c => c.Style.Borders.Right),
            Bottom = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn, static c => c.Row, static c => c.Column, static c => c.Style.Borders.Bottom),
            Left = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow, static c => c.Column, static c => c.Row, static c => c.Style.Borders.Left)
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
            if ((bestBorder is null) || (GetBorderPriority(border.Style) > GetBorderPriority(bestBorder.Style)))
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

    private static string BuildLineKey(ReportLine line) =>
        String.Create(
            CultureInfo.InvariantCulture,
            $"{Math.Round(Math.Min(line.X1, line.X2), 4)}:{Math.Round(Math.Min(line.Y1, line.Y2), 4)}:{Math.Round(Math.Max(line.X1, line.X2), 4)}:{Math.Round(Math.Max(line.Y1, line.Y2), 4)}");

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);

        public bool IsEmpty => Left.Length == 0 && Center.Length == 0 && Right.Length == 0;
    }
}
