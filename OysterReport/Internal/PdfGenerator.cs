namespace OysterReport.Internal;

using System.Globalization;
using System.Runtime.InteropServices;
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
        var options = context.RenderingOptions;

        var now = options.TimeProvider.GetLocalNow();
        var culture = options.HeaderFooterCulture ?? CultureInfo.CurrentCulture;

        var writer = new PdfDocumentWriter(output, context.CompressContentStreams, context.CompressionLevel);
        if (context.EmbedDocumentMetadata)
        {
            writer.Title = context.Workbook.Metadata.TemplateName;
            writer.CreationDate = now;
        }

        var fonts = new PdfFontManager(writer, context.FontResolver, FontLocator.Create(options));
        var images = new PdfImageCache();

        for (var sheetIndex = 0; sheetIndex < context.SheetPlans.Count; sheetIndex++)
        {
            var sheetPlan = context.SheetPlans[sheetIndex];
            var sourceSheet = context.Workbook.Sheets[sheetIndex];
            var missingGlyphs = new Dictionary<PdfEmbeddedFont, SortedSet<int>>();
            foreach (var pagePlan in sheetPlan.Pages)
            {
                var page = writer.BeginPage(pagePlan.PageBounds.Width, pagePlan.PageBounds.Height);
                page.FillRect(PdfColor.White, pagePlan.PageBounds.X, pagePlan.PageBounds.Y, pagePlan.PageBounds.Width, pagePlan.PageBounds.Height);
                DrawCells(page, fonts, missingGlyphs, pagePlan.Cells, context);

                DrawImages(page, writer, images, sheetPlan.SheetName, sheetPlan.Images.Where(static x => x.InCell), options);
                DrawBorders(page, sourceSheet, pagePlan.Cells, options);
                DrawImages(page, writer, images, sheetPlan.SheetName, sheetPlan.Images.Where(static x => !x.InCell), options);
                var fields = new HeaderFooterFields(pagePlan.PageNumber, pagePlan.TotalPages, sheetPlan.SheetName, sheetPlan.FileName, now, culture);
                DrawHeaderFooter(page, fonts, missingGlyphs, pagePlan.HeaderFooter, fields, options);
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

    private static void DrawCells(PdfPageCanvas page, PdfFontManager fonts, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs, IReadOnlyList<PdfCellRenderInfo> cells, ReportRenderContext context)
    {
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
            page.FillRects(ToColor(colorHex), CollectionsMarshal.AsSpan(rects));
        }

        foreach (var renderCell in cells)
        {
            var sourceCell = renderCell.Cell;
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

    private static void DrawBorders(PdfPageCanvas page, ReportSheet sourceSheet, IReadOnlyList<PdfCellRenderInfo> cells, ReportRenderOption renderOption)
    {
        var collectedLines = new Dictionary<LineKey, (ReportLine Line, ReportBorder Border)>(EstimateBorderSegments(cells));

        foreach (var renderCell in cells)
        {
            var sourceCell = renderCell.Cell;
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

        var junctions = CollectJunctions(collectedLines, renderOption);

        // Stable by priority (same order as a stable sort of the insertion order)
        for (var priority = MinBorderPriority; priority <= MaxBorderPriority; priority++)
        {
            foreach (var (line, border) in collectedLines.Values)
            {
                if (GetBorderPriority(border.Style) == priority)
                {
                    DrawBorderLine(page, border, line, junctions, renderOption);
                }
            }
        }
    }

    private static int EstimateBorderSegments(IEnumerable<PdfCellRenderInfo> cells)
    {
        var sides = 0;
        foreach (var renderCell in cells)
        {
            var borders = renderCell.Cell.Style.Borders;
            sides += (borders.Top.Style != BorderLineStyle.None ? 1 : 0) +
                (borders.Right.Style != BorderLineStyle.None ? 1 : 0) +
                (borders.Bottom.Style != BorderLineStyle.None ? 1 : 0) +
                (borders.Left.Style != BorderLineStyle.None ? 1 : 0);
        }

        return sides / 2;
    }

    private static void CollectBorderSide(ReportBorder border, ReportLine line, Dictionary<LineKey, (ReportLine Line, ReportBorder Border)> collectedLines)
    {
        if (border.Style == BorderLineStyle.None)
        {
            return;
        }

        ref var entry = ref CollectionsMarshal.GetValueRefOrAddDefault(collectedLines, LineKey.Create(line), out var exists);
        if (exists && (GetBorderPriority(entry.Border.Style) >= GetBorderPriority(border.Style)))
        {
            return;
        }

        entry = (line, border);
    }

    private static Dictionary<PointKey, Junction> CollectJunctions(Dictionary<LineKey, (ReportLine Line, ReportBorder Border)> collectedLines, ReportRenderOption renderOption)
    {
        var junctions = new Dictionary<PointKey, Junction>(collectedLines.Count);
        foreach (var (line, border) in collectedLines.Values)
        {
            var arm = new BorderArm(GetBorderPriority(border.Style), ResolveBorderShape(border.Style, renderOption));
            if (IsHorizontal(line))
            {
                SetArm(ref GetJunction(junctions, Math.Min(line.X1, line.X2), line.Y1).Right, arm);
                SetArm(ref GetJunction(junctions, Math.Max(line.X1, line.X2), line.Y1).Left, arm);
            }
            else
            {
                SetArm(ref GetJunction(junctions, line.X1, Math.Min(line.Y1, line.Y2)).Down, arm);
                SetArm(ref GetJunction(junctions, line.X1, Math.Max(line.Y1, line.Y2)).Up, arm);
            }
        }

        return junctions;
    }

    private static ref Junction GetJunction(Dictionary<PointKey, Junction> junctions, double x, double y) =>
        ref CollectionsMarshal.GetValueRefOrAddDefault(junctions, PointKey.Create(x, y), out _);

    private static void SetArm(ref BorderArm? arm, BorderArm candidate)
    {
        // Overlapping segments (merged cells) keep the one drawn on top
        if ((arm is null) || (arm.Value.Priority < candidate.Priority))
        {
            arm = candidate;
        }
    }

    private static void DrawBorderLine(PdfPageCanvas page, ReportBorder border, ReportLine line, Dictionary<PointKey, Junction> junctions, ReportRenderOption renderOption)
    {
        var color = ToColor(border.ColorHex);
        var (width, offset) = ResolveBorderShape(border.Style, renderOption);
        junctions.TryGetValue(PointKey.Create(line.X1, line.Y1), out var start);
        junctions.TryGetValue(PointKey.Create(line.X2, line.Y2), out var end);
        if (border.Style == BorderLineStyle.Double)
        {
            DrawSolidBorder(page, color, width, ResolveStroke(line, start, end, -offset, width));
            DrawSolidBorder(page, color, width, ResolveStroke(line, start, end, offset, width));
            return;
        }

        var stroke = ResolveStroke(line, start, end, 0d, width);
        if (IsSolidBorder(border.Style))
        {
            DrawSolidBorder(page, color, width, stroke);
            return;
        }

        Span<double> dash = stackalloc double[4];
        var dashLength = border.Style switch
        {
            BorderLineStyle.Dashed => SetDash(dash, 3 * width, width),
            BorderLineStyle.Dotted => SetDash(dash, width),
            BorderLineStyle.DashDot => SetDash(dash, 3 * width, width, width, width),
            _ => 0
        };
        page.StrokeLine(color, width, dash[..dashLength], stroke.X1, stroke.Y1, stroke.X2, stroke.Y2);
    }

    private static int SetDash(Span<double> destination, params ReadOnlySpan<double> pattern)
    {
        pattern.CopyTo(destination);
        return pattern.Length;
    }

    private static ReportLine ResolveStroke(ReportLine line, in Junction start, in Junction end, double offset, double width)
    {
        var halfWidth = width / 2d;
        if (IsHorizontal(line))
        {
            var (x1, x2) = Extend(line.X1, line.X2, ResolveExtension(start.Up, start.Down, offset, halfWidth), ResolveExtension(end.Up, end.Down, offset, halfWidth));
            return new ReportLine { X1 = x1, Y1 = line.Y1 + offset, X2 = x2, Y2 = line.Y2 + offset };
        }

        var (y1, y2) = Extend(line.Y1, line.Y2, ResolveExtension(start.Left, start.Right, offset, halfWidth), ResolveExtension(end.Left, end.Right, offset, halfWidth));
        return new ReportLine { X1 = line.X1 + offset, Y1 = y1, X2 = line.X2 + offset, Y2 = y2 };
    }

    private static double ResolveExtension(BorderArm? negative, BorderArm? positive, double offset, double halfWidth)
    {
        var extension = Double.NegativeInfinity;
        if ((offset - halfWidth) <= 0)
        {
            extension = negative?.Shape.InnerReach ?? positive?.Shape.HalfExtent ?? 0d;
        }

        if ((offset + halfWidth) >= 0)
        {
            extension = Math.Max(extension, positive?.Shape.InnerReach ?? negative?.Shape.HalfExtent ?? 0d);
        }

        return extension;
    }

    private static (double From, double To) Extend(double from, double to, double fromExtension, double toExtension)
    {
        var direction = to >= from ? 1d : -1d;
        var extendedFrom = from - (direction * fromExtension);
        var extendedTo = to + (direction * toExtension);
        return ((extendedTo - extendedFrom) * direction) >= 0 ? (extendedFrom, extendedTo) : (extendedFrom, extendedFrom);
    }

    private static bool IsHorizontal(ReportLine line) =>
        Math.Abs(line.Y1 - line.Y2) < RenderConstants.StraightLineTolerance;

    private static void DrawSolidBorder(PdfPageCanvas page, PdfColor color, double width, ReportLine line)
    {
        if (IsHorizontal(line))
        {
            page.FillRect(color, Math.Min(line.X1, line.X2), line.Y1 - (width / 2d), Math.Abs(line.X2 - line.X1), width);
            return;
        }

        page.FillRect(color, line.X1 - (width / 2d), Math.Min(line.Y1, line.Y2), width, Math.Abs(line.Y2 - line.Y1));
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    private static void DrawImages(PdfPageCanvas page, PdfDocumentWriter writer, PdfImageCache cache, string sheetName, IEnumerable<PdfImageRenderInfo> images, ReportRenderOption renderOption)
    {
        foreach (var image in images)
        {
            if (image.ImageBytes.IsEmpty)
            {
                continue;
            }

            try
            {
                if (!cache.TryGet(image.ImageBytes, out var pdfImage))
                {
                    pdfImage = AddImage(writer, image.ImageBytes, renderOption);
                    cache.Add(image.ImageBytes, pdfImage);
                }

                var bounds = ResolveImageBounds(image, pdfImage.Width, pdfImage.Height);
                page.DrawImage(pdfImage.Id, bounds.X, bounds.Y, bounds.Width, bounds.Height);
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

    private static PdfImage AddImage(PdfDocumentWriter writer, ReadOnlyMemory<byte> bytes, ReportRenderOption renderOption)
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

    private static ReportRect ResolveImageBounds(PdfImageRenderInfo image, int width, int height)
    {
        var bounds = image.Bounds;
        if ((image.Fit != ReportImageFit.Contain) || (width <= 0) || (height <= 0))
        {
            return bounds;
        }

        var scale = Math.Min(bounds.Width / width, bounds.Height / height);
        return new ReportRect
        {
            X = bounds.X + ((bounds.Width - (width * scale)) / 2d),
            Y = bounds.Y + ((bounds.Height - (height * scale)) / 2d),
            Width = width * scale,
            Height = height * scale
        };
    }

    //--------------------------------------------------------------------------------
    // Header / Footer
    //--------------------------------------------------------------------------------

    private static void DrawHeaderFooter(PdfPageCanvas page, PdfFontManager fonts, Dictionary<PdfEmbeddedFont, SortedSet<int>> missingGlyphs, PdfHeaderFooterRenderInfo headerFooter, HeaderFooterFields fields, ReportRenderOption renderOption)
    {
        var headerSections = ResolveHeaderFooterSections(headerFooter.HeaderText, fields);
        var footerSections = ResolveHeaderFooterSections(headerFooter.FooterText, fields);

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

    private static HeaderFooterSections ResolveHeaderFooterSections(string? text, HeaderFooterFields fields)
    {
        if (String.IsNullOrWhiteSpace(text))
        {
            return HeaderFooterSections.Empty;
        }

        var left = new StringBuilder();
        var center = new StringBuilder();
        var right = new StringBuilder();
        var current = center;
        Span<char> scratch = stackalloc char[64];

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
                    if (fields.PageNumber.TryFormat(scratch, out var pn, default, CultureInfo.InvariantCulture))
                    {
                        current.Append(scratch[..pn]);
                    }
                    break;
                case 'N':
                    if (fields.TotalPages.TryFormat(scratch, out var tp, default, CultureInfo.InvariantCulture))
                    {
                        current.Append(scratch[..tp]);
                    }
                    break;
                case 'D':
                    if (fields.Now.TryFormat(scratch, out var date, "d", fields.Culture))
                    {
                        current.Append(scratch[..date]);
                    }
                    break;
                case 'T':
                    if (fields.Now.TryFormat(scratch, out var time, "t", fields.Culture))
                    {
                        current.Append(scratch[..time]);
                    }
                    break;
                case 'A': current.Append(fields.SheetName); break;
                case 'F': current.Append(fields.FileName); break;
                case '&':
                    current.Append('&');
                    break;
                case '"':
                    // Font name and style (&"name,style") are not applied: skip to the closing quote
                    var closingQuote = text.IndexOf('"', index + 1);
                    index = closingQuote < 0 ? text.Length - 1 : closingQuote;
                    break;
                case 'K':
                    // Font color (&KRRGGBB, or &KTTSNNN for a theme color) is not applied
                    index = Math.Min(index + 6, text.Length - 1);
                    break;
                case >= '0' and <= '9':
                    while ((index + 1 < text.Length) && Char.IsAsciiDigit(text[index + 1]))
                    {
                        index++;
                    }
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
        if (String.IsNullOrWhiteSpace(colorHex))
        {
            return default;
        }

        var hex = colorHex.AsSpan().Trim().TrimStart('#');
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

    private static bool IsTransparentColor(string colorHex)
    {
        if (String.IsNullOrWhiteSpace(colorHex))
        {
            return true;
        }

        var trimmed = colorHex.AsSpan().Trim();
        return (trimmed[0] == '#' ? trimmed[1..] : trimmed).StartsWith("00", StringComparison.Ordinal);
    }

    private static double ResolveBorderWidth(BorderLineStyle style, ReportRenderOption renderOption) =>
        style switch
        {
            BorderLineStyle.Thick => renderOption.ThickBorderWidth,
            BorderLineStyle.Medium => renderOption.MediumBorderWidth,
            BorderLineStyle.Double => renderOption.NormalBorderWidth,
            BorderLineStyle.Hair => renderOption.HairBorderWidth,
            _ => renderOption.NormalBorderWidth
        };

    private static BorderShape ResolveBorderShape(BorderLineStyle style, ReportRenderOption renderOption)
    {
        var width = ResolveBorderWidth(style, renderOption);
        return style == BorderLineStyle.Double
            ? new BorderShape(width, Math.Max(RenderConstants.MinimumDoubleBorderGap, width * RenderConstants.DoubleBorderGapWidthMultiplier) / 2d)
            : new BorderShape(width, 0d);
    }

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
        var mergedCells = CollectMergedCells(sourceSheet, mergeInfo.Range);

        return new ReportBorders
        {
            Top = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn, static c => c.Row, static c => c.Column, static c => c.Style.Borders.Top),
            Right = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow, static c => c.Column, static c => c.Row, static c => c.Style.Borders.Right),
            Bottom = ResolveMergedBorder(mergedCells, mergeInfo.Range.EndRow, mergeInfo.Range.StartColumn, mergeInfo.Range.EndColumn, static c => c.Row, static c => c.Column, static c => c.Style.Borders.Bottom),
            Left = ResolveMergedBorder(mergedCells, mergeInfo.Range.StartColumn, mergeInfo.Range.StartRow, mergeInfo.Range.EndRow, static c => c.Column, static c => c.Row, static c => c.Style.Borders.Left)
        };
    }

    private static List<ReportCell> CollectMergedCells(ReportSheet sheet, ReportRange range)
    {
        if (!sheet.IsCellOrderSorted)
        {
            return sheet.Cells.Where(cell => range.Contains(cell.Row, cell.Column)).ToList();
        }

        // Cells are in (row, column) order, so walking the range through the index yields the same sequence
        var result = new List<ReportCell>();
        for (var row = range.StartRow; row <= range.EndRow; row++)
        {
            for (var column = range.StartColumn; column <= range.EndColumn; column++)
            {
                if (sheet.FindCell(row, column) is { } cell)
                {
                    result.Add(cell);
                }
            }
        }

        return result;
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

    private const int MinBorderPriority = 0;

    private const int MaxBorderPriority = 6;

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

    private readonly record struct LineKey(long Left, long Top, long Right, long Bottom)
    {
        public static LineKey Create(ReportLine line) =>
            new(
                BitConverter.DoubleToInt64Bits(Math.Round(Math.Min(line.X1, line.X2), 4)),
                BitConverter.DoubleToInt64Bits(Math.Round(Math.Min(line.Y1, line.Y2), 4)),
                BitConverter.DoubleToInt64Bits(Math.Round(Math.Max(line.X1, line.X2), 4)),
                BitConverter.DoubleToInt64Bits(Math.Round(Math.Max(line.Y1, line.Y2), 4)));
    }

    private readonly record struct PointKey(long X, long Y)
    {
        public static PointKey Create(double x, double y) =>
            new(BitConverter.DoubleToInt64Bits(Math.Round(x, 4)), BitConverter.DoubleToInt64Bits(Math.Round(y, 4)));
    }

    private readonly record struct BorderShape(double Width, double Offset)
    {
        // From the center line to the outer edge
        public double HalfExtent => Offset + (Width / 2d);

        // Past the center line to the far edge of the nearest stroke, where a line ending inside this one stops
        public double InnerReach => (Width / 2d) - Offset;
    }

    private readonly record struct BorderArm(int Priority, BorderShape Shape);

    private struct Junction
    {
        public BorderArm? Left;
        public BorderArm? Up;
        public BorderArm? Right;
        public BorderArm? Down;
    }

    private sealed record HeaderFooterSections(string Left, string Center, string Right)
    {
        public static HeaderFooterSections Empty { get; } = new(string.Empty, string.Empty, string.Empty);

        public bool IsEmpty => Left.Length == 0 && Center.Length == 0 && Right.Length == 0;
    }

    private readonly record struct HeaderFooterFields(int PageNumber, int TotalPages, string SheetName, string FileName, DateTimeOffset Now, CultureInfo Culture);
}
