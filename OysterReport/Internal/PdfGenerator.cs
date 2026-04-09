namespace OysterReport.Internal;

using System.Globalization;
using System.Reflection;
using System.Text;

using ClosedXML.Excel;

using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;

internal static class PdfGenerator
{
    private static int fontPlatformConfigured;
    private const double BoldSimulationOffsetPoints = 0.35d;

    private sealed record ResolvedFontRenderInfo
    {
        public XFont Font { get; init; } = null!;

        public bool SimulateBold { get; init; }
    }

    // レンダリングコンテキストをもとに PDF ドキュメントを生成し、出力ストリームへ書き込む。
    // PDFSharp のフォント設定を初期化し、全シート・全ページを順に描画する。
    internal static void WritePdf(
        ReportRenderContext context,
        Stream output)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(output);

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

    // PDFSharp のフォントリゾルバーが未設定の場合に
    // ReportFontResolverAdapter を登録する (初回のみ実行)。
    private static void EnsurePdfSharpFontConfiguration()
    {
        if (Interlocked.Exchange(ref fontPlatformConfigured, 1) == 1)
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
            fontResolverProperty?.SetValue(null, new ReportFontResolverAdapter());
        }
    }

    // ページ全体を白い背景矩形で塗りつぶす。
    private static void DrawPageBackground(XGraphics graphics, ReportRect pageBounds)
    {
        graphics.DrawRectangle(XBrushes.White, pageBounds.X, pageBounds.Y, pageBounds.Width, pageBounds.Height);
    }

    // セルの背景色とテキストを描画する。
    // 背景は同色のセルをまとめて描画し、テキストは折り返し・中央揃え等に応じて配置する。
    private static void DrawCells(
        XGraphics graphics,
        ReportSheet sourceSheet,
        IReadOnlyList<PdfCellRenderInfo> cells,
        ReportRenderContext context)
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

            var resolvedFont = ResolveFont(sourceCell.Style.Font, context);
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

                DrawCellText(
                    graphics,
                    sourceCell.DisplayText,
                    resolvedFont,
                    textBrush,
                    textRect,
                    sourceCell);
            }
            finally
            {
                graphics.Restore(clipState);
            }
        }
    }

    // セルの罫線を描画する。重複する辺は優先度の高い罫線スタイルを採用し一度だけ描画する。
    private static void DrawBorders(XGraphics graphics, ReportSheet sourceSheet, IEnumerable<PdfCellRenderInfo> cells, ReportRenderingOptions renderingOptions)
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
            DrawBorderLine(graphics, border, line, renderingOptions);
        }
    }

    // 1 辺の罫線情報を収集する。同一辺に複数スタイルがある場合は優先度の高いほうを残す。
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

    // 罫線スタイルに応じて実線・破線・二重線を描画する。
    private static void DrawBorderLine(XGraphics graphics, ReportBorder border, ReportLine line, ReportRenderingOptions renderingOptions)
    {
        var borderColor = ToColor(border.ColorHex);
        var borderWidth = ResolveBorderWidth(border.Style, renderingOptions);
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

    // シートに埋め込まれた画像をページに描画する。デコード失敗時はスキップする。
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

    // ヘッダー・フッターのテキストを左・中央・右の各領域に描画する。
    private static void DrawHeaderFooter(
        XGraphics graphics,
        PdfHeaderFooterRenderInfo headerFooter,
        int pageNumber,
        int totalPages,
        ReportRenderingOptions renderingOptions)
    {
        var headerSections = ResolveHeaderFooterSections(headerFooter.HeaderText, pageNumber, totalPages);
        var footerSections = ResolveHeaderFooterSections(headerFooter.FooterText, pageNumber, totalPages);
        var font = CreateFallbackFont(renderingOptions.HeaderFooterFontSizePoints, renderingOptions.HeaderFooterFallbackFontNames);

        DrawHeaderFooterSections(graphics, headerSections, headerFooter.HeaderBounds, font);
        DrawHeaderFooterSections(graphics, footerSections, headerFooter.FooterBounds, font);
    }

    // 左・中央・右に分割されたヘッダー/フッター文字列をそれぞれ描画する。
    private static void DrawHeaderFooterSections(
        XGraphics graphics,
        HeaderFooterSections sections,
        ReportRect bounds,
        XFont font)
    {
        if (!string.IsNullOrWhiteSpace(sections.Left))
        {
            graphics.DrawString(sections.Left, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopLeft);
        }

        if (!string.IsNullOrWhiteSpace(sections.Center))
        {
            graphics.DrawString(sections.Center, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopCenter);
        }

        if (!string.IsNullOrWhiteSpace(sections.Right))
        {
            graphics.DrawString(sections.Right, font, XBrushes.Black, new XRect(bounds.X, bounds.Y, bounds.Width, bounds.Height), XStringFormats.TopRight);
        }
    }

    // Excel ヘッダー/フッター書式文字列 (&L, &C, &R, &P, &N) を解析し左・中央・右とページ番号に分解する。
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

    // ReportFont 属性とフォントリゾルバーから PDFSharp 用 XFont を生成する。
    // リゾルバーが返した名前を XFont に渡すことで PDFSharp のフォントキャッシュを
    // リゾルバーごとに分離し、埋め込みフォントが確実に使われるようにする。
    private static ResolvedFontRenderInfo ResolveFont(ReportFont font, ReportRenderContext context)
    {
        var fontSize = font.Size <= 0 ? context.RenderingOptions.DefaultCellFontSizePoints : font.Size;
        var nameToUse = font.Name;
        var simulateBold = false;

        if (context.FontResolver is not null)
        {
            var request = new ReportFontRequest
            {
                FontName = font.Name
            };

            var resolvedName = context.FontResolver.ResolveFaceName(request);
            if (!string.IsNullOrWhiteSpace(resolvedName))
            {
                var embeddedFontData = context.FontResolver.GetFontData(resolvedName);

                if (embeddedFontData is { } fontData)
                {
                    // 埋め込みフォントをアダプタに事前登録する。
                    // 同じバイト列を複数回登録してもべき等であるため問題ない。
                    ReportFontResolverAdapter.RegisterEmbeddedFont(resolvedName, fontData);

                    // 単一の埋め込みフォント資源を返した場合、Bold は描画時にシミュレーションする。
                    // Italic は ReportFontResolverAdapter が PDFsharp の公式シミュレーションへ委譲する。
                    simulateBold = font.Bold;
                }

                nameToUse = resolvedName;
            }
        }

        if (!simulateBold && font.Bold && ReportFontResolverAdapter.NeedsBoldSimulationForInstalledFont(nameToUse, font.Italic))
        {
            simulateBold = true;
        }

        var style = BuildActualFontStyle(font, simulateBold);

        if (!string.IsNullOrWhiteSpace(nameToUse) && TryCreateFont(nameToUse, fontSize, style, out var resolvedFont))
        {
            return new ResolvedFontRenderInfo
            {
                Font = resolvedFont,
                SimulateBold = simulateBold
            };
        }

        // リゾルバーが返した名前で失敗した場合は元の Excel フォント名にフォールバックする。
        if (!string.Equals(nameToUse, font.Name, StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(font.Name) &&
            TryCreateFont(font.Name, fontSize, BuildActualFontStyle(font, simulateBold: false), out var fallbackFont))
        {
            return new ResolvedFontRenderInfo
            {
                Font = fallbackFont
            };
        }

        throw new InvalidOperationException($"No appropriate font found for family name '{font.Name}'.");
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
                : new XRect(textRect.X + BoldSimulationOffsetPoints, textRect.Y, textRect.Width, textRect.Height);

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
                        ResolveStringFormat(sourceCell));
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

    // ヘッダー/フッター用のフォールバックフォントを候補一覧から作成する。
    private static XFont CreateFallbackFont(double size, IEnumerable<string> fontNames)
    {
        foreach (var fontName in fontNames)
        {
            if (TryCreateFont(fontName, size, XFontStyleEx.Regular, out var font))
            {
                return font;
            }
        }

        throw new InvalidOperationException("No appropriate fallback font found for header or image drawing.");
    }

    // 指定名で XFont の作成を試み、失敗した場合は false を返す。
    private static bool TryCreateFont(string fontName, double size, XFontStyleEx style, out XFont font)
    {
        try
        {
            font = new XFont(fontName, size, style, new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.TryComputeSubset));
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

    // ARGB 形式のカラー文字列を XColor に変換する。
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

    // アルファで完全透明な色かどうかを判定する。
    private static bool IsTransparentColor(string colorHex) =>
        ColorHelper.NormalizeHex(colorHex).StartsWith("#00", StringComparison.Ordinal);

    // 罫線スタイルから描画幅 (pt) を決定する。
    private static double ResolveBorderWidth(XLBorderStyleValues style, ReportRenderingOptions renderingOptions) =>
        style switch
        {
            XLBorderStyleValues.Thick => renderingOptions.ThickBorderWidthPoints,
            XLBorderStyleValues.Medium => renderingOptions.MediumBorderWidthPoints,
            XLBorderStyleValues.Double => renderingOptions.NormalBorderWidthPoints,
            XLBorderStyleValues.Hair => renderingOptions.HairBorderWidthPoints,
            _ => renderingOptions.NormalBorderWidthPoints
        };

    // 罫線スタイルに応じた破線パターンを XPen に適用する。
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

    // 二重線罫線を並列に 2 本の実線で描画する。
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

    // セルの水平・垂直配置から XStringFormat を生成する。
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

    // セルの水平配置から段落整列エニュムを返す。
    private static XParagraphAlignment ResolveParagraphAlignment(ReportCell cell) =>
        ResolveHorizontalAlignment(cell) switch
        {
            XLAlignmentHorizontalValues.Center => XParagraphAlignment.Center,
            XLAlignmentHorizontalValues.Right => XParagraphAlignment.Right,
            XLAlignmentHorizontalValues.Justify => XParagraphAlignment.Justify,
            _ => XParagraphAlignment.Left
        };

    // セルの配置設定から水平整列を決定する。General の場合は値種別の既定値を使う。
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

    // 実線罫線を矩形塗りつぶしで描画する（血流防止のため DrawLine の代わりに使用）。
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

    // 実線として矩形描画すべき罫線スタイルかどうかを判定する。
    private static bool IsSolidBorder(XLBorderStyleValues style) =>
        style is XLBorderStyleValues.Hair or XLBorderStyleValues.Thin or XLBorderStyleValues.Medium or XLBorderStyleValues.Thick;

    // マージセルの外周罫線をマージ内の各セルから最高優先度で利用するものに解決する。
    private static ReportBorders ResolveMergedBorders(ReportSheet sourceSheet, ReportCell ownerCell)
    {
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
