// Main entry point: opens the .xlsx package via DocumentFormat.OpenXml and produces a ReportWorkbook directly,
// bypassing ClosedXML and the second ExcelReader pass.

namespace OysterReport.Prototype;

using DocumentFormat.OpenXml.Packaging;

using OysterReport.Internal;

internal static class OpenXmlLoader
{
    public static ReportWorkbook Load(Stream stream, ReportRenderOption? renderOption = null)
    {
        var effectiveOptions = renderOption ?? new ReportRenderOption();
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        return Build(doc, effectiveOptions);
    }

    public static ReportWorkbook Load(string path, ReportRenderOption? renderOption = null)
    {
        var effectiveOptions = renderOption ?? new ReportRenderOption();
        using var doc = SpreadsheetDocument.Open(path, isEditable: false);
        return Build(doc, effectiveOptions);
    }

    private static ReportWorkbook Build(SpreadsheetDocument doc, ReportRenderOption renderOption)
    {
        var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("Workbook part missing.");
        var styles = StyleCatalog.Load(workbookPart);
        var sharedStrings = SharedStringCatalog.Load(workbookPart.SharedStringTablePart);
        var measurementProfile = CreateMeasurementProfile(styles, renderOption);
        var printAreas = ReadPrintAreas(workbookPart);

        var metadata = new ReportMetadata
        {
            TemplateName = doc.PackageProperties.Title ?? "Workbook"
        };

        var workbook = new ReportWorkbook
        {
            Metadata = metadata,
            MeasurementProfile = measurementProfile
        };

        var loader = new WorksheetLoader(styles, sharedStrings, measurementProfile);

        if (workbookPart.Workbook.Sheets is null)
        {
            return workbook;
        }

        var sheetIndex = 0;
        foreach (var sheetRef in workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
        {
            if (sheetRef.Id?.Value is not { } relId)
            {
                continue;
            }

            if (workbookPart.GetPartById(relId) is not WorksheetPart wsPart)
            {
                continue;
            }

            var name = sheetRef.Name?.Value ?? string.Empty;
            printAreas.TryGetValue(sheetIndex, out var printArea);
            var reportSheet = loader.Load(wsPart, name, printArea);

            foreach (var img in DrawingLoader.Load(wsPart))
            {
                reportSheet.AddImage(img);
            }

            ApplyTableStyles(reportSheet, wsPart, styles.ColorResolver);

            workbook.AddSheet(reportSheet);
            sheetIndex++;
        }

        return workbook;
    }

    private static Dictionary<int, ReportPrintArea> ReadPrintAreas(WorkbookPart workbookPart)
    {
        var result = new Dictionary<int, ReportPrintArea>();
        var definedNames = workbookPart.Workbook.DefinedNames;
        if (definedNames is null)
        {
            return result;
        }

        foreach (var dn in definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>())
        {
            if (dn.Name?.Value != "_xlnm.Print_Area" || dn.LocalSheetId?.Value is not { } sheetId)
            {
                continue;
            }

            var text = dn.Text?.Trim();
            if (String.IsNullOrEmpty(text))
            {
                continue;
            }

            // Expected form: SheetName!$A$1:$E$58  (strip everything up to and including !)
            var bang = text.LastIndexOf('!');
            var refPart = bang >= 0 ? text[(bang + 1)..] : text;
            refPart = refPart.Replace("$", string.Empty, StringComparison.Ordinal);
            var range = ParsePrintAreaRange(refPart);
            if (range is not null)
            {
                result[(int)sheetId] = new ReportPrintArea { Range = range.Value };
            }
        }

        return result;
    }

    private static ReportRange? ParsePrintAreaRange(string reference)
    {
        var colonIdx = reference.IndexOf(':', StringComparison.Ordinal);
        if (colonIdx < 0)
        {
            var (row, col) = AddressHelper.ParseAddress(reference);
            return new ReportRange { StartRow = row, StartColumn = col, EndRow = row, EndColumn = col };
        }

        var (r1, c1) = AddressHelper.ParseAddress(reference[..colonIdx]);
        var (r2, c2) = AddressHelper.ParseAddress(reference[(colonIdx + 1)..]);
        return new ReportRange
        {
            StartRow = Math.Min(r1, r2),
            StartColumn = Math.Min(c1, c2),
            EndRow = Math.Max(r1, r2),
            EndColumn = Math.Max(c1, c2)
        };
    }

    private static ReportMeasurementProfile CreateMeasurementProfile(StyleCatalog styles, ReportRenderOption renderOption)
    {
        // Use the "Normal" style's font (fonts[0] per ECMA-376 convention) — matches ClosedXML's workbook.Style.Font.
        var maxDigitWidth = ResolveMaxDigitWidth(styles.DefaultFontName, styles.DefaultFontSize, renderOption);
        return new ReportMeasurementProfile
        {
            MaxDigitWidth = maxDigitWidth,
            ColumnWidthAdjustment = renderOption.ColumnWidthAdjustment
        };
    }

    private static double ResolveMaxDigitWidth(string? fontName, double fontSize, ReportRenderOption renderOption)
    {
        if (String.IsNullOrWhiteSpace(fontName) || fontSize <= 0d)
        {
            return renderOption.FallbackMaxDigitWidth;
        }

        var directMeasured = FontMetricsHelper.MeasureMaxDigitWidth(fontName, fontSize);
        if (directMeasured is > 0d)
        {
            return Math.Max(renderOption.FallbackMaxDigitWidth, directMeasured.Value);
        }

        return renderOption.FallbackMaxDigitWidth;
    }

    private static void ApplyTableStyles(ReportSheet sheet, WorksheetPart wsPart, ColorResolver colorResolver)
    {
        // Mirror ExcelReader.ApplyTableStyles: for each table with ShowRowStripes, paint every other data row
        // with the theme's band1 fill — but only for cells that are currently transparent.
        foreach (var table in TableLoader.Load(wsPart))
        {
            if (!table.ShowRowStripes || String.IsNullOrEmpty(table.ThemeName))
            {
                continue;
            }

            if (!TryResolveStripeHex(table.ThemeName, colorResolver, out var stripeHex))
            {
                continue;
            }

            var firstDataRow = table.Range.StartRow + (table.ShowHeader ? 1 : 0);
            var lastDataRow = table.Range.EndRow - (table.ShowTotals ? 1 : 0);

            for (var r = firstDataRow; r <= lastDataRow; r++)
            {
                if (((r - firstDataRow) % 2) != 0)
                {
                    continue;
                }

                foreach (var cell in sheet.Cells)
                {
                    if (cell.Row != r || cell.Column < table.Range.StartColumn || cell.Column > table.Range.EndColumn)
                    {
                        continue;
                    }

                    if (!cell.Style.Fill.BackgroundColorHex.StartsWith("#00", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    cell.Style = cell.Style with
                    {
                        Fill = new ReportFill { BackgroundColorHex = stripeHex }
                    };
                }
            }
        }
    }

    // Mirror OysterReport.Internal.TableStyleCatalog.Band1RowByStyleName — theme color index (0=lt1…9=accent6) + tint.
    // Neutral styles (Light1/8/15, Medium1/8/15/22) are omitted as they have no stripe.
    private static readonly Dictionary<string, (int ThemeIndex, double Tint)> StripeBandByStyleName =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["TableStyleLight2"] = (4, 0.8),
            ["TableStyleLight3"] = (5, 0.8),
            ["TableStyleLight4"] = (6, 0.8),
            ["TableStyleLight5"] = (7, 0.8),
            ["TableStyleLight6"] = (8, 0.8),
            ["TableStyleLight7"] = (9, 0.8),
            ["TableStyleLight9"] = (4, 0.8),
            ["TableStyleLight10"] = (5, 0.8),
            ["TableStyleLight11"] = (6, 0.8),
            ["TableStyleLight12"] = (7, 0.8),
            ["TableStyleLight13"] = (8, 0.8),
            ["TableStyleLight14"] = (9, 0.8),
            ["TableStyleLight16"] = (4, 0.8),
            ["TableStyleLight17"] = (5, 0.8),
            ["TableStyleLight18"] = (6, 0.8),
            ["TableStyleLight19"] = (7, 0.8),
            ["TableStyleLight20"] = (8, 0.8),
            ["TableStyleLight21"] = (9, 0.8),
            ["TableStyleMedium2"] = (4, 0.2),
            ["TableStyleMedium3"] = (5, 0.2),
            ["TableStyleMedium4"] = (6, 0.2),
            ["TableStyleMedium5"] = (7, 0.2),
            ["TableStyleMedium6"] = (8, 0.2),
            ["TableStyleMedium7"] = (9, 0.2),
            ["TableStyleMedium9"] = (4, 0.2),
            ["TableStyleMedium10"] = (5, 0.2),
            ["TableStyleMedium11"] = (6, 0.2),
            ["TableStyleMedium12"] = (7, 0.2),
            ["TableStyleMedium13"] = (8, 0.2),
            ["TableStyleMedium14"] = (9, 0.2),
            ["TableStyleMedium16"] = (4, 0.2),
            ["TableStyleMedium17"] = (5, 0.2),
            ["TableStyleMedium18"] = (6, 0.2),
            ["TableStyleMedium19"] = (7, 0.2),
            ["TableStyleMedium20"] = (8, 0.2),
            ["TableStyleMedium21"] = (9, 0.2),
            ["TableStyleMedium23"] = (4, 0.2),
            ["TableStyleMedium24"] = (5, 0.2),
            ["TableStyleMedium25"] = (6, 0.2),
            ["TableStyleMedium26"] = (7, 0.2),
            ["TableStyleMedium27"] = (8, 0.2),
            ["TableStyleMedium28"] = (9, 0.2)
        };

    private static bool TryResolveStripeHex(string themeName, ColorResolver colorResolver, out string hex)
    {
        if (!StripeBandByStyleName.TryGetValue(themeName, out var band) ||
            !colorResolver.TryGetThemeColor(band.ThemeIndex, out var baseColor))
        {
            hex = string.Empty;
            return false;
        }

        hex = ColorHelper.ToHex(ColorHelper.ApplyTint(baseColor, band.Tint));
        return true;
    }
}
