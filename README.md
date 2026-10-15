# OysterReport - Excel template to PDF converter

[![NuGet](https://img.shields.io/nuget/v/OysterReport.svg)](https://www.nuget.org/packages/OysterReport/)

## What is this?

A .NET library that converts Excel (.xlsx) templates to PDF.

| Excel |  | PDF |
| :---: | :---: | :---: |
| <img src="Document/excel.png" /> | → | <img src="Document/pdf.png" /> |

## Quick Start

```csharp
var engine = new OysterReportEngine();

using var workbook = new TemplateWorkbook("Invoice.xlsx");
var sheet = workbook.GetSheet("Invoice");

// Replace simple placeholders
sheet.ReplacePlaceholders(new Dictionary<string, string?>
{
    ["CustomerName"] = "UsaUsa Corp",
    ["IssueDate"] = "2025-01-15"
});

// Fill detail rows sequentially from the marker positions
sheet.ReplacePlaceholders(items.Select(static item => new Dictionary<string, string?>
{
    ["ItemName"] = item.Name,
    ["Amount"] = item.Amount.ToString()
}));

using var output = File.Create("invoice.pdf");
engine.GeneratePdf(workbook, output);
```

## Supported features

| Category | Detail |
|---|---|
| **Font** | Size, Bold/Italic/Bold-Italic, Color |
| **Fill** | Background color |
| **Borders** | Border width, Custom color |
| **Text alignment** | Horizontal, Vertical |
| **Merged cells** | Horizontal, Vertical |
| **Images** | Cell-anchored and free-floating |
| **Page setup** | Paper size, Margins |
| **Header / footer** | Header and Footer text |
| **Multi-sheet** | Each sheet |
| **Print area** | Defined print area |
| **Embedded fonts** | Custom font resolver |

## Fonts

Fonts are resolved in this order:

1. `OysterReportEngine.FontResolver` (`IReportFontResolver`) supplies font data from the application.
   For font collections (TTC/OTC), `FontResolveInfo.FaceIndex` selects the face; when it is not set, the face whose name matches `FaceName` is used.
2. `ReportRenderOption.FontDirectories` are searched.
3. OS font directories (Windows, macOS, Linux) are searched when `ReportRenderOption.UseSystemFonts` is `true` (default).

`ReportRenderOption.FontAliases` maps a missing family to substitutes. The default maps Office fonts to metric-compatible fonts (e.g. Calibri → Carlito, Arial → Liberation Sans).
TrueType and OpenType (CFF) fonts are subset when embedded.

Pass the same options to `TemplateWorkbook` so that column widths are measured with the same fonts:

```csharp
var options = new ReportRenderOption { FontDirectories = ["/app/fonts"] };
using var workbook = new TemplateWorkbook("Invoice.xlsx", options);
var engine = new OysterReportEngine { RenderingOptions = options };
```

## Images

JPEG, PNG, BMP and GIF are embedded directly. Other formats (e.g. WebP) can be converted to PNG or JPEG by `ReportRenderOption.ImageConverter`; otherwise the image is skipped with a warning.

## Warnings

`ReportRenderOption.OnRenderWarning` receives non-fatal issues:

| Kind | Description |
|---|---|
| `ImageDecodeFailed` | The image could not be decoded and was skipped |
| `MissingGlyph` | Characters not contained in the font (drawn as the font's missing glyph) |
| `FontEmbeddingRestricted` | The font license restricts embedding (the font is embedded anyway) |

## Dependencies

- DocumentFormat.OpenXml

PDF generation (font subsetting, images, text layout) is implemented in managed code without native libraries.
Some parts are ported from MIT-licensed projects; see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
