# OysterReport

A .NET library that converts Excel (.xlsx) templates to PDF.  

| Excel |  | PDF |
| :---: | :---: | :---: |
| <img src="Document/excel.png" /> | → | <img src="Document/pdf.png" /> |

# Quick Start

```csharp
var engine = new OysterReportEngine();

using var workbook = new TemplateWorkbook("Invoice.xlsx");
var sheet = workbook.GetSheet("Invoice");

// Replace simple placeholders
sheet.ReplacePlaceholder("CustomerName", "Acme Corp");
sheet.ReplacePlaceholder("IssueDate", "2025-01-15");

// Expand a detail row
var templateRow = sheet.FindRow("ItemName");
var row = templateRow;
foreach (var item in items)
{
    row = templateRow.InsertCopyAfter(row);
    row.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["ItemName"] = item.Name,
        ["Amount"]   = item.Amount.ToString()
    });
}
templateRow.Delete();

using var output = File.Create("invoice.pdf");
engine.GeneratePdf(workbook, output);
```

# Supported features

| Category | Detail |
|---|---|
| **Font** | Size, bold, italic, bold-italic, color (RGB / theme) |
| **Fill** | Background color (RGB / theme) |
| **Borders** | Thin, medium, thick, double, dashed, dotted; custom color |
| **Text alignment** | Horizontal (left / center / right), vertical (top / middle / bottom), wrap |
| **Merged cells** | Horizontal, vertical, rectangular; no duplicate text rendering |
| **Hidden rows / columns** | Hidden rows and columns are excluded from output |
| **Images** | Cell-anchored and free-floating (PNG / JPEG) |
| **Page setup** | Paper size (A4 portrait / landscape, etc.), margins, `CenterHorizontally`, page breaks |
| **Header / footer** | Left / center / right header and footer text |
| **Multi-sheet** | Each sheet → one PDF page; full-workbook or single-sheet output |
| **Print area** | Only cells within the defined print area are rendered |
| **Embedded fonts** | Custom `IReportFontResolver` with TTF byte data; bold/italic flag forwarding |

# Requirements

- .NET 10 or later
- Dependencies:
  - [ClosedXML](https://github.com/ClosedXML/ClosedXML)
  - [PDFsharp](https://github.com/empira/PDFsharp)
  - [SkiaSharp](https://github.com/mono/SkiaSharp)
