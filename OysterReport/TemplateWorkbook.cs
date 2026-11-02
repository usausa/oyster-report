namespace OysterReport;

using OysterReport.Internal;
using OysterReport.Internal.OpenXml;

public sealed class TemplateWorkbook : IDisposable
{
    private readonly List<TemplateSheet> sheets;

    internal ReportWorkbook ReportWorkbook { get; }

    public IReadOnlyList<TemplateSheet> Sheets => sheets;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    public TemplateWorkbook(string filePath, ReportRenderOption? renderOption = null)
        : this(OpenXmlLoader.Load(filePath, renderOption))
    {
    }

    public TemplateWorkbook(Stream stream, ReportRenderOption? renderOption = null)
        : this(OpenXmlLoader.Load(stream, renderOption))
    {
    }

    private TemplateWorkbook(ReportWorkbook reportWorkbook)
    {
        ReportWorkbook = reportWorkbook;
        sheets = ReportWorkbook.Sheets.Select(x => new TemplateSheet(ReportWorkbook, x)).ToList();
    }

    public void Dispose()
    {
        // No unmanaged resources held; keeps IDisposable for API compatibility.
    }

    //--------------------------------------------------------------------------------
    // Copy
    //--------------------------------------------------------------------------------

    public TemplateWorkbook Clone() => new(ReportWorkbook.DeepClone());

    //--------------------------------------------------------------------------------
    // Sheet
    //--------------------------------------------------------------------------------

    public TemplateSheet GetSheet(string name)
    {
        return sheets.FirstOrDefault(x => String.Equals(x.Name, name, StringComparison.Ordinal)) ??
               throw new ArgumentException($"Sheet not found. name=[{name}]", nameof(name));
    }

    public TemplateSheet GetSheet(int index) => sheets[index];

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public TemplateSheet CopySheet(string sourceSheetName, string newSheetName)
    {
        var sourceSheet = GetSheet(sourceSheetName);
        var cloned = sourceSheet.UnderlyingSheet.Clone(newSheetName);
        ReportWorkbook.AddSheet(cloned);
        var newSheet = new TemplateSheet(ReportWorkbook, cloned);
        sheets.Add(newSheet);
        return newSheet;
    }

    public void RemoveSheet(string name)
    {
        var sheet = GetSheet(name);
        ReportWorkbook.RemoveSheet(sheet.UnderlyingSheet);
        sheets.Remove(sheet);
    }

    public int ReplacePlaceholder(string marker, string value)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.ReplacePlaceholder(marker, value);
        }
        return count;
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.ReplacePlaceholders(values);
        }
        return count;
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    public int ReplacePlaceholderWithImage(string marker, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit = ReportImageFit.Contain)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.ReplacePlaceholderWithImage(marker, imageBytes, fit);
        }
        return count;
    }

    public int ReplaceImage(string name, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit = ReportImageFit.Contain)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.ReplaceImage(name, imageBytes, fit);
        }
        return count;
    }

    public int RemoveImage(string name)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.RemoveImage(name);
        }
        return count;
    }
}
