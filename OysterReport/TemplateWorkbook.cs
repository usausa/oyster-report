namespace OysterReport;

using ClosedXML.Excel;

public sealed class TemplateWorkbook : IDisposable
{
    private readonly XLWorkbook workbook;

    private readonly List<TemplateSheet> sheets;

    internal IXLWorkbook UnderlyingWorkbook => workbook;

    public IReadOnlyList<TemplateSheet> Sheets => sheets;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    public TemplateWorkbook(string filePath)
    {
        workbook = new XLWorkbook(filePath);
        sheets = workbook.Worksheets.Select(static x => new TemplateSheet(x)).ToList();
    }

    public TemplateWorkbook(Stream stream)
    {
        workbook = new XLWorkbook(stream);
        sheets = workbook.Worksheets.Select(static x => new TemplateSheet(x)).ToList();
    }

    public void Dispose() => workbook.Dispose();

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
        var sourceWorksheet = workbook.Worksheet(sourceSheetName);
        var newWorksheet = sourceWorksheet.CopyTo(newSheetName);
        var newSheet = new TemplateSheet(newWorksheet);
        sheets.Add(newSheet);
        return newSheet;
    }

    public void RemoveSheet(string name)
    {
        var sheet = GetSheet(name);
        workbook.Worksheet(name).Delete();
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
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }
}
