namespace OysterReport;

// Sheet
public class SheetInfo
{
    public string SheetName { get; set; } = string.Empty;

    public string UsedRange { get; set; } = string.Empty;

    //public List<CellInfo> Cells { get; set; } = new();
}

// Cell
public sealed class CellInfo
{
    public int Row { get; set; }

    public int Column { get; set; }

    public string CellAddress { get; set; } = string.Empty;

    public string Value { get; set; } = string.Empty;

    public double Width { get; set; }

    public double Height { get; set; }

    // TODO
}
