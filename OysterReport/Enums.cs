namespace OysterReport;

public enum ReportDumpFormat
{
    Json,
    Markdown
}

public enum ReportDiagnosticSeverity
{
    Info,
    Warning,
    Error
}

public enum ReportPaperSize
{
    Custom,
    A4,
    Letter,
    Legal
}

public enum ReportPageOrientation
{
    Portrait,
    Landscape
}

public enum ReportAnchorType
{
    MoveAndSizeWithCells,
    MoveWithCells,
    Absolute
}

public enum ReportHorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Justify
}

public enum ReportVerticalAlignment
{
    Top,
    Center,
    Bottom,
    Justify
}

public enum ReportBorderStyle
{
    None,
    Thin,
    Medium,
    Thick,
    DoubleLine,
    Dashed,
    Dotted,
    Hair,
    DashDot
}

public enum ReportCellValueKind
{
    Blank,
    Text,
    Number,
    DateTime,
    Boolean,
    Error
}
