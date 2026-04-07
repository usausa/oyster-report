namespace OysterReport.Generator.Models;

internal sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];
    private readonly List<ReportDiagnostic> diagnostics = [];

    public ReportWorkbook(ReportMetadata? metadata = null, ReportMeasurementProfile? measurementProfile = null)
    {
        Metadata = metadata ?? new ReportMetadata();
        MeasurementProfile = measurementProfile ?? new ReportMeasurementProfile();
    }

    public IReadOnlyList<ReportSheet> Sheets => sheets; // List of sheets in the workbook

    public ReportMetadata Metadata { get; } // Metadata for the entire report workbook

    public ReportMeasurementProfile MeasurementProfile { get; } // Measurement settings and environment normalization profile

    public IReadOnlyList<ReportDiagnostic> Diagnostics => diagnostics; // Diagnostics collected during reading

    public ReportSheet AddSheet(string name)
    {
        var sheet = new ReportSheet(name);
        AddSheet(sheet);
        return sheet;
    }

    public void AddSheet(ReportSheet sheet)
    {
        sheets.Add(sheet);
    }

    internal void AddDiagnostic(ReportDiagnostic diagnostic)
    {
        diagnostics.Add(diagnostic);
    }
}
