namespace OysterReport.Generator.Models;

internal sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];

    public ReportWorkbook(ReportMetadata? metadata = null, ReportMeasurementProfile? measurementProfile = null)
    {
        Metadata = metadata ?? new ReportMetadata();
        MeasurementProfile = measurementProfile ?? new ReportMeasurementProfile();
    }

    public IReadOnlyList<ReportSheet> Sheets => sheets;

    public ReportMetadata Metadata { get; }

    public ReportMeasurementProfile MeasurementProfile { get; }

    internal void AddSheet(ReportSheet sheet)
    {
        sheets.Add(sheet);
    }
}
