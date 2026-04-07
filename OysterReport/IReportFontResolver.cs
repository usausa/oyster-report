namespace OysterReport;

public interface IReportFontResolver
{
    ReportFontResolveResult Resolve(ReportFontRequest request);
}
