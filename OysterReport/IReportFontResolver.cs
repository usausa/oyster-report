namespace OysterReport;

public interface IReportFontResolver
{
    FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic);

    ReadOnlyMemory<byte>? GetFont(string faceName) => null;
}
