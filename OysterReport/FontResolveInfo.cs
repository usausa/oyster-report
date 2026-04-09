namespace OysterReport;

public sealed record FontResolveInfo
{
    public string FaceName { get; }

    public bool MustSimulateBold { get; set; }

    public bool MustSimulateItalic { get; set; }

    public FontResolveInfo(string faceName)
    {
        FaceName = faceName;
    }
}
