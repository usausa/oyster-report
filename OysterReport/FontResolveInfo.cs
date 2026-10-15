namespace OysterReport;

public sealed record FontResolveInfo
{
    public string FaceName { get; }

    public bool MustSimulateBold { get; set; }

    public bool MustSimulateItalic { get; set; }

    // Face index in a font collection (TTC/OTC); when null the face whose name matches FaceName is used (or the first face)
    public int? FaceIndex { get; set; }

    public FontResolveInfo(string faceName)
    {
        FaceName = faceName;
    }
}
