namespace OysterReport.Generator.Models;

internal sealed record RowExpansionRequest
{
    public int TemplateStartRowIndex { get; init; } // Start row index of the template block

    public int TemplateEndRowIndex { get; init; } // End row index of the template block

    public int RepeatCount { get; init; } // Number of additional repetitions to insert

    public IReadOnlyList<IReadOnlyDictionary<string, string?>> PlaceholderValuesByIteration { get; init; } =
        Array.Empty<IReadOnlyDictionary<string, string?>>(); // Placeholder values applied to each repeated row

    internal int GetRepeatCount()
    {
        if (RepeatCount > 0)
        {
            return RepeatCount;
        }

        if (PlaceholderValuesByIteration.Count > 0)
        {
            return PlaceholderValuesByIteration.Count;
        }

        throw new InvalidOperationException("RepeatCount or PlaceholderValuesByIteration must be specified.");
    }

    internal IReadOnlyDictionary<string, string?> GetPlaceholderValues(int iteration) =>
        iteration < PlaceholderValuesByIteration.Count
            ? PlaceholderValuesByIteration[iteration]
            : new Dictionary<string, string?>();
}
