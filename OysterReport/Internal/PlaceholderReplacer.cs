namespace OysterReport.Internal;

using System.Text;

// Replaces {{key}} placeholders for multiple keys in a single pass over the cells.
// The returned count matches the per-key API: one per (cell, key) pair with at least one occurrence.
internal static class PlaceholderReplacer
{
    public static int ReplaceInRows(ReportSheet sheet, int startRow, int endRow, IReadOnlyDictionary<string, string?> values)
    {
        if (values.Count == 0)
        {
            return 0;
        }

        var lookup = EnsureOrdinalLookup(values);
        var count = 0;
        foreach (var cell in sheet.Cells)
        {
            if ((cell.Row < startRow) || (cell.Row > endRow))
            {
                continue;
            }

            var replaced = ReplaceText(cell.DisplayText, lookup, out var replacedKeys);
            if (replacedKeys > 0)
            {
                TemplateSheet.SetCellText(cell, replaced);
                count += replacedKeys;
            }
        }

        return count;
    }

    // Keeps the ordinal marker comparison used by the single-key replacement API
    private static IReadOnlyDictionary<string, string?> EnsureOrdinalLookup(IReadOnlyDictionary<string, string?> values)
    {
        if (values is Dictionary<string, string?> dictionary &&
            (ReferenceEquals(dictionary.Comparer, EqualityComparer<string>.Default) ||
             ReferenceEquals(dictionary.Comparer, StringComparer.Ordinal)))
        {
            return values;
        }

        var copy = new Dictionary<string, string?>(values.Count, StringComparer.Ordinal);
        foreach (var (key, value) in values)
        {
            copy[key] = value;
        }

        return copy;
    }

    private static string ReplaceText(string text, IReadOnlyDictionary<string, string?> values, out int replacedKeys)
    {
        replacedKeys = 0;

        StringBuilder? builder = null;
        HashSet<string>? keys = null;
        var copiedUpTo = 0;
        var searchPos = 0;

        while (true)
        {
            var open = text.IndexOf("{{", searchPos, StringComparison.Ordinal);
            if (open < 0)
            {
                break;
            }

            var close = text.IndexOf("}}", open + 2, StringComparison.Ordinal);
            if (close < 0)
            {
                break;
            }

            var key = text[(open + 2)..close];
            if (values.TryGetValue(key, out var value))
            {
                builder ??= new StringBuilder(text.Length);
                builder.Append(text, copiedUpTo, open - copiedUpTo);
                builder.Append(value ?? string.Empty);
                copiedUpTo = close + 2;
                searchPos = close + 2;

                keys ??= [];
                if (keys.Add(key))
                {
                    replacedKeys++;
                }
            }
            else
            {
                // Unknown marker: keep it and continue right after the opening braces
                // so overlapping candidates (e.g. "{{A{{B}}") are still found
                searchPos = open + 2;
            }
        }

        if (builder is null)
        {
            return text;
        }

        builder.Append(text, copiedUpTo, text.Length - copiedUpTo);
        return builder.ToString();
    }
}
