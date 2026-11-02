namespace OysterReport.Internal;

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;

internal static class PlaceholderReplacer
{
    public static int ReplaceInRows(ReportSheet sheet, int startRow, int endRow, IReadOnlyDictionary<string, string?> values)
    {
        if (values.Count == 0)
        {
            return 0;
        }

        var lookup = EnsureOrdinalLookup(values);
        var replacedKeys = new List<string>();
        var count = 0;
        foreach (var cell in sheet.CellsInStorageOrder)
        {
            if ((cell.Row < startRow) || (cell.Row > endRow))
            {
                continue;
            }

            var text = cell.DisplayText;
            if (!text.Contains("{{", StringComparison.Ordinal))
            {
                continue;
            }

            replacedKeys.Clear();
            var replaced = ReplaceText(text, lookup, replacedKeys);
            if (replacedKeys.Count > 0)
            {
                TemplateSheet.SetCellText(cell, replaced);
                count += replacedKeys.Count;
            }
        }

        return count;
    }

    public static int ReplaceWithImageInRows(ReportSheet sheet, int startRow, int endRow, string marker, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit)
    {
        var placeholder = "{{" + marker + "}}";
        var count = 0;
        foreach (var cell in sheet.CellsInStorageOrder)
        {
            if ((cell.Row < startRow) || (cell.Row > endRow))
            {
                continue;
            }

            var text = cell.DisplayText;
            if (!text.Contains(placeholder, StringComparison.Ordinal))
            {
                continue;
            }

            TemplateSheet.SetCellText(cell, text.Replace(placeholder, string.Empty, StringComparison.Ordinal));
            TemplateSheet.AddCellImage(sheet, cell.Row, cell.Column, cell.Merge, marker, imageBytes, fit);
            count++;
        }

        return count;
    }

    public static int ReplaceInHeaderFooter(ReportSheet sheet, string marker, string value)
    {
        if (!ContainsPlaceholder(sheet.HeaderFooter))
        {
            return 0;
        }

        return ReplaceInHeaderFooter(sheet, new Dictionary<string, string?>(StringComparer.Ordinal) { [marker] = value });
    }

    public static int ReplaceInHeaderFooter(ReportSheet sheet, IReadOnlyDictionary<string, string?> values)
    {
        var headerFooter = sheet.HeaderFooter;
        if ((values.Count == 0) || !ContainsPlaceholder(headerFooter))
        {
            return 0;
        }

        var escaped = new Dictionary<string, string?>(values.Count, StringComparer.Ordinal);
        foreach (var (key, value) in values)
        {
            escaped[key] = value?.Replace("&", "&&", StringComparison.Ordinal);
        }

        var replacedKeys = new List<string>();
        var count = 0;
        sheet.HeaderFooter = headerFooter with
        {
            OddHeader = Replace(headerFooter.OddHeader),
            OddFooter = Replace(headerFooter.OddFooter),
            EvenHeader = Replace(headerFooter.EvenHeader),
            EvenFooter = Replace(headerFooter.EvenFooter),
            FirstHeader = Replace(headerFooter.FirstHeader),
            FirstFooter = Replace(headerFooter.FirstFooter)
        };

        return count;

        string? Replace(string? text)
        {
            if (!ContainsPlaceholder(text))
            {
                return text;
            }

            replacedKeys.Clear();
            var replaced = ReplaceText(text, escaped, replacedKeys);
            count += replacedKeys.Count;
            return replaced;
        }
    }

    private static bool ContainsPlaceholder(ReportHeaderFooter headerFooter) =>
        ContainsPlaceholder(headerFooter.OddHeader) || ContainsPlaceholder(headerFooter.OddFooter) ||
        ContainsPlaceholder(headerFooter.EvenHeader) || ContainsPlaceholder(headerFooter.EvenFooter) ||
        ContainsPlaceholder(headerFooter.FirstHeader) || ContainsPlaceholder(headerFooter.FirstFooter);

    private static bool ContainsPlaceholder([NotNullWhen(true)] string? text) => (text is not null) && text.Contains("{{", StringComparison.Ordinal);

    // Keeps the ordinal marker comparison used by the single-key replacement API
    private static Dictionary<string, string?> EnsureOrdinalLookup(IReadOnlyDictionary<string, string?> values)
    {
        if ((values is Dictionary<string, string?> dictionary) &&
            (ReferenceEquals(dictionary.Comparer, EqualityComparer<string>.Default) ||
             ReferenceEquals(dictionary.Comparer, StringComparer.Ordinal)))
        {
            return dictionary;
        }

        var copy = new Dictionary<string, string?>(values.Count, StringComparer.Ordinal);
        foreach (var (key, value) in values)
        {
            copy[key] = value;
        }

        return copy;
    }

    [SkipLocalsInit]
    private static string ReplaceText(string text, Dictionary<string, string?> values, List<string> replacedKeys)
    {
#if NET9_0_OR_GREATER
        var lookup = values.GetAlternateLookup<ReadOnlySpan<char>>();
#endif
        var builder = new DefaultInterpolatedStringHandler(0, 0, CultureInfo.InvariantCulture, stackalloc char[256]);
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

            var name = text.AsSpan(open + 2, close - open - 2);
#if NET9_0_OR_GREATER
            if (lookup.TryGetValue(name, out var key, out var value))
#else
            var key = name.ToString();
            if (values.TryGetValue(key, out var value))
#endif
            {
                builder.AppendFormatted(text.AsSpan(copiedUpTo, open - copiedUpTo));
                builder.AppendFormatted(value);
                copiedUpTo = close + 2;
                searchPos = close + 2;
                AddDistinct(replacedKeys, key);
            }
            else
            {
                // Unknown marker: keep it and continue right after the opening braces
                // so overlapping candidates (e.g. "{{A{{B}}") are still found
                searchPos = open + 2;
            }
        }

        if (replacedKeys.Count == 0)
        {
            builder.ToStringAndClear();
            return text;
        }

        builder.AppendFormatted(text.AsSpan(copiedUpTo));
        return builder.ToStringAndClear();
    }

    private static void AddDistinct(List<string> keys, string key)
    {
        foreach (var existing in keys)
        {
            if (String.Equals(existing, key, StringComparison.Ordinal))
            {
                return;
            }
        }

        keys.Add(key);
    }
}
