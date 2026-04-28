namespace OysterReport.Internal.OpenXml;

using System.Globalization;

internal static class ExcelFormatCode
{
    private const int MaxSections = 4;

    public static bool IsDateTime(ReadOnlySpan<char> code)
    {
        if (code.IsEmpty || code is "General" || code is "@")
        {
            return false;
        }

        var inQuote = false;
        for (var i = 0; i < code.Length; i++)
        {
            var c = code[i];
            if (c == '\\' && i + 1 < code.Length)
            {
                i++;
                continue;
            }
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '[')
            {
                var rel = code[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                if (IsElapsedDirective(code.Slice(i + 1, rel)))
                {
                    return true;
                }
                i += rel + 1;
                continue;
            }
            if (c is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M')
            {
                return true;
            }
        }
        return false;
    }

    public static string Format(double value, ReadOnlySpan<char> formatCode)
    {
        if (formatCode.IsEmpty || formatCode is "General")
        {
            return value.ToString("G", CultureInfo.InvariantCulture);
        }

        try
        {
            Span<Range> sections = stackalloc Range[MaxSections];
            var sectionCount = SplitSections(formatCode, sections);
            var idx = SelectSectionIndex(sectionCount, value);
            var section = formatCode[sections[idx]];
            if (section.IsEmpty)
            {
                return string.Empty;
            }

            var work = (idx == 1 && sectionCount >= 2) ? Math.Abs(value) : value;

            if (IsDateTimeSection(section))
            {
                return FormatDateTime(work, section);
            }
            return FormatNumeric(work, value, section);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException or OverflowException or IndexOutOfRangeException)
        {
            return value.ToString("G", CultureInfo.InvariantCulture);
        }
    }

    //--------------------------------------------------------------------------------
    // Section split / selection
    //--------------------------------------------------------------------------------

    private static int SelectSectionIndex(int count, double value)
    {
        if (count == 1)
        {
            return 0;
        }
        if (value > 0)
        {
            return 0;
        }
        if (value < 0)
        {
            return count >= 2 ? 1 : 0;
        }
        return count >= 3 ? 2 : 0;
    }

    private static int SplitSections(ReadOnlySpan<char> format, Span<Range> sections)
    {
        var count = 0;
        var sectionStart = 0;
        var inQuote = false;
        var inBracket = false;

        for (var i = 0; i < format.Length; i++)
        {
            var c = format[i];
            if (c == '\\' && i + 1 < format.Length)
            {
                i++;
                continue;
            }
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (!inQuote)
            {
                if (c == '[')
                {
                    inBracket = true;
                }
                else if (c == ']')
                {
                    inBracket = false;
                }
            }
            if (c == ';' && !inQuote && !inBracket && count < sections.Length - 1)
            {
                sections[count++] = sectionStart..i;
                sectionStart = i + 1;
            }
        }
        sections[count++] = sectionStart..format.Length;
        return count;
    }

    //--------------------------------------------------------------------------------
    // Format kind detection
    //--------------------------------------------------------------------------------

    private static bool IsDateTimeSection(ReadOnlySpan<char> section)
    {
        var inQuote = false;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            if (c == '\\' && i + 1 < section.Length)
            {
                i++;
                continue;
            }
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '[')
            {
                var rel = section[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                if (IsElapsedDirective(section.Slice(i + 1, rel)))
                {
                    return true;
                }
                i += rel + 1;
                continue;
            }
            if (c is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M')
            {
                return true;
            }
        }
        return false;
    }

    private static bool IsElapsedDirective(ReadOnlySpan<char> inner)
    {
        if (inner.IsEmpty)
        {
            return false;
        }
        var first = Char.ToLowerInvariant(inner[0]);
        if (first != 'h' && first != 'm' && first != 's')
        {
            return false;
        }
        for (var i = 1; i < inner.Length; i++)
        {
            if (Char.ToLowerInvariant(inner[i]) != first)
            {
                return false;
            }
        }
        return true;
    }

    //--------------------------------------------------------------------------------
    // Numeric
    //--------------------------------------------------------------------------------

    private static string FormatNumeric(double work, double original, ReadOnlySpan<char> body)
    {
        if (HasFractionPattern(body))
        {
            return original.ToString("G", CultureInfo.InvariantCulture);
        }

        Span<char> prefix = stackalloc char[64];
        Span<char> pattern = stackalloc char[32];
        Span<char> suffix = stackalloc char[64];
        var prefixLen = 0;
        var patternLen = 0;
        var suffixLen = 0;

        BuildNumericParts(work, body, prefix, ref prefixLen, pattern, ref patternLen, suffix, ref suffixLen);

        var prefixSpan = prefix[..prefixLen];
        var patternSpan = pattern[..patternLen];
        var suffixSpan = suffix[..suffixLen];

        if (patternSpan.IsEmpty)
        {
            return String.Concat(prefixSpan, suffixSpan);
        }

        Span<char> formatted = stackalloc char[64];
        if (work.TryFormat(formatted, out var written, patternSpan, CultureInfo.InvariantCulture))
        {
            return String.Concat(prefixSpan, formatted[..written], suffixSpan);
        }

        var fallback = work.ToString(patternSpan.ToString(), CultureInfo.InvariantCulture);
        return String.Concat(prefixSpan, fallback.AsSpan(), suffixSpan);
    }

    private static void BuildNumericParts(
        double work,
        ReadOnlySpan<char> body,
        Span<char> prefix,
        ref int prefixLen,
        Span<char> pattern,
        ref int patternLen,
        Span<char> suffix,
        ref int suffixLen)
    {
        var phase = 0;
        var prevPattern = '\0';
        Span<char> scratch = stackalloc char[32];

        for (var i = 0; i < body.Length; i++)
        {
            var c = body[i];

            if (c == '\\' && i + 1 < body.Length)
            {
                AppendLiteral(body[i + 1], prefix, ref prefixLen, suffix, ref suffixLen, ref phase);
                i++;
                continue;
            }
            if (c == '"')
            {
                var rel = body[(i + 1)..].IndexOf('"');
                if (rel < 0)
                {
                    break;
                }
                AppendLiteral(body.Slice(i + 1, rel), prefix, ref prefixLen, suffix, ref suffixLen, ref phase);
                i += rel + 1;
                continue;
            }
            if (c == '[')
            {
                var rel = body[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                i += rel + 1;
                continue;
            }
            if (c == '_' && i + 1 < body.Length)
            {
                AppendLiteral(' ', prefix, ref prefixLen, suffix, ref suffixLen, ref phase);
                i++;
                continue;
            }
            if (c == '*' && i + 1 < body.Length)
            {
                i++;
                continue;
            }
            if (c == '@')
            {
                if (work.TryFormat(scratch, out var written, "G", CultureInfo.InvariantCulture))
                {
                    AppendLiteral(scratch[..written], prefix, ref prefixLen, suffix, ref suffixLen, ref phase);
                }
                continue;
            }

            if (IsNumericPatternChar(c, prevPattern))
            {
                if (phase == 2)
                {
                    suffix[suffixLen++] = c;
                }
                else
                {
                    pattern[patternLen++] = c == '?' ? '#' : c;
                    phase = 1;
                    prevPattern = c;
                }
            }
            else
            {
                if (phase == 0)
                {
                    prefix[prefixLen++] = c;
                }
                else
                {
                    phase = 2;
                    suffix[suffixLen++] = c;
                }
            }
        }
    }

    private static bool IsNumericPatternChar(char c, char prevPattern)
    {
        if (c is '0' or '#' or '?' or '.' or ',' or '%' or 'E' or 'e')
        {
            return true;
        }
        if ((c == '+' || c == '-') && (prevPattern == 'E' || prevPattern == 'e'))
        {
            return true;
        }
        return false;
    }

    private static void AppendLiteral(char c, Span<char> prefix, ref int prefixLen, Span<char> suffix, ref int suffixLen, ref int phase)
    {
        if (phase == 0)
        {
            prefix[prefixLen++] = c;
        }
        else
        {
            phase = 2;
            suffix[suffixLen++] = c;
        }
    }

    private static void AppendLiteral(scoped ReadOnlySpan<char> text, Span<char> prefix, ref int prefixLen, Span<char> suffix, ref int suffixLen, ref int phase)
    {
        if (phase == 0)
        {
            text.CopyTo(prefix[prefixLen..]);
            prefixLen += text.Length;
        }
        else
        {
            phase = 2;
            text.CopyTo(suffix[suffixLen..]);
            suffixLen += text.Length;
        }
    }

    private static bool HasFractionPattern(ReadOnlySpan<char> body)
    {
        var inQuote = false;
        for (var i = 0; i < body.Length; i++)
        {
            var c = body[i];
            if (c == '\\' && i + 1 < body.Length)
            {
                i++;
                continue;
            }
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '[')
            {
                var rel = body[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                i += rel + 1;
                continue;
            }
            if (c == '/' && i > 0 && body[i - 1] is '0' or '#' or '?')
            {
                return true;
            }
        }
        return false;
    }

    //--------------------------------------------------------------------------------
    // Date/Time
    //--------------------------------------------------------------------------------

    private static string FormatDateTime(double value, ReadOnlySpan<char> body)
    {
        DateTime dt;
        try
        {
            dt = DateTime.FromOADate(value);
        }
        catch (ArgumentException)
        {
            return value.ToString("G", CultureInfo.InvariantCulture);
        }

        var hasAmPm = HasAmPmMarker(body);
        Span<char> output = stackalloc char[128];
        var outputLen = 0;

        BuildDateTimeOutput(dt, value, body, hasAmPm, output, ref outputLen);
        return new string(output[..outputLen]);
    }

    private static void BuildDateTimeOutput(
        DateTime dt,
        double value,
        ReadOnlySpan<char> body,
        bool hasAmPm,
        Span<char> output,
        ref int outputLen)
    {
        var culture = CultureInfo.InvariantCulture;
        Span<char> scratch = stackalloc char[32];

        var i = 0;
        while (i < body.Length)
        {
            var c = body[i];

            if (c == '\\' && i + 1 < body.Length)
            {
                output[outputLen++] = body[i + 1];
                i += 2;
                continue;
            }
            if (c == '"')
            {
                var rel = body[(i + 1)..].IndexOf('"');
                if (rel < 0)
                {
                    break;
                }
                var literal = body.Slice(i + 1, rel);
                literal.CopyTo(output[outputLen..]);
                outputLen += literal.Length;
                i += rel + 2;
                continue;
            }
            if (c == '[')
            {
                var rel = body[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                var inner = body.Slice(i + 1, rel);
                if (IsElapsedDirective(inner))
                {
                    AppendElapsed(output, ref outputLen, value, inner, scratch);
                }
                i += rel + 2;
                continue;
            }

            if (TryAmPm(body, i, dt, out var amPmText, out var consumed))
            {
                amPmText.CopyTo(output[outputLen..]);
                outputLen += amPmText.Length;
                i += consumed;
                continue;
            }

            var lower = Char.ToLowerInvariant(c);
            if (lower == 'y')
            {
                var len = CountTokenLen(body, i);
                AppendInt(output, ref outputLen, scratch, len >= 3 ? dt.Year : dt.Year % 100, len >= 3 ? 4 : 2);
                i += len;
                continue;
            }
            if (lower == 'd')
            {
                var len = CountTokenLen(body, i);
                if (len <= 2)
                {
                    AppendInt(output, ref outputLen, scratch, dt.Day, len);
                }
                else
                {
                    var name = len == 3
                        ? culture.DateTimeFormat.GetAbbreviatedDayName(dt.DayOfWeek)
                        : culture.DateTimeFormat.GetDayName(dt.DayOfWeek);
                    name.CopyTo(output[outputLen..]);
                    outputLen += name.Length;
                }
                i += len;
                continue;
            }
            if (lower == 'm')
            {
                var len = CountTokenLen(body, i);
                if (IsMinuteContext(body, i, len))
                {
                    AppendInt(output, ref outputLen, scratch, dt.Minute, len);
                }
                else
                {
                    AppendMonth(output, ref outputLen, scratch, culture, dt.Month, len);
                }
                i += len;
                continue;
            }
            if (lower == 'h')
            {
                var len = CountTokenLen(body, i);
                var hour = hasAmPm ? Convert12Hour(dt.Hour) : dt.Hour;
                AppendInt(output, ref outputLen, scratch, hour, len);
                i += len;
                continue;
            }
            if (lower == 's')
            {
                var len = CountTokenLen(body, i);
                AppendInt(output, ref outputLen, scratch, dt.Second, len);
                i += len;
                if (i < body.Length && body[i] == '.' && i + 1 < body.Length && body[i + 1] == '0')
                {
                    output[outputLen++] = '.';
                    i++;
                    var fracLen = 0;
                    while (i < body.Length && body[i] == '0')
                    {
                        fracLen++;
                        i++;
                    }
                    AppendFractional(output, ref outputLen, scratch, dt.Millisecond, fracLen);
                }
                continue;
            }

            output[outputLen++] = c;
            i++;
        }
    }

    private static void AppendInt(Span<char> output, ref int outputLen, scoped Span<char> scratch, int value, int width)
    {
        ReadOnlySpan<char> format = width <= 1 ? "D1" : width == 2 ? "D2" : width == 3 ? "D3" : "D4";
        if (value.TryFormat(scratch, out var written, format, CultureInfo.InvariantCulture))
        {
            scratch[..written].CopyTo(output[outputLen..]);
            outputLen += written;
        }
    }

    private static void AppendMonth(Span<char> output, ref int outputLen, scoped Span<char> scratch, CultureInfo culture, int month, int len)
    {
        if (len <= 2)
        {
            AppendInt(output, ref outputLen, scratch, month, len);
            return;
        }
        var name = len == 3
            ? culture.DateTimeFormat.GetAbbreviatedMonthName(month)
            : culture.DateTimeFormat.GetMonthName(month);
        if (len >= 5)
        {
            output[outputLen++] = name[0];
        }
        else
        {
            name.CopyTo(output[outputLen..]);
            outputLen += name.Length;
        }
    }

    private static void AppendFractional(Span<char> output, ref int outputLen, scoped Span<char> scratch, int millisecond, int fracLen)
    {
        if (fracLen <= 0)
        {
            return;
        }
        Span<char> formatBuf = stackalloc char[4];
        formatBuf[0] = 'F';
        if (!fracLen.TryFormat(formatBuf[1..], out var fmtLen, default, CultureInfo.InvariantCulture))
        {
            return;
        }
        if ((millisecond / 1000.0).TryFormat(scratch, out var written, formatBuf[..(fmtLen + 1)], CultureInfo.InvariantCulture))
        {
            var slice = scratch[2..written];
            slice.CopyTo(output[outputLen..]);
            outputLen += slice.Length;
        }
    }

    private static int CountTokenLen(ReadOnlySpan<char> s, int start)
    {
        var c = Char.ToLowerInvariant(s[start]);
        var n = 1;
        while (start + n < s.Length && Char.ToLowerInvariant(s[start + n]) == c)
        {
            n++;
        }
        return n;
    }

    private static int Convert12Hour(int hour)
    {
        var h = hour % 12;
        return h == 0 ? 12 : h;
    }

    private static bool HasAmPmMarker(ReadOnlySpan<char> body)
    {
        var inQuote = false;
        for (var i = 0; i < body.Length; i++)
        {
            var c = body[i];
            if (c == '\\' && i + 1 < body.Length)
            {
                i++;
                continue;
            }
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '[')
            {
                var rel = body[(i + 1)..].IndexOf(']');
                if (rel < 0)
                {
                    break;
                }
                i += rel + 1;
                continue;
            }
            if (TryAmPm(body, i, default, out _, out _))
            {
                return true;
            }
        }
        return false;
    }

    private static bool TryAmPm(ReadOnlySpan<char> body, int i, DateTime dt, out string text, out int consumed)
    {
        if (i + 5 <= body.Length)
        {
            var s5 = body.Slice(i, 5);
            if (s5 is "AM/PM")
            {
                text = dt.Hour < 12 ? "AM" : "PM";
                consumed = 5;
                return true;
            }
            if (s5 is "am/pm")
            {
                text = dt.Hour < 12 ? "am" : "pm";
                consumed = 5;
                return true;
            }
        }
        if (i + 3 <= body.Length)
        {
            var s3 = body.Slice(i, 3);
            if (s3 is "A/P")
            {
                text = dt.Hour < 12 ? "A" : "P";
                consumed = 3;
                return true;
            }
            if (s3 is "a/p")
            {
                text = dt.Hour < 12 ? "a" : "p";
                consumed = 3;
                return true;
            }
        }
        text = string.Empty;
        consumed = 0;
        return false;
    }

    private static bool IsMinuteContext(ReadOnlySpan<char> body, int i, int len)
    {
        var p = i - 1;
        while (p >= 0)
        {
            var c = body[p];
            if (c == 'h' || c == 'H')
            {
                return true;
            }
            if (c == ':' || c == '.')
            {
                p--;
                continue;
            }
            break;
        }

        var q = i + len;
        while (q < body.Length)
        {
            var c = body[q];
            if (c == 's' || c == 'S')
            {
                return true;
            }
            if (c == ':' || c == '.')
            {
                q++;
                continue;
            }
            break;
        }
        return false;
    }

    private static void AppendElapsed(Span<char> output, ref int outputLen, double totalDays, ReadOnlySpan<char> inner, scoped Span<char> scratch)
    {
        if (inner.IsEmpty)
        {
            return;
        }
        var first = Char.ToLowerInvariant(inner[0]);
        var len = inner.Length;

        var totalSeconds = (long)Math.Floor(Math.Abs(totalDays) * 86400);

        var v = first switch
        {
            'h' => totalSeconds / 3600,
            'm' => totalSeconds / 60,
            's' => totalSeconds,
            _ => 0L
        };

        if (totalDays < 0)
        {
            output[outputLen++] = '-';
        }

        var width = Math.Max(len, 1);
        if (v.TryFormat(scratch, out var written, default, CultureInfo.InvariantCulture))
        {
            for (var pad = written; pad < width; pad++)
            {
                output[outputLen++] = '0';
            }
            scratch[..written].CopyTo(output[outputLen..]);
            outputLen += written;
        }
    }
}
