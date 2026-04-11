namespace OysterReport.Internal;

using System.Buffers;

public ref struct ValueStringBuilder
{
    private char[]? arrayFromPool;
    private Span<char> span;
    private int pos;

    public ValueStringBuilder(Span<char> initialBuffer)
    {
        arrayFromPool = null;
        span = initialBuffer;
        pos = 0;
    }

    public void Dispose()
    {
        if (arrayFromPool != null)
        {
            ArrayPool<char>.Shared.Return(arrayFromPool);
            arrayFromPool = null;
        }
    }

    public void Append(string str)
    {
        if (str.Length > span.Length - pos)
        {
            Grow(str.Length);
        }
        str.AsSpan().CopyTo(span[pos..]);
        pos += str.Length;
    }

    public void Append(ReadOnlySpan<char> str)
    {
        if (str.Length > span.Length - pos)
        {
            Grow(str.Length);
        }
        str.CopyTo(span[pos..]);
        pos += str.Length;
    }

    public void Append(char c)
    {
        if (pos >= span.Length)
        {
            Grow(1);
        }
        span[pos++] = c;
    }

    private void Grow(int additional)
    {
        var newSize = Math.Max(span.Length * 2, pos + additional);
        var newArray = ArrayPool<char>.Shared.Rent(newSize);
        span.CopyTo(newArray);

        if (arrayFromPool != null)
        {
            ArrayPool<char>.Shared.Return(arrayFromPool);
        }

        arrayFromPool = newArray;
        span = arrayFromPool.AsSpan(0, newSize);
    }

    public override readonly string ToString()
    {
        return new string(span[..pos].ToArray());
    }
}
