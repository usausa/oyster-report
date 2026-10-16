namespace OysterReport.Internal.Pdf;

using System.Buffers;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;

internal sealed class PdfByteBuffer : IBufferWriter<byte>, IDisposable
{
    private const int MaxNumberLength = 32;

    private byte[] buffer;

    public PdfByteBuffer(int initialCapacity = 4096)
    {
        buffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
    }

    public int Length { get; private set; }

    public ReadOnlySpan<byte> WrittenSpan => buffer.AsSpan(0, Length);

    public void Dispose()
    {
        var array = buffer;
        buffer = [];
        Length = 0;
        if (array.Length > 0)
        {
            ArrayPool<byte>.Shared.Return(array);
        }
    }

    public void Clear() => Length = 0;

    public void Advance(int count) => Length += count;

    public Memory<byte> GetMemory(int sizeHint = 0)
    {
        EnsureCapacity(sizeHint);
        return buffer.AsMemory(Length);
    }

    public Span<byte> GetSpan(int sizeHint = 0)
    {
        EnsureCapacity(sizeHint);
        return buffer.AsSpan(Length);
    }

    public void Write(byte value)
    {
        if (Length >= buffer.Length)
        {
            Grow(1);
        }

        buffer[Length++] = value;
    }

    public void Write(ReadOnlySpan<byte> value)
    {
        EnsureCapacity(value.Length);
        value.CopyTo(buffer.AsSpan(Length));
        Length += value.Length;
    }

    public void WriteAscii(ReadOnlySpan<char> value)
    {
        EnsureCapacity(value.Length);
        var destination = buffer.AsSpan(Length, value.Length);
        if (Ascii.FromUtf16(value, destination, out _) != OperationStatus.Done)
        {
            Encoding.ASCII.GetBytes(value, destination);
        }

        Length += value.Length;
    }

    public void WriteLatin1(ReadOnlySpan<char> value)
    {
        EnsureCapacity(value.Length);
        Length += Encoding.Latin1.GetBytes(value, buffer.AsSpan(Length));
    }

    public void WriteInteger(long value)
    {
        EnsureCapacity(MaxNumberLength);
        value.TryFormat(buffer.AsSpan(Length), out var written, default, CultureInfo.InvariantCulture);
        Length += written;
    }

    public void WriteNumber(double value)
    {
        EnsureCapacity(MaxNumberLength);
        if (PdfNumberFormatter.TryFormat(value, buffer.AsSpan(Length), out var written))
        {
            Length += written;
            return;
        }

        WriteNumberSlow(value);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void WriteNumberSlow(double value)
    {
        int written;
        while (!value.TryFormat(buffer.AsSpan(Length), out written, "0.####", CultureInfo.InvariantCulture))
        {
            Grow(buffer.Length - Length + 1);
        }

        Length += written;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureCapacity(int sizeHint)
    {
        if (buffer.Length - Length < Math.Max(sizeHint, 1))
        {
            Grow(sizeHint);
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void Grow(int sizeHint)
    {
        var required = (long)Length + Math.Max(sizeHint, 1);
        var newSize = (int)Math.Min(Math.Max(required, (long)buffer.Length * 2), Array.MaxLength);
        var newBuffer = ArrayPool<byte>.Shared.Rent(newSize);
        buffer.AsSpan(0, Length).CopyTo(newBuffer);
        if (buffer.Length > 0)
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }

        buffer = newBuffer;
    }
}

internal static class PdfNumberFormatter
{
    private const double Scale = 10000d;

    private const double FastPathLimit = 100000d;

    private const double TieMargin = 1e-5d;

    public static bool TryFormat(double value, Span<byte> destination, out int written)
    {
        if ((value == 0d) && !Double.IsNegative(value))
        {
            destination[0] = (byte)'0';
            written = 1;
            return true;
        }

        var absolute = Math.Abs(value);
        if (!(absolute < FastPathLimit))
        {
            written = 0;
            return false;
        }

        var scaled = absolute * Scale;
        var floor = Math.Floor(scaled);
        var fraction = scaled - floor;
        if (Math.Abs(fraction - 0.5d) < TieMargin)
        {
            written = 0;
            return false;
        }

        var units = (long)floor + (fraction > 0.5d ? 1 : 0);
        if (units == 0)
        {
            // "-0" handling is left to the BCL
            written = 0;
            return false;
        }

        var position = 0;
        if (value < 0d)
        {
            destination[position++] = (byte)'-';
        }

        var integer = units / (long)Scale;
        var decimals = (int)(units - (integer * (long)Scale));
        integer.TryFormat(destination[position..], out var count, default, CultureInfo.InvariantCulture);
        position += count;
        if (decimals != 0)
        {
            destination[position++] = (byte)'.';
            decimals.TryFormat(destination[position..], out count, "D4", CultureInfo.InvariantCulture);
            position += count;
            while (destination[position - 1] == (byte)'0')
            {
                position--;
            }
        }

        written = position;
        return true;
    }
}

internal sealed class PdfByteBufferStream(PdfByteBuffer target) : Stream
{
    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => target.Length;

    public override long Position
    {
        get => target.Length;
        set => throw new NotSupportedException();
    }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => target.Write(buffer.AsSpan(offset, count));

    public override void Write(ReadOnlySpan<byte> buffer) => target.Write(buffer);

    public override void WriteByte(byte value) => target.Write(value);
}
