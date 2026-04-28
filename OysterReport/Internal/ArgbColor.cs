namespace OysterReport.Internal;

internal readonly struct ArgbColor : IEquatable<ArgbColor>
{
    public static readonly ArgbColor Black = new(0xFF, 0x00, 0x00, 0x00);
    public static readonly ArgbColor White = new(0xFF, 0xFF, 0xFF, 0xFF);

    public uint Value { get; }

    public ArgbColor(uint argb)
    {
        Value = argb;
    }

    public ArgbColor(byte a, byte r, byte g, byte b)
    {
        Value = ((uint)a << 24) | ((uint)r << 16) | ((uint)g << 8) | b;
    }

    public byte A => (byte)(Value >> 24);

    public byte R => (byte)(Value >> 16);

    public byte G => (byte)(Value >> 8);

    public byte B => (byte)Value;

    public bool Equals(ArgbColor other) => Value == other.Value;

    public override bool Equals(object? obj) => obj is ArgbColor other && Equals(other);

    public override int GetHashCode() => Value.GetHashCode();

    public static bool operator ==(ArgbColor left, ArgbColor right) => left.Value == right.Value;

    public static bool operator !=(ArgbColor left, ArgbColor right) => left.Value != right.Value;
}
