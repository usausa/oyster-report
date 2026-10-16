namespace OysterReport.Internal.Pdf;

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

internal sealed class PdfImageCache
{
    private readonly Dictionary<int, List<(ReadOnlyMemory<byte> Source, int Id)>> imagesByLength = [];

    public bool TryGet(ReadOnlyMemory<byte> source, out int id)
    {
        if (imagesByLength.TryGetValue(source.Length, out var candidates))
        {
            var span = source.Span;
            foreach (var (candidate, candidateId) in candidates)
            {
                var candidateSpan = candidate.Span;
                if (Unsafe.AreSame(ref MemoryMarshal.GetReference(candidateSpan), ref MemoryMarshal.GetReference(span)) ||
                    candidateSpan.SequenceEqual(span))
                {
                    id = candidateId;
                    return true;
                }
            }
        }

        id = 0;
        return false;
    }

    public void Add(ReadOnlyMemory<byte> source, int id)
    {
        ref var candidates = ref CollectionsMarshal.GetValueRefOrAddDefault(imagesByLength, source.Length, out _);
        candidates ??= [];
        candidates.Add((source, id));
    }
}
