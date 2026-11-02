namespace OysterReport.Internal.Pdf;

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

internal sealed class PdfImageCache
{
    private readonly Dictionary<int, List<(ReadOnlyMemory<byte> Source, PdfImage Image)>> imagesByLength = [];

    public bool TryGet(ReadOnlyMemory<byte> source, out PdfImage image)
    {
        if (imagesByLength.TryGetValue(source.Length, out var candidates))
        {
            var span = source.Span;
            foreach (var (candidate, candidateImage) in candidates)
            {
                var candidateSpan = candidate.Span;
                if (Unsafe.AreSame(ref MemoryMarshal.GetReference(candidateSpan), ref MemoryMarshal.GetReference(span)) ||
                    candidateSpan.SequenceEqual(span))
                {
                    image = candidateImage;
                    return true;
                }
            }
        }

        image = default;
        return false;
    }

    public void Add(ReadOnlyMemory<byte> source, PdfImage image)
    {
        ref var candidates = ref CollectionsMarshal.GetValueRefOrAddDefault(imagesByLength, source.Length, out _);
        candidates ??= [];
        candidates.Add((source, image));
    }
}
