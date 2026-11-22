using Mnshor.Models;

namespace Mnshor.Engines;

public interface IPdfImageConverter
{
    Task<IReadOnlyList<ExtractedPdfImage>> ExtractEmbeddedImagesAsync(
        byte[] pdfData,
        ImageOutputFormat targetFormat = ImageOutputFormat.Png,
        IProgress<(int Current, int Total)>? progress = null);

    Task<ExecutionResult> PackageImagesAsync(
        string originalFileName,
        IReadOnlyList<RenderedPdfPageImage> images,
        ImageOutputFormat format);
}
