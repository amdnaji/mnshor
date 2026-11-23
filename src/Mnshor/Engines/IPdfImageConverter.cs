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

    Task<ImageToPdfInputItem?> InspectImageAsync(string fileName, string contentType, byte[] data);

    Task<ExecutionResult> ConvertImagesToPdfAsync(
        IReadOnlyList<ImageToPdfInputItem> images,
        ImagesToPdfOptions? options = null,
        IProgress<(int Current, int Total)>? progress = null);
}
