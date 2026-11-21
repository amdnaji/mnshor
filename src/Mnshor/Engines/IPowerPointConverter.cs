using Mnshor.Models;

namespace Mnshor.Engines;

public interface IPowerPointConverter
{
    PptxPresentationInfo InspectPresentation(byte[] pptxData, IProgress<(int Current, int Total)>? progress = null);

    Task<ExecutionResult> ConvertPptxToPdfAsync(
        DocumentItem document,
        PowerPointPdfOptions? options = null,
        IProgress<(int Current, int Total, string Stage)>? progress = null);

    Task<ExecutionResult> ConvertSlidesToPptxAsync(
        string originalFileName,
        IReadOnlyList<PdfPageSlideImage> slides,
        PowerPointPdfOptions? options = null,
        IProgress<(int Current, int Total, string Stage)>? progress = null);
}
