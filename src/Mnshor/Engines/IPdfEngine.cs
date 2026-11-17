using Mnshor.Models;

namespace Mnshor.Engines;

public interface IPdfEngine
{
    int GetPageCount(byte[] pdfData);
    Task<ExecutionResult> MergeAsync(IReadOnlyList<DocumentItem> documents);
    Task<ExecutionResult> SplitRangeAsync(DocumentItem document, string rangeExpression);
    Task<ExecutionResult> BurstToZipAsync(DocumentItem document);
    Task<ExecutionResult> RotateAsync(DocumentItem document, int angleDegrees, string pageScope = "all", string? customRange = null);
    Task<ExecutionResult> RotatePagesAsync(DocumentItem document, IReadOnlyList<int> pageRotations);
    Task<ExecutionResult> RemovePagesAsync(DocumentItem document, IReadOnlyCollection<int> pagesToRemove);
    Task<ExecutionResult> RemovePagesAsync(DocumentItem document, string rangeExpression);
    Task<ExecutionResult> ReorderPagesAsync(DocumentItem document, IReadOnlyList<int> pageOrder);
    Task<ExecutionResult> ReversePagesAsync(DocumentItem document);
    PdfMetadataModel GetMetadata(byte[] pdfData);
    Task<ExecutionResult> UpdateMetadataAsync(DocumentItem document, PdfMetadataModel metadata, bool sanitizeAll = false);
    Task<ExecutionResult> WatermarkAsync(DocumentItem document, PdfWatermarkOptions options);
    Task<ExecutionResult> SignPdfAsync(DocumentItem document, PdfSignOptions options);
}
