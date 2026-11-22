namespace Mnshor.Models;

public enum PdfImageConversionMode
{
    PagesToImages,
    ExtractEmbeddedImages
}

public enum ImageOutputFormat
{
    Jpg,
    Png,
    WebP
}

public enum ImageRenderQuality
{
    Normal, // ~150 DPI (scale 2.08)
    High,   // ~216 DPI (scale 3.0)
    Ultra   // ~300 DPI (scale 4.17)
}

public sealed class PdfToImagesOptions
{
    public PdfImageConversionMode Mode { get; set; } = PdfImageConversionMode.PagesToImages;
    public ImageOutputFormat Format { get; set; } = ImageOutputFormat.Jpg;
    public ImageRenderQuality Quality { get; set; } = ImageRenderQuality.Normal;
    public string PageScope { get; set; } = "all"; // "all" | "selected" | "range"
    public string CustomRange { get; set; } = string.Empty;
    public HashSet<int> SelectedPages { get; set; } = new();

    public float GetRenderScale() => Quality switch
    {
        ImageRenderQuality.High => 3.0f,
        ImageRenderQuality.Ultra => 4.17f,
        _ => 2.08f
    };

    public float GetJpgQuality() => Quality switch
    {
        ImageRenderQuality.Ultra => 0.95f,
        ImageRenderQuality.High => 0.92f,
        _ => 0.88f
    };

    public string GetMimeType() => Format switch
    {
        ImageOutputFormat.Png => "image/png",
        ImageOutputFormat.WebP => "image/webp",
        _ => "image/jpeg"
    };

    public string GetFileExtension() => Format switch
    {
        ImageOutputFormat.Png => "png",
        ImageOutputFormat.WebP => "webp",
        _ => "jpg"
    };
}

public sealed class RenderedPdfPageImage
{
    public int PageNumber { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string MimeType { get; set; } = "image/jpeg";
    public byte[] Data { get; set; } = Array.Empty<byte>();
}

public sealed class ExtractedPdfImage
{
    public int Index { get; set; }
    public int PageNumber { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string MimeType { get; set; } = "image/png";
    public long SizeBytes => Data.Length;
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public string PreviewDataUrl { get; set; } = string.Empty;
    public bool IsSelected { get; set; } = true;

    public string FormattedSize => SizeBytes switch
    {
        < 1024 => $"{SizeBytes} B",
        < 1024 * 1024 => $"{SizeBytes / 1024.0:F1} KB",
        _ => $"{SizeBytes / (1024.0 * 1024.0):F2} MB"
    };
}
