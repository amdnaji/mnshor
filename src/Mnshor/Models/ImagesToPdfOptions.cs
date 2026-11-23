namespace Mnshor.Models;

public enum ImagePdfPageSize
{
    FitToImage,
    A4,
    Letter
}

public enum ImagePdfOrientation
{
    Portrait,
    Landscape
}

public enum ImagePdfMargin
{
    None,  // 0 mm (0 pt)
    Small, // 10 mm (~28.35 pt)
    Big    // 25 mm (~70.87 pt)
}

public sealed class ImagesToPdfOptions
{
    public ImagePdfPageSize PageSize { get; set; } = ImagePdfPageSize.A4;
    public ImagePdfOrientation Orientation { get; set; } = ImagePdfOrientation.Portrait;
    public ImagePdfMargin Margin { get; set; } = ImagePdfMargin.None;
    public bool MergeIntoSinglePdf { get; set; } = true;

    public double GetMarginPoints() => Margin switch
    {
        ImagePdfMargin.Small => 28.3465, // 10 mm
        ImagePdfMargin.Big => 70.8661,   // 25 mm
        _ => 0.0
    };
}

public sealed class ImageToPdfInputItem
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = "image/jpeg";
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public int Width { get; set; }
    public int Height { get; set; }
    public int RotationDegrees { get; set; } = 0; // 0, 90, 180, 270
    public string PreviewDataUrl { get; set; } = string.Empty;

    public long SizeBytes => Data.Length;

    public string FormattedSize => SizeBytes switch
    {
        < 1024 => $"{SizeBytes} B",
        < 1024 * 1024 => $"{SizeBytes / 1024.0:F1} KB",
        _ => $"{SizeBytes / (1024.0 * 1024.0):F2} MB"
    };
}
