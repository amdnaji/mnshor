namespace Mnshor.Models;

public enum PowerPointConversionMode
{
    PptxToPdf,
    PdfToPptx
}

public enum PptxSlideSizePreset
{
    MatchSource,
    Widescreen16x9,
    Standard4x3
}

public class PowerPointPdfOptions
{
    public PowerPointConversionMode Mode { get; set; } = PowerPointConversionMode.PptxToPdf;

    // PPTX -> PDF options
    public HashSet<int> SelectedSlideNumbers { get; set; } = new();
    public bool IncludeSlideNumbers { get; set; } = true;
    public bool IncludeSpeakerNotes { get; set; } = false;

    // PDF -> PPTX options
    public PptxSlideSizePreset SlideSizePreset { get; set; } = PptxSlideSizePreset.MatchSource;
    public string PageScope { get; set; } = "all"; // "all" or "range"
    public string CustomRange { get; set; } = string.Empty;
    public float RenderScale { get; set; } = 2.0f; // ~144-200 DPI
}

public class PptxPresentationInfo
{
    public bool IsValid { get; set; }
    public string? ErrorMessage { get; set; }
    public int SlideCount { get; set; }
    public float SlideWidthPoints { get; set; }
    public float SlideHeightPoints { get; set; }
    public List<PptxSlideSummary> Slides { get; set; } = new();
}

public class PptxSlideSummary
{
    public int SlideNumber { get; set; }
    public string Title { get; set; } = string.Empty;
    public string PreviewSnippet { get; set; } = string.Empty;
    public int ShapeCount { get; set; }
    public int ImageCount { get; set; }
    public int TableCount { get; set; }
    public bool HasSpeakerNotes { get; set; }
    public bool IsRtlDetected { get; set; }
    public string PreviewDataUrl { get; set; } = string.Empty;
}

public class PdfPageSlideImage
{
    public int PageNumber { get; set; }
    public double WidthPoints { get; set; }
    public double HeightPoints { get; set; }
    public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
}
