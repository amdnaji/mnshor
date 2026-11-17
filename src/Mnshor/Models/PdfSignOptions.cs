namespace Mnshor.Models;

public enum SignatureInputMethod
{
    TypeOsFont,
    DrawFreehand,
    UploadImage
}

public enum SignFieldType
{
    Signature,
    Initials,
    Name,
    Date,
    CustomText,
    Stamp
}

public enum SignPageScope
{
    SpecificPage,
    AllPages,
    LastPage,
    CustomRange
}

public enum OsFontCategory
{
    SignatureScript,
    ArabicCalligraphy,
    SystemFont
}

public class OsFontDescriptor
{
    public string FamilyName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public OsFontCategory Category { get; set; } = OsFontCategory.SystemFont;
}

public class PdfSignatureField
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public SignFieldType FieldType { get; set; } = SignFieldType.Signature;

    public string Text { get; set; } = string.Empty;

    public string FontFamily { get; set; } = "Segoe Script";

    public string TextColorHex { get; set; } = "#0F172A";

    public bool IsBold { get; set; } = false;

    public bool IsItalic { get; set; } = false;

    /// <summary>
    /// High-DPI transparent PNG bytes representing this field (from OS Font shaper, Freehand canvas, or Uploaded image).
    /// </summary>
    public byte[]? RenderedPngBytes { get; set; }

    /// <summary>
    /// Base64 data URL (data:image/png;base64,...) for instant visual display on the page stage.
    /// </summary>
    public string PreviewDataUrl =>
        RenderedPngBytes != null && RenderedPngBytes.Length > 0
            ? $"data:image/png;base64,{Convert.ToBase64String(RenderedPngBytes)}"
            : string.Empty;

    /// <summary>
    /// Normalized horizontal position (0.0 = left edge, 1.0 = right edge of PDF page).
    /// </summary>
    public double XPercent { get; set; } = 0.62;

    /// <summary>
    /// Normalized vertical position (0.0 = top edge, 1.0 = bottom edge of PDF page).
    /// </summary>
    public double YPercent { get; set; } = 0.78;

    /// <summary>
    /// Normalized width as a fraction of the PDF page width (e.g. 0.24 = 24% of page width).
    /// </summary>
    public double WidthPercent { get; set; } = 0.24;

    /// <summary>
    /// Width / Height aspect ratio of the rendered signature or field image.
    /// </summary>
    public double AspectRatio { get; set; } = 3.0;

    public int RotationDegrees { get; set; } = 0;

    public int OpacityPercent { get; set; } = 100;

    /// <summary>
    /// 1-based page number where this field is placed when PageScope == SpecificPage.
    /// </summary>
    public int TargetPage { get; set; } = 1;

    public SignPageScope PageScope { get; set; } = SignPageScope.SpecificPage;

    public string CustomRange { get; set; } = string.Empty;
}

public class PdfSignOptions
{
    public string SignerFullName { get; set; } = "Ahmed Al-Mansoor";

    public string SignerInitials { get; set; } = "AM";

    public List<PdfSignatureField> Fields { get; set; } = new();
}
