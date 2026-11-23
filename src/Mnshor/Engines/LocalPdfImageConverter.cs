using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Mnshor.Models;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using SkiaSharp;

namespace Mnshor.Engines;

/// <summary>
/// Pure-.NET, 100% in-memory PDF-to-Images engine supporting:
/// 1) Packaging client-rendered PDF pages (JPG / PNG / WebP) as a single image or in-memory ZIP archive.
/// 2) Extracting embedded raster images (/XObject /Subtype /Image, including Form XObjects & /SMask alpha)
///    directly from PDF streams at native resolution.
/// </summary>
public sealed class LocalPdfImageConverter : IPdfImageConverter
{
    public async Task<IReadOnlyList<ExtractedPdfImage>> ExtractEmbeddedImagesAsync(
        byte[] pdfData,
        ImageOutputFormat targetFormat = ImageOutputFormat.Png,
        IProgress<(int Current, int Total)>? progress = null)
    {
        if (pdfData == null || pdfData.Length == 0)
            return Array.Empty<ExtractedPdfImage>();

        return await Task.Run(() =>
        {
            var extracted = new List<ExtractedPdfImage>();
            var visitedObjects = new HashSet<PdfDictionary>();

            using var ms = new MemoryStream(pdfData, writable: false);
            using var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

            int totalPages = doc.PageCount;
            int imageIndex = 0;

            for (int p = 0; p < totalPages; p++)
            {
                int pageNum = p + 1;
                var page = doc.Pages[p];
                var resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    CollectImagesFromResources(
                        resources,
                        pageNum,
                        targetFormat,
                        visitedObjects,
                        extracted,
                        ref imageIndex,
                        depth: 0);
                }

                progress?.Report((pageNum, totalPages));
            }

            return (IReadOnlyList<ExtractedPdfImage>)extracted;
        });
    }

    public async Task<ExecutionResult> PackageImagesAsync(
        string originalFileName,
        IReadOnlyList<RenderedPdfPageImage> images,
        ImageOutputFormat format)
    {
        if (images == null || images.Count == 0)
        {
            return ExecutionResult.Fail("No images were provided to package.");
        }

        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                string baseName = Path.GetFileNameWithoutExtension(
                    string.IsNullOrWhiteSpace(originalFileName) ? "Document.pdf" : originalFileName);
                string ext = GetExtension(format);
                string mime = GetMimeType(format);

                // Single image -> download directly without wrapping in a ZIP
                if (images.Count == 1)
                {
                    var single = images[0];
                    string singleName = !string.IsNullOrWhiteSpace(single.FileName)
                        ? single.FileName
                        : $"{baseName}_page_{single.PageNumber:D3}.{ext}";
                    sw.Stop();
                    return ExecutionResult.Ok(
                        single.Data,
                        singleName,
                        $"Exported 1 image ({ext.ToUpperInvariant()}) in {sw.ElapsedMilliseconds} ms",
                        sw.ElapsedMilliseconds,
                        string.IsNullOrWhiteSpace(single.MimeType) ? mime : single.MimeType);
                }

                // Multiple images -> package into an in-memory ZIP archive
                using var zipMs = new MemoryStream();
                using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, leaveOpen: true))
                {
                    var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    int seq = 1;

                    foreach (var img in images)
                    {
                        if (img.Data == null || img.Data.Length == 0) continue;

                        string entryName = !string.IsNullOrWhiteSpace(img.FileName)
                            ? img.FileName
                            : $"{baseName}_page_{img.PageNumber:D3}.{ext}";

                        if (!usedNames.Add(entryName))
                        {
                            string stem = Path.GetFileNameWithoutExtension(entryName);
                            string eExt = Path.GetExtension(entryName);
                            entryName = $"{stem}_{seq}{eExt}";
                            usedNames.Add(entryName);
                        }

                        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
                        using var entryStream = entry.Open();
                        entryStream.Write(img.Data, 0, img.Data.Length);
                        seq++;
                    }
                }

                sw.Stop();
                string zipName = $"{baseName}_images_{DateTime.Now:yyyyMMdd_HHmmss}.zip";
                return ExecutionResult.Ok(
                    zipMs.ToArray(),
                    zipName,
                    $"Packaged {images.Count} image(s) ({ext.ToUpperInvariant()}) into ZIP in {sw.ElapsedMilliseconds} ms",
                    sw.ElapsedMilliseconds,
                    "application/zip");
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Image packaging error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private static void CollectImagesFromResources(
        PdfDictionary resources,
        int pageNum,
        ImageOutputFormat targetFormat,
        HashSet<PdfDictionary> visited,
        List<ExtractedPdfImage> output,
        ref int imageIndex,
        int depth)
    {
        if (depth > 6) return;

        var xObjects = resources.Elements.GetDictionary("/XObject");
        if (xObjects == null) return;

        foreach (var key in xObjects.Elements.Keys)
        {
            var dict = ResolveDictionary(xObjects.Elements[key]);
            if (dict == null || !visited.Add(dict)) continue;

            var subtype = dict.Elements.GetName("/Subtype");
            if (string.Equals(subtype, "/Form", StringComparison.Ordinal))
            {
                var formRes = dict.Elements.GetDictionary("/Resources");
                if (formRes != null)
                {
                    CollectImagesFromResources(formRes, pageNum, targetFormat, visited, output, ref imageIndex, depth + 1);
                }
                continue;
            }

            if (!string.Equals(subtype, "/Image", StringComparison.Ordinal))
                continue;

            // Skip soft-mask-only 1-bit stencil masks
            if (dict.Elements.GetBoolean("/ImageMask"))
                continue;

            try
            {
                var item = TryExtractSingleImage(dict, pageNum, targetFormat, imageIndex + 1);
                if (item != null)
                {
                    imageIndex++;
                    output.Add(item);
                }
            }
            catch
            {
                // Skip unsupported or malformed image streams cleanly
            }
        }
    }

    private static ExtractedPdfImage? TryExtractSingleImage(
        PdfDictionary imgDict,
        int pageNum,
        ImageOutputFormat targetFormat,
        int index)
    {
        if (imgDict.Stream?.Value == null || imgDict.Stream.Value.Length == 0)
            return null;

        int width = imgDict.Elements.GetInteger("/Width");
        int height = imgDict.Elements.GetInteger("/Height");
        if (width < 4 || height < 4) return null;

        string filter = GetPrimaryFilterName(imgDict);
        var smaskDict = ResolveDictionary(imgDict.Elements["/SMask"]);
        byte[]? smaskAlpha = TryReadSMaskAlpha(smaskDict, width, height);

        string ext = GetExtension(targetFormat);
        string mime = GetMimeType(targetFormat);

        // Case 1: Native JPEG (/DCTDecode) with no alpha mask -> direct or re-encoded
        if (string.Equals(filter, "/DCTDecode", StringComparison.Ordinal))
        {
            byte[] jpegBytes = imgDict.Stream.Value;
            if (smaskAlpha == null && targetFormat == ImageOutputFormat.Jpg)
            {
                using var skImg = SKImage.FromEncodedData(jpegBytes);
                if (skImg == null) return null;
                return BuildExtractedItem(index, pageNum, skImg.Width, skImg.Height, jpegBytes, ext, mime, skImg);
            }

            using var decodedBmp = SKBitmap.Decode(jpegBytes);
            if (decodedBmp == null) return null;
            using var finalImg = ApplyAlphaMaskIfPresent(decodedBmp, smaskAlpha);
            byte[] encoded = EncodeSkImage(finalImg, targetFormat);
            if (encoded.Length == 0) return null;
            return BuildExtractedItem(index, pageNum, finalImg.Width, finalImg.Height, encoded, ext, mime, finalImg);
        }

        // Case 2: Raw / Flate-compressed raster pixels (DeviceRGB, DeviceGray, DeviceCMYK)
        int bpc = imgDict.Elements.GetInteger("/BitsPerComponent");
        if (bpc != 8) return null;

        byte[] raw = imgDict.Stream.UnfilteredValue;
        if (raw == null || raw.Length == 0) return null;

        string colorSpace = GetColorSpaceName(imgDict);
        using var bmp = DecodeRawPixelsToBitmap(raw, width, height, colorSpace, smaskAlpha);
        if (bmp == null) return null;

        using var skImage = SKImage.FromBitmap(bmp);
        byte[] outBytes = EncodeSkImage(skImage, targetFormat);
        if (outBytes.Length == 0) return null;

        return BuildExtractedItem(index, pageNum, width, height, outBytes, ext, mime, skImage);
    }

    private static ExtractedPdfImage BuildExtractedItem(
        int index,
        int pageNum,
        int width,
        int height,
        byte[] data,
        string ext,
        string mime,
        SKImage skImg)
    {
        string previewUrl = BuildThumbnailDataUrl(skImg);
        return new ExtractedPdfImage
        {
            Index = index,
            PageNumber = pageNum,
            Width = width,
            Height = height,
            FileName = $"image_{index:D3}_p{pageNum}.{ext}",
            MimeType = mime,
            Data = data,
            PreviewDataUrl = previewUrl,
            IsSelected = true
        };
    }

    private static string BuildThumbnailDataUrl(SKImage source)
    {
        const int maxSide = 320;
        if (source.Width <= maxSide && source.Height <= maxSide)
        {
            using var d = source.Encode(SKEncodedImageFormat.Png, 85);
            return d != null ? $"data:image/png;base64,{Convert.ToBase64String(d.ToArray())}" : string.Empty;
        }

        float scale = Math.Min((float)maxSide / source.Width, (float)maxSide / source.Height);
        int tw = Math.Max(16, (int)Math.Round(source.Width * scale));
        int th = Math.Max(16, (int)Math.Round(source.Height * scale));

        using var surface = SKSurface.Create(new SKImageInfo(tw, th, SKColorType.Rgba8888, SKAlphaType.Premul));
        if (surface == null) return string.Empty;
        surface.Canvas.Clear(SKColors.Transparent);
        surface.Canvas.DrawImage(source, new SKRect(0, 0, tw, th), new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
        using var snap = surface.Snapshot();
        using var enc = snap.Encode(SKEncodedImageFormat.Png, 85);
        return enc != null ? $"data:image/png;base64,{Convert.ToBase64String(enc.ToArray())}" : string.Empty;
    }

    private static SKImage ApplyAlphaMaskIfPresent(SKBitmap src, byte[]? smaskAlpha)
    {
        if (smaskAlpha == null || smaskAlpha.Length < src.Width * src.Height)
            return SKImage.FromBitmap(src);

        using var rgba = new SKBitmap(src.Width, src.Height, SKColorType.Rgba8888, SKAlphaType.Unpremul);
        int pixelCount = src.Width * src.Height;
        var dstPixels = new int[pixelCount];
        for (int y = 0, idx = 0; y < src.Height; y++)
        {
            for (int x = 0; x < src.Width; x++, idx++)
            {
                var c = src.GetPixel(x, y);
                byte a = smaskAlpha[idx];
                dstPixels[idx] = (a << 24) | (c.Blue << 16) | (c.Green << 8) | c.Red;
            }
        }
        Marshal.Copy(dstPixels, 0, rgba.GetPixels(), pixelCount);
        return SKImage.FromBitmap(rgba);
    }

    private static SKBitmap? DecodeRawPixelsToBitmap(
        byte[] raw,
        int width,
        int height,
        string colorSpace,
        byte[]? smaskAlpha)
    {
        int pixelCount = width * height;
        var pixels = new int[pixelCount];

        if (string.Equals(colorSpace, "/DeviceRGB", StringComparison.Ordinal))
        {
            if (raw.Length < pixelCount * 3) return null;
            for (int i = 0, p = 0; p < pixelCount; i += 3, p++)
            {
                byte r = raw[i];
                byte g = raw[i + 1];
                byte b = raw[i + 2];
                byte a = smaskAlpha != null && p < smaskAlpha.Length ? smaskAlpha[p] : (byte)255;
                pixels[p] = (a << 24) | (b << 16) | (g << 8) | r;
            }
        }
        else if (string.Equals(colorSpace, "/DeviceGray", StringComparison.Ordinal))
        {
            if (raw.Length < pixelCount) return null;
            for (int p = 0; p < pixelCount; p++)
            {
                byte v = raw[p];
                byte a = smaskAlpha != null && p < smaskAlpha.Length ? smaskAlpha[p] : (byte)255;
                pixels[p] = (a << 24) | (v << 16) | (v << 8) | v;
            }
        }
        else if (string.Equals(colorSpace, "/DeviceCMYK", StringComparison.Ordinal))
        {
            if (raw.Length < pixelCount * 4) return null;
            for (int i = 0, p = 0; p < pixelCount; i += 4, p++)
            {
                float c = raw[i] / 255f;
                float m = raw[i + 1] / 255f;
                float y = raw[i + 2] / 255f;
                float k = raw[i + 3] / 255f;
                byte r = (byte)Math.Clamp((int)Math.Round(255f * (1f - c) * (1f - k)), 0, 255);
                byte g = (byte)Math.Clamp((int)Math.Round(255f * (1f - m) * (1f - k)), 0, 255);
                byte b = (byte)Math.Clamp((int)Math.Round(255f * (1f - y) * (1f - k)), 0, 255);
                byte a = smaskAlpha != null && p < smaskAlpha.Length ? smaskAlpha[p] : (byte)255;
                pixels[p] = (a << 24) | (b << 16) | (g << 8) | r;
            }
        }
        else
        {
            return null;
        }

        var alphaType = smaskAlpha != null ? SKAlphaType.Unpremul : SKAlphaType.Opaque;
        var bmp = new SKBitmap(width, height, SKColorType.Rgba8888, alphaType);
        Marshal.Copy(pixels, 0, bmp.GetPixels(), pixelCount);
        return bmp;
    }

    private static byte[]? TryReadSMaskAlpha(PdfDictionary? smaskDict, int expectedW, int expectedH)
    {
        if (smaskDict?.Stream?.Value == null) return null;
        int mw = smaskDict.Elements.GetInteger("/Width");
        int mh = smaskDict.Elements.GetInteger("/Height");
        int bpc = smaskDict.Elements.GetInteger("/BitsPerComponent");
        if (mw != expectedW || mh != expectedH || bpc != 8) return null;

        try
        {
            byte[] raw = smaskDict.Stream.UnfilteredValue;
            return raw != null && raw.Length >= expectedW * expectedH ? raw : null;
        }
        catch
        {
            return null;
        }
    }

    private static byte[] EncodeSkImage(SKImage img, ImageOutputFormat format)
    {
        if (format == ImageOutputFormat.Jpg)
        {
            // Composite onto white background so transparent pixels don't turn black in JPEG
            using var surface = SKSurface.Create(new SKImageInfo(img.Width, img.Height, SKColorType.Rgba8888, SKAlphaType.Opaque));
            if (surface != null)
            {
                surface.Canvas.Clear(SKColors.White);
                surface.Canvas.DrawImage(img, 0, 0);
                using var snap = surface.Snapshot();
                using var jpg = snap.Encode(SKEncodedImageFormat.Jpeg, 92);
                return jpg?.ToArray() ?? Array.Empty<byte>();
            }
            using var fallback = img.Encode(SKEncodedImageFormat.Jpeg, 92);
            return fallback?.ToArray() ?? Array.Empty<byte>();
        }

        var skFmt = format == ImageOutputFormat.WebP ? SKEncodedImageFormat.Webp : SKEncodedImageFormat.Png;
        using var data = img.Encode(skFmt, format == ImageOutputFormat.WebP ? 90 : 95);
        return data?.ToArray() ?? Array.Empty<byte>();
    }

    private static PdfDictionary? ResolveDictionary(PdfItem? item)
    {
        if (item is PdfReference pdfRef)
            return pdfRef.Value as PdfDictionary;
        return item as PdfDictionary;
    }

    private static string GetPrimaryFilterName(PdfDictionary dict)
    {
        var filterItem = dict.Elements["/Filter"];
        if (filterItem is PdfName name) return name.Value;
        if (filterItem is PdfArray arr && arr.Elements.Count > 0 && arr.Elements[0] is PdfName first)
            return first.Value;
        return string.Empty;
    }

    private static string GetColorSpaceName(PdfDictionary dict)
    {
        var csItem = dict.Elements["/ColorSpace"];
        if (csItem is PdfReference csRef) csItem = csRef.Value;
        if (csItem is PdfName name) return name.Value;
        if (csItem is PdfArray arr && arr.Elements.Count > 0 && arr.Elements[0] is PdfName first)
        {
            // Handle ICCBased [ /ICCBased stream ] by checking /N component count
            if (string.Equals(first.Value, "/ICCBased", StringComparison.Ordinal) && arr.Elements.Count > 1)
            {
                var iccDict = ResolveDictionary(arr.Elements[1]);
                int n = iccDict?.Elements.GetInteger("/N") ?? 0;
                return n switch
                {
                    1 => "/DeviceGray",
                    3 => "/DeviceRGB",
                    4 => "/DeviceCMYK",
                    _ => string.Empty
                };
            }
            return first.Value;
        }
        return string.Empty;
    }

    private static string GetExtension(ImageOutputFormat format) => format switch
    {
        ImageOutputFormat.Png => "png",
        ImageOutputFormat.WebP => "webp",
        _ => "jpg"
    };

    private static string GetMimeType(ImageOutputFormat format) => format switch
    {
        ImageOutputFormat.Png => "image/png",
        ImageOutputFormat.WebP => "image/webp",
        _ => "image/jpeg"
    };

    // =========================================================================
    // Images -> PDF (JPG / PNG / WebP / BMP / GIF -> Single PDF or Separate ZIP)
    // =========================================================================

    public async Task<ImageToPdfInputItem?> InspectImageAsync(string fileName, string contentType, byte[] data)
    {
        if (data == null || data.Length == 0)
            return null;

        return await Task.Run(() =>
        {
            try
            {
                using var decoded = DecodeWithExifOrientation(data, extraRotationDegrees: 0);
                if (decoded == null || decoded.Width <= 0 || decoded.Height <= 0)
                    return null;

                using var skImg = SKImage.FromBitmap(decoded);
                string previewUrl = BuildThumbnailDataUrl(skImg);

                return new ImageToPdfInputItem
                {
                    FileName = string.IsNullOrWhiteSpace(fileName) ? "image.png" : fileName,
                    ContentType = string.IsNullOrWhiteSpace(contentType) ? "image/png" : contentType,
                    Data = data,
                    Width = decoded.Width,
                    Height = decoded.Height,
                    RotationDegrees = 0,
                    PreviewDataUrl = previewUrl
                };
            }
            catch
            {
                return null;
            }
        });
    }

    public async Task<ExecutionResult> ConvertImagesToPdfAsync(
        IReadOnlyList<ImageToPdfInputItem> images,
        ImagesToPdfOptions? options = null,
        IProgress<(int Current, int Total)>? progress = null)
    {
        if (images == null || images.Count == 0)
        {
            return ExecutionResult.Fail("Please add at least one image to convert to PDF.");
        }

        options ??= new ImagesToPdfOptions();
        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                double marginPt = options.GetMarginPoints();

                // Mode A: Merge all images into a single PDF (or single image uploaded)
                if (options.MergeIntoSinglePdf || images.Count == 1)
                {
                    using var doc = new PdfDocument();
                    doc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
                    doc.Info.Creator = "Mnshor Local PDF Suite";

                    int added = 0;
                    for (int i = 0; i < images.Count; i++)
                    {
                        var item = images[i];
                        if (item.Data == null || item.Data.Length == 0) continue;

                        if (AddImagePageToDocument(doc, item, options, marginPt))
                        {
                            added++;
                        }

                        progress?.Report((i + 1, images.Count));
                    }

                    if (added == 0)
                    {
                        return ExecutionResult.Fail("None of the uploaded images could be decoded into PDF pages.");
                    }

                    using var outMs = new MemoryStream();
                    doc.Save(outMs, false);
                    sw.Stop();

                    string baseName = images.Count == 1
                        ? Path.GetFileNameWithoutExtension(images[0].FileName)
                        : "Images_Merged";
                    if (string.IsNullOrWhiteSpace(baseName)) baseName = "Images";
                    string fileName = $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

                    return ExecutionResult.Ok(
                        outMs.ToArray(),
                        fileName,
                        $"Converted {added} image(s) into a {added}-page PDF in {sw.ElapsedMilliseconds} ms",
                        sw.ElapsedMilliseconds,
                        "application/pdf");
                }

                // Mode B: Separate PDF per image packaged into an in-memory ZIP archive
                using var zipMs = new MemoryStream();
                int exportedPdfs = 0;

                using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, leaveOpen: true))
                {
                    var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    for (int i = 0; i < images.Count; i++)
                    {
                        var item = images[i];
                        if (item.Data == null || item.Data.Length == 0) continue;

                        using var singleDoc = new PdfDocument();
                        singleDoc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
                        singleDoc.Info.Creator = "Mnshor Local PDF Suite";

                        if (AddImagePageToDocument(singleDoc, item, options, marginPt))
                        {
                            using var singleMs = new MemoryStream();
                            singleDoc.Save(singleMs, false);
                            byte[] pdfBytes = singleMs.ToArray();

                            string stem = Path.GetFileNameWithoutExtension(item.FileName);
                            if (string.IsNullOrWhiteSpace(stem)) stem = $"image_{i + 1:D3}";
                            string entryName = $"{stem}.pdf";
                            if (!usedNames.Add(entryName))
                            {
                                entryName = $"{stem}_{i + 1:D3}.pdf";
                                usedNames.Add(entryName);
                            }

                            var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
                            using var entryStream = entry.Open();
                            entryStream.Write(pdfBytes, 0, pdfBytes.Length);
                            exportedPdfs++;
                        }

                        progress?.Report((i + 1, images.Count));
                    }
                }

                if (exportedPdfs == 0)
                {
                    return ExecutionResult.Fail("None of the uploaded images could be converted to PDF.");
                }

                sw.Stop();
                string zipName = $"Images_PDFs_{DateTime.Now:yyyyMMdd_HHmmss}.zip";
                return ExecutionResult.Ok(
                    zipMs.ToArray(),
                    zipName,
                    $"Converted {exportedPdfs} image(s) into {exportedPdfs} separate PDF(s) (ZIP) in {sw.ElapsedMilliseconds} ms",
                    sw.ElapsedMilliseconds,
                    "application/zip");
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"Images to PDF error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private static bool AddImagePageToDocument(
        PdfDocument doc,
        ImageToPdfInputItem item,
        ImagesToPdfOptions options,
        double marginPt)
    {
        using var bmp = DecodeWithExifOrientation(item.Data, item.RotationDegrees);
        if (bmp == null || bmp.Width <= 0 || bmp.Height <= 0)
            return false;

        // Determine whether the bitmap has any non-opaque pixel so opaque photos use compact JPEG
        // while transparent PNGs/WebPs preserve transparency cleanly (or composite cleanly).
        bool hasAlpha = HasTransparency(bmp);
        byte[] normalizedBytes;
        using (var skImg = SKImage.FromBitmap(bmp))
        {
            if (hasAlpha)
            {
                using var png = skImg.Encode(SKEncodedImageFormat.Png, 95);
                normalizedBytes = png?.ToArray() ?? Array.Empty<byte>();
            }
            else
            {
                using var jpg = skImg.Encode(SKEncodedImageFormat.Jpeg, 92);
                normalizedBytes = jpg?.ToArray() ?? Array.Empty<byte>();
            }
        }

        if (normalizedBytes.Length == 0)
            return false;

        double pageW;
        double pageH;

        if (options.PageSize == ImagePdfPageSize.FitToImage)
        {
            // Map pixels at 96 DPI -> 72 pt/in (0.75 pt per px) with safe clamping for very large photos
            const double pxToPt = 72.0 / 96.0;
            double imgWpt = bmp.Width * pxToPt;
            double imgHpt = bmp.Height * pxToPt;

            // Keep extreme high-res images within reasonable PDF page bounds (max 1440 pt ~ 20 in)
            double maxDim = Math.Max(imgWpt, imgHpt);
            if (maxDim > 1440.0)
            {
                double clampScale = 1440.0 / maxDim;
                imgWpt *= clampScale;
                imgHpt *= clampScale;
            }

            pageW = Math.Max(72.0, imgWpt + (marginPt * 2.0));
            pageH = Math.Max(72.0, imgHpt + (marginPt * 2.0));
        }
        else
        {
            (double baseW, double baseH) = options.PageSize == ImagePdfPageSize.Letter
                ? (612.0, 792.0)      // US Letter (8.5 x 11 in)
                : (595.28, 841.89);   // A4 (210 x 297 mm)

            if (options.Orientation == ImagePdfOrientation.Landscape)
            {
                pageW = baseH;
                pageH = baseW;
            }
            else
            {
                pageW = baseW;
                pageH = baseH;
            }
        }

        var page = doc.AddPage();
        page.Width = XUnit.FromPoint(pageW);
        page.Height = XUnit.FromPoint(pageH);

        double availW = Math.Max(10.0, pageW - (marginPt * 2.0));
        double availH = Math.Max(10.0, pageH - (marginPt * 2.0));

        double scale = Math.Min(availW / bmp.Width, availH / bmp.Height);
        double drawW = bmp.Width * scale;
        double drawH = bmp.Height * scale;
        double drawX = marginPt + (availW - drawW) / 2.0;
        double drawY = marginPt + (availH - drawH) / 2.0;

        using var imgStream = new MemoryStream(normalizedBytes, writable: false);
        using var xImg = XImage.FromStream(imgStream);
        using var gfx = XGraphics.FromPdfPage(page);
        gfx.DrawImage(xImg, drawX, drawY, drawW, drawH);

        return true;
    }

    private static SKBitmap? DecodeWithExifOrientation(byte[] data, int extraRotationDegrees)
    {
        using var skData = SKData.CreateCopy(data);
        using var codec = SKCodec.Create(skData);
        if (codec == null)
            return null;

        var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Unpremul);
        using var rawBmp = new SKBitmap(info);
        var result = codec.GetPixels(info, rawBmp.GetPixels());
        if (result != SKCodecResult.Success && result != SKCodecResult.IncompleteInput)
            return null;

        SKBitmap oriented = ApplyExifOrigin(rawBmp, codec.EncodedOrigin);
        int normRot = ((extraRotationDegrees % 360) + 360) % 360;
        if (normRot == 0)
            return oriented;

        try
        {
            return RotateBitmap(oriented, normRot);
        }
        finally
        {
            oriented.Dispose();
        }
    }

    private static SKBitmap ApplyExifOrigin(SKBitmap src, SKEncodedOrigin origin)
    {
        if (origin == SKEncodedOrigin.TopLeft || origin == SKEncodedOrigin.Default)
            return src.Copy();

        bool swap = origin is SKEncodedOrigin.RightTop or SKEncodedOrigin.LeftBottom
            or SKEncodedOrigin.LeftTop or SKEncodedOrigin.RightBottom;

        int dstW = swap ? src.Height : src.Width;
        int dstH = swap ? src.Width : src.Height;

        var dst = new SKBitmap(dstW, dstH, src.ColorType, src.AlphaType);
        using var canvas = new SKCanvas(dst);
        canvas.Clear(SKColors.Transparent);

        switch (origin)
        {
            case SKEncodedOrigin.TopRight:
                canvas.Scale(-1, 1, dstW / 2f, 0);
                break;
            case SKEncodedOrigin.BottomRight:
                canvas.RotateDegrees(180, dstW / 2f, dstH / 2f);
                break;
            case SKEncodedOrigin.BottomLeft:
                canvas.Scale(1, -1, 0, dstH / 2f);
                break;
            case SKEncodedOrigin.LeftTop:
                canvas.RotateDegrees(90, dstW / 2f, dstH / 2f);
                canvas.Translate((dstW - src.Width) / 2f, (dstH - src.Height) / 2f);
                canvas.Scale(1, -1, src.Width / 2f, src.Height / 2f);
                canvas.DrawBitmap(src, 0, 0);
                return dst;
            case SKEncodedOrigin.RightTop:
                canvas.Translate(dstW, 0);
                canvas.RotateDegrees(90);
                break;
            case SKEncodedOrigin.RightBottom:
                canvas.RotateDegrees(270, dstW / 2f, dstH / 2f);
                canvas.Translate((dstW - src.Width) / 2f, (dstH - src.Height) / 2f);
                canvas.Scale(1, -1, src.Width / 2f, src.Height / 2f);
                canvas.DrawBitmap(src, 0, 0);
                return dst;
            case SKEncodedOrigin.LeftBottom:
                canvas.Translate(0, dstH);
                canvas.RotateDegrees(270);
                break;
        }

        canvas.DrawBitmap(src, 0, 0);
        return dst;
    }

    private static SKBitmap RotateBitmap(SKBitmap src, int degrees)
    {
        bool swap = degrees is 90 or 270;
        int dstW = swap ? src.Height : src.Width;
        int dstH = swap ? src.Width : src.Height;

        var dst = new SKBitmap(dstW, dstH, src.ColorType, src.AlphaType);
        using var canvas = new SKCanvas(dst);
        canvas.Clear(SKColors.Transparent);

        if (degrees == 90)
        {
            canvas.Translate(dstW, 0);
            canvas.RotateDegrees(90);
        }
        else if (degrees == 180)
        {
            canvas.Translate(dstW, dstH);
            canvas.RotateDegrees(180);
        }
        else if (degrees == 270)
        {
            canvas.Translate(0, dstH);
            canvas.RotateDegrees(270);
        }

        canvas.DrawBitmap(src, 0, 0);
        return dst;
    }

    private static bool HasTransparency(SKBitmap bmp)
    {
        // Fast strided sample check (exact on small images, < 1ms on 4K images)
        int stepX = Math.Max(1, bmp.Width / 64);
        int stepY = Math.Max(1, bmp.Height / 64);
        for (int y = 0; y < bmp.Height; y += stepY)
        {
            for (int x = 0; x < bmp.Width; x += stepX)
            {
                if (bmp.GetPixel(x, y).Alpha < 250)
                    return true;
            }
        }
        return false;
    }
}
