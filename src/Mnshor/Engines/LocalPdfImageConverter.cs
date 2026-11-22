using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Mnshor.Models;
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
}
