using System.Diagnostics;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Mnshor.Models;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using SkiaSharp;
using SkiaSharp.HarfBuzz;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Mnshor.Engines;

public class LocalPowerPointConverter : IPowerPointConverter
{
    private const double EmusPerPoint = 12700.0;
    private readonly OsFontService _fontService;

    public LocalPowerPointConverter(OsFontService? fontService = null)
    {
        _fontService = fontService ?? new OsFontService();
    }

    // =========================================================================
    // Slide Context (Holds Slide -> SlideLayout -> SlideMaster -> Theme Chain)
    // =========================================================================

    private sealed class SlideRenderContext
    {
        public required PresentationPart PresentationPart { get; init; }
        public required SlidePart SlidePart { get; init; }
        public SlideLayoutPart? LayoutPart { get; init; }
        public SlideMasterPart? MasterPart { get; init; }
        public ThemePart? ThemePart { get; init; }
        public required Dictionary<string, SKColor> ThemeColors { get; init; }
        public required Dictionary<string, string> ColorMap { get; init; }
        public required string MajorLatinFont { get; init; }
        public required string MinorLatinFont { get; init; }
        public required string MajorComplexFont { get; init; }
        public required string MinorComplexFont { get; init; }
        public float SlideWidth { get; init; }
        public float SlideHeight { get; init; }
        public int SlideNumber { get; init; }
    }

    private static SlideRenderContext BuildSlideContext(
        PresentationPart presPart,
        SlidePart slidePart,
        float slideW,
        float slideH,
        int slideNum)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        var masterPart = layoutPart?.SlideMasterPart ?? presPart.SlideMasterParts.FirstOrDefault();
        var themePart = masterPart?.ThemePart ?? presPart.ThemePart;

        var themeColors = ExtractThemeColors(themePart);
        var colorMap = ExtractColorMap(masterPart?.SlideMaster?.ColorMap, layoutPart?.SlideLayout?.ColorMapOverride, slidePart.Slide?.ColorMapOverride);

        // Apply colorMap aliases (bg1, tx1, bg2, tx2, hlink, folHlink) into themeColors
        foreach (var (alias, targetSlot) in colorMap)
        {
            if (themeColors.TryGetValue(targetSlot, out var mappedColor))
            {
                themeColors[alias] = mappedColor;
            }
        }

        var fontScheme = themePart?.Theme?.ThemeElements?.FontScheme;
        string majorLatin = fontScheme?.MajorFont?.LatinFont?.Typeface?.Value ?? "Segoe UI";
        string minorLatin = fontScheme?.MinorFont?.LatinFont?.Typeface?.Value ?? "Segoe UI";
        string majorCs = fontScheme?.MajorFont?.ComplexScriptFont?.Typeface?.Value ?? majorLatin;
        string minorCs = fontScheme?.MinorFont?.ComplexScriptFont?.Typeface?.Value ?? minorLatin;

        if (string.IsNullOrWhiteSpace(majorCs)) majorCs = majorLatin;
        if (string.IsNullOrWhiteSpace(minorCs)) minorCs = minorLatin;

        return new SlideRenderContext
        {
            PresentationPart = presPart,
            SlidePart = slidePart,
            LayoutPart = layoutPart,
            MasterPart = masterPart,
            ThemePart = themePart,
            ThemeColors = themeColors,
            ColorMap = colorMap,
            MajorLatinFont = majorLatin,
            MinorLatinFont = minorLatin,
            MajorComplexFont = majorCs,
            MinorComplexFont = minorCs,
            SlideWidth = slideW,
            SlideHeight = slideH,
            SlideNumber = slideNum
        };
    }

    // =========================================================================
    // 1. Inspect Presentation (.pptx) + Render Visual Thumbnail Previews
    // =========================================================================

    public PptxPresentationInfo InspectPresentation(byte[] pptxData, IProgress<(int Current, int Total)>? progress = null)
    {
        if (pptxData == null || pptxData.Length == 0)
        {
            return new PptxPresentationInfo
            {
                IsValid = false,
                ErrorMessage = "Invalid or empty PowerPoint (.pptx) file."
            };
        }

        try
        {
            using var ms = new MemoryStream(pptxData, writable: false);
            using var doc = PresentationDocument.Open(ms, isEditable: false);

            var presPart = doc.PresentationPart;
            var pres = presPart?.Presentation;
            if (presPart == null || pres == null)
            {
                return new PptxPresentationInfo
                {
                    IsValid = false,
                    ErrorMessage = "PowerPoint file has no readable presentation structure."
                };
            }

            var (slideW, slideH) = GetSlideDimensionsPoints(pres);
            var slideIds = pres.SlideIdList?.Elements<P.SlideId>().ToList() ?? new List<P.SlideId>();

            if (slideIds.Count == 0)
            {
                return new PptxPresentationInfo
                {
                    IsValid = false,
                    ErrorMessage = "The PowerPoint presentation contains no slides."
                };
            }

            var info = new PptxPresentationInfo
            {
                IsValid = true,
                SlideCount = slideIds.Count,
                SlideWidthPoints = slideW,
                SlideHeightPoints = slideH
            };

            var defaultOpts = new PowerPointPdfOptions
            {
                IncludeSlideNumbers = false,
                IncludeSpeakerNotes = false
            };

            int slideNum = 1;
            foreach (var sId in slideIds)
            {
                string? relId = sId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relId) || presPart.GetPartById(relId) is not SlidePart slidePart)
                {
                    slideNum++;
                    continue;
                }

                var summary = ExtractSlideSummary(slidePart, slideNum);

                // Render a visual thumbnail preview for the first 36 slides
                if (slideNum <= 36)
                {
                    try
                    {
                        var ctx = BuildSlideContext(presPart, slidePart, slideW, slideH, slideNum);
                        float thumbW = 360f;
                        float scale = thumbW / Math.Max(1f, slideW);
                        int bmpW = Math.Max(160, (int)Math.Round(slideW * scale));
                        int bmpH = Math.Max(90, (int)Math.Round(slideH * scale));

                        using var surface = SKSurface.Create(new SKImageInfo(bmpW, bmpH, SKColorType.Rgba8888, SKAlphaType.Premul));
                        if (surface != null)
                        {
                            surface.Canvas.Scale(scale, scale);
                            RenderSlideToCanvas(surface.Canvas, ctx, defaultOpts);
                            using var img = surface.Snapshot();
                            using var enc = img.Encode(SKEncodedImageFormat.Jpeg, 82);
                            summary.PreviewDataUrl = $"data:image/jpeg;base64,{Convert.ToBase64String(enc.ToArray())}";
                        }
                    }
                    catch
                    {
                        // Ignore preview thumbnail error
                    }
                }

                info.Slides.Add(summary);
                progress?.Report((slideNum, slideIds.Count));
                slideNum++;
            }

            return info;
        }
        catch (Exception ex)
        {
            return new PptxPresentationInfo
            {
                IsValid = false,
                ErrorMessage = $"Failed to read PowerPoint (.pptx): {ex.Message}"
            };
        }
    }

    private static PptxSlideSummary ExtractSlideSummary(SlidePart slidePart, int slideNumber)
    {
        var slide = slidePart.Slide;
        var spTree = slide?.CommonSlideData?.ShapeTree;

        string title = string.Empty;
        var snippets = new List<string>();
        int shapeCount = 0;
        int imageCount = 0;
        int tableCount = 0;
        bool isRtl = false;

        if (spTree != null)
        {
            foreach (var shape in spTree.Descendants<P.Shape>())
            {
                shapeCount++;
                var (layoutShape, _) = FindInheritedPlaceholderShapes(shape, slidePart);
                var phType = GetEffectivePlaceholderType(shape, layoutShape);

                string text = ExtractShapeText(shape);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (!isRtl && OsFontService.ContainsArabic(text))
                    isRtl = true;

                if (string.IsNullOrEmpty(title) &&
                    (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle))
                {
                    title = text.Split('\n')[0].Trim();
                }
                else
                {
                    snippets.Add(text.Replace('\n', ' ').Trim());
                }
            }

            imageCount = spTree.Descendants<P.Picture>().Count();
            tableCount = spTree.Descendants<A.Table>().Count();
        }

        if (string.IsNullOrEmpty(title))
        {
            title = snippets.FirstOrDefault() ?? $"Slide {slideNumber}";
            if (title.Length > 48) title = title[..45] + "...";
        }

        string preview = string.Join(" • ", snippets.Take(2));
        if (preview.Length > 95) preview = preview[..92] + "...";

        bool hasNotes = slidePart.NotesSlidePart != null &&
                        !string.IsNullOrWhiteSpace(ExtractNotesText(slidePart.NotesSlidePart));

        return new PptxSlideSummary
        {
            SlideNumber = slideNumber,
            Title = title,
            PreviewSnippet = preview,
            ShapeCount = shapeCount,
            ImageCount = imageCount,
            TableCount = tableCount,
            HasSpeakerNotes = hasNotes,
            IsRtlDetected = isRtl
        };
    }

    // =========================================================================
    // 2. Convert PowerPoint (.pptx) -> High-Fidelity Vector PDF
    // =========================================================================

    public async Task<ExecutionResult> ConvertPptxToPdfAsync(
        DocumentItem document,
        PowerPointPdfOptions? options = null,
        IProgress<(int Current, int Total, string Stage)>? progress = null)
    {
        if (document == null || document.Data == null || document.Data.Length == 0)
        {
            return ExecutionResult.Fail("Invalid or empty PowerPoint (.pptx) document.");
        }

        options ??= new PowerPointPdfOptions { Mode = PowerPointConversionMode.PptxToPdf };
        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                // Primary pure-.NET engine: OfficeIMO.PowerPoint.Pdf (vectors, charts, SmartArt, shapes, tables, compact PDF)
                var officeImoResult = TryConvertWithOfficeImo(document, options, sw, progress);
                if (officeImoResult != null)
                {
                    return officeImoResult;
                }

                // Fallback in-memory SkiaSharp renderer
                using var pptxStream = new MemoryStream(document.Data, writable: false);
                using var presDoc = PresentationDocument.Open(pptxStream, isEditable: false);

                var presPart = presDoc.PresentationPart;
                var pres = presPart?.Presentation;
                if (presPart == null || pres == null)
                {
                    return ExecutionResult.Fail("PowerPoint document has no readable presentation.");
                }

                var slideIds = pres.SlideIdList?.Elements<P.SlideId>().ToList() ?? new List<P.SlideId>();
                if (slideIds.Count == 0)
                {
                    return ExecutionResult.Fail("Presentation contains no slides.");
                }

                var (slideW, slideH) = GetSlideDimensionsPoints(pres);
                float notesBandHeight = options.IncludeSpeakerNotes ? 86f : 0f;
                float pageW = slideW;
                float pageH = slideH + notesBandHeight;

                var selectedSet = options.SelectedSlideNumbers;
                string docTitle = Path.GetFileNameWithoutExtension(document.FileName);

                using var outStream = new MemoryStream();
                var metadata = new SKDocumentPdfMetadata
                {
                    Title = docTitle,
                    Creator = "Mnshor Local PDF Suite",
                    Producer = "Mnshor High-Fidelity PowerPoint Engine",
                    RasterDpi = 300f
                };

                int renderedSlides = 0;

                using (var pdfDoc = SKDocument.CreatePdf(outStream, metadata))
                {
                    for (int idx = 0; idx < slideIds.Count; idx++)
                    {
                        int slideNum = idx + 1;
                        if (selectedSet.Count > 0 && !selectedSet.Contains(slideNum))
                            continue;

                        string? relId = slideIds[idx].RelationshipId?.Value;
                        if (string.IsNullOrEmpty(relId) || presPart.GetPartById(relId) is not SlidePart slidePart)
                            continue;

                        var ctx = BuildSlideContext(presPart, slidePart, slideW, slideH, slideNum);
                        var canvas = pdfDoc.BeginPage(pageW, pageH);
                        RenderSlideToCanvas(canvas, ctx, options);
                        pdfDoc.EndPage();
                        renderedSlides++;
                    }

                    if (renderedSlides == 0)
                    {
                        return ExecutionResult.Fail("Please select at least one slide to convert.");
                    }

                    pdfDoc.Close();
                }

                sw.Stop();
                string fileName = $"{docTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                string summary = $"Converted {renderedSlides} PowerPoint slide(s) to PDF in {sw.ElapsedMilliseconds} ms";
                return ExecutionResult.Ok(outStream.ToArray(), fileName, summary, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"PowerPoint conversion error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private static ExecutionResult? TryConvertWithOfficeImo(
        DocumentItem document,
        PowerPointPdfOptions options,
        Stopwatch sw,
        IProgress<(int Current, int Total, string Stage)>? progress = null)
    {
        try
        {
            byte[] sourceBytes = document.Data;
            int slideCount;

            using var editMs = new MemoryStream();
            editMs.Write(sourceBytes, 0, sourceBytes.Length);
            editMs.Position = 0;

            using (var editDoc = PresentationDocument.Open(editMs, true))
            {
                var slideIdList = editDoc.PresentationPart?.Presentation?.SlideIdList;
                if (slideIdList == null)
                    return null;

                var allIds = slideIdList.Elements<P.SlideId>().ToList();
                if (options.SelectedSlideNumbers.Count > 0)
                {
                    int kept = 0;
                    for (int i = 0; i < allIds.Count; i++)
                    {
                        if (options.SelectedSlideNumbers.Contains(i + 1))
                        {
                            kept++;
                        }
                        else
                        {
                            allIds[i].Remove();
                        }
                    }

                    if (kept == 0)
                    {
                        return ExecutionResult.Fail("Please select at least one slide to convert.");
                    }

                    slideCount = kept;
                }
                else
                {
                    slideCount = allIds.Count;
                    if (slideCount == 0)
                        return ExecutionResult.Fail("Presentation contains no slides.");
                }

                // Pre-clip any <p:pic> shapes that use direct or SlideLayout-inherited custom/preset clip masks
                // (e.g. hexagon/ellipse/custom polygon picture placeholders) into transparent PNGs before OfficeIMO renders.
                PreClipMaskedPicturesInMemory(editDoc, progress);

                editDoc.PresentationPart!.Presentation!.Save();
            }

            sourceBytes = editMs.ToArray();

            progress?.Report((slideCount, slideCount, "rendering"));

            using var inStream = new MemoryStream(sourceBytes, writable: false);
            using var pptDoc = PowerPointDocument.Load(inStream, new PowerPointLoadOptions());

            var exportOptions = new PowerPointPdfExportOptions
            {
                IncludePictures = true,
                IncludeAutoShapes = true,
                IncludeTextBoxes = true,
                IncludeSlideBackgrounds = true,
                IncludeTables = true,
                IncludeCharts = true,
                IncludeSmartArt = true,
                IncludeSpeakerNotes = options.IncludeSpeakerNotes,
                PageLayout = options.IncludeSpeakerNotes
                    ? PowerPointPdfPageLayout.NotesPages
                    : PowerPointPdfPageLayout.Slides
            };

            using var outStream = new MemoryStream();
            pptDoc.SaveAsPdf(outStream, exportOptions);

            byte[] pdfBytes = outStream.ToArray();
            if (pdfBytes.Length == 0)
                return null;

            progress?.Report((slideCount, slideCount, "optimizing"));
            pdfBytes = OptimizePdfSizeInMemory(pdfBytes);

            sw.Stop();
            string docTitle = Path.GetFileNameWithoutExtension(document.FileName);
            string fileName = $"{docTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            string summary = $"Converted {slideCount} PowerPoint slide(s) to PDF via OfficeIMO Engine in {sw.ElapsedMilliseconds} ms";
            return ExecutionResult.Ok(pdfBytes, fileName, summary, sw.ElapsedMilliseconds);
        }
        catch
        {
            // Fall back cleanly to the built-in SkiaSharp renderer if OfficeIMO cannot parse a specific construct
            return null;
        }
    }

    private static void PreClipMaskedPicturesInMemory(
        PresentationDocument presDoc,
        IProgress<(int Current, int Total, string Stage)>? progress = null)
    {
        var presPart = presDoc.PresentationPart;
        var slideIds = presPart?.Presentation?.SlideIdList?.Elements<P.SlideId>().ToList();
        if (presPart == null || slideIds == null) return;

        // 1. Pre-rasterize decorative (non-placeholder, text-less) <a:custGeom> & hexagon/polygon shapes
        //    on SlideMasterParts and SlideLayoutParts (e.g. slideLayout16.xml hexagons on the Thank You slide)
        var processedMasters = new HashSet<SlideMasterPart>();
        var processedLayouts = new HashSet<SlideLayoutPart>();

        int idx = 0;
        foreach (var sId in slideIds)
        {
            idx++;
            progress?.Report((idx, slideIds.Count, "preparing"));
            string? relId = sId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relId) || presPart.GetPartById(relId) is not SlidePart slidePart)
                continue;

            var layoutPart = slidePart.SlideLayoutPart;
            var masterPart = layoutPart?.SlideMasterPart ?? presPart.SlideMasterParts.FirstOrDefault();
            var ctx = BuildSlideContext(presPart, slidePart, 960f, 540f, 1);

            if (masterPart != null && processedMasters.Add(masterPart))
            {
                var masterTree = masterPart.SlideMaster?.CommonSlideData?.ShapeTree;
                bool mMod = PreRasterizeCustomShapesInTree(masterPart, masterTree, ctx, skipPlaceholders: true);
                mMod |= PreClipMaskedPicturesInTree(masterPart, masterTree, null, null);
                if (mMod) masterPart.SlideMaster?.Save();
            }

            if (layoutPart != null && processedLayouts.Add(layoutPart))
            {
                var layoutTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree;
                var mShapes = masterPart?.SlideMaster?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
                bool lMod = PreRasterizeCustomShapesInTree(layoutPart, layoutTree, ctx, skipPlaceholders: true);
                lMod |= PreClipMaskedPicturesInTree(layoutPart, layoutTree, null, mShapes);
                if (lMod) layoutPart.SlideLayout?.Save();
            }

            // Also pre-rasterize any text-less custom/polygon shapes and masked pictures directly on the Slide
            var spTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (spTree == null) continue;

            bool slideModified = PreRasterizeCustomShapesInTree(slidePart, spTree, ctx, skipPlaceholders: false);

            var layoutShapes = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
            var masterShapes = masterPart?.SlideMaster?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
            slideModified |= PreClipMaskedPicturesInTree(slidePart, spTree, layoutShapes, masterShapes);

            if (slideModified)
            {
                slidePart.Slide!.Save();
            }
        }
    }

    private static bool PreClipMaskedPicturesInTree(
        OpenXmlPart ownerPart,
        P.ShapeTree? spTree,
        List<P.Shape>? layoutShapes,
        List<P.Shape>? masterShapes)
    {
        if (spTree == null) return false;
        bool modified = false;

        foreach (var pic in spTree.Descendants<P.Picture>().ToList())
        {
            var (custGeom, prstGeom, inheritedXfrm) = ResolvePictureClipGeometry(pic, layoutShapes, masterShapes);
            bool hasCustomClip = custGeom != null ||
                                 (prstGeom?.Preset?.Value != null && prstGeom.Preset.Value != A.ShapeTypeValues.Rectangle);
            if (!hasCustomClip) continue;

            string? embedId = pic.BlipFill?.Blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) continue;
            if (ownerPart.GetPartById(embedId) is not ImagePart srcImgPart) continue;

            try
            {
                byte[] rawImgBytes;
                using (var s = srcImgPart.GetStream())
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    rawImgBytes = ms.ToArray();
                }

                using var skImg = SKImage.FromEncodedData(rawImgBytes);
                if (skImg == null || skImg.Width <= 0 || skImg.Height <= 0) continue;

                // Compute target output dimensions based on picture (or inherited placeholder) extents
                var xfrm = pic.ShapeProperties?.Transform2D ?? inheritedXfrm;
                float boxW = xfrm?.Extents?.Cx != null ? (float)(xfrm.Extents.Cx.Value / EmusPerPoint) : skImg.Width;
                float boxH = xfrm?.Extents?.Cy != null ? (float)(xfrm.Extents.Cy.Value / EmusPerPoint) : skImg.Height;
                if (boxW <= 1f || boxH <= 1f) continue;

                float scale = Math.Min(2.2f, 1200f / Math.Max(boxW, boxH));
                int outW = Math.Clamp((int)Math.Round(boxW * scale), 32, 1400);
                int outH = Math.Clamp((int)Math.Round(boxH * scale), 32, 1400);

                var destRect = new SKRect(0, 0, outW, outH);
                using var clipPath = BuildShapePath(destRect, custGeom, prstGeom);

                using var surface = SKSurface.Create(new SKImageInfo(outW, outH, SKColorType.Rgba8888, SKAlphaType.Premul));
                if (surface == null) continue;

                var canvas = surface.Canvas;
                canvas.Clear(SKColors.Transparent);
                canvas.ClipPath(clipPath, SKClipOperation.Intersect, antialias: true);

                // Apply <a:srcRect> crop if present
                var srcRectEl = pic.BlipFill?.SourceRectangle;
                if (srcRectEl != null &&
                    (srcRectEl.Left?.Value != null || srcRectEl.Top?.Value != null || srcRectEl.Right?.Value != null || srcRectEl.Bottom?.Value != null))
                {
                    float cL = (srcRectEl.Left?.Value ?? 0) / 100000f;
                    float cT = (srcRectEl.Top?.Value ?? 0) / 100000f;
                    float cR = (srcRectEl.Right?.Value ?? 0) / 100000f;
                    float cB = (srcRectEl.Bottom?.Value ?? 0) / 100000f;

                    float sx1 = Math.Clamp(cL * skImg.Width, 0f, skImg.Width - 1f);
                    float sy1 = Math.Clamp(cT * skImg.Height, 0f, skImg.Height - 1f);
                    float sx2 = Math.Clamp((1f - cR) * skImg.Width, sx1 + 1f, skImg.Width);
                    float sy2 = Math.Clamp((1f - cB) * skImg.Height, sy1 + 1f, skImg.Height);

                    canvas.DrawImage(skImg, new SKRect(sx1, sy1, sx2, sy2), destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
                }
                else
                {
                    canvas.DrawImage(skImg, destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
                }

                using var snap = surface.Snapshot();
                using var encodedPng = snap.Encode(SKEncodedImageFormat.Png, 95);
                if (encodedPng == null) continue;

                ImagePart newImgPart = ownerPart switch
                {
                    SlidePart sp => sp.AddImagePart(ImagePartType.Png),
                    SlideLayoutPart lp => lp.AddImagePart(ImagePartType.Png),
                    SlideMasterPart mp => mp.AddImagePart(ImagePartType.Png),
                    _ => null!
                };
                if (newImgPart == null) continue;

                using (var pngStream = new MemoryStream(encodedPng.ToArray(), writable: false))
                {
                    newImgPart.FeedData(pngStream);
                }

                string newRelId = ownerPart.GetIdOfPart(newImgPart);
                pic.BlipFill!.Blip!.Embed = newRelId;
                pic.BlipFill.SourceRectangle?.Remove();

                // Replace geometry with rectangle since the mask is now baked into the transparent PNG alpha channel
                pic.ShapeProperties?.GetFirstChild<A.CustomGeometry>()?.Remove();
                pic.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Remove();
                pic.ShapeProperties?.AppendChild(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

                modified = true;
            }
            catch
            {
                // Skip problematic picture and continue
            }
        }

        return modified;
    }

    /// <summary>
    /// Replaces text-less decorative shapes using CustomGeometry or polygon PresetGeometry (Hexagon, Diamond, etc.)
    /// with crisp pre-rendered transparent PNG Picture elements so OfficeIMO exports their exact polygon silhouette.
    /// </summary>
    private static bool PreRasterizeCustomShapesInTree(
        OpenXmlPart ownerPart,
        OpenXmlCompositeElement? container,
        SlideRenderContext ctx,
        bool skipPlaceholders,
        float parentScaleX = 1f,
        float parentScaleY = 1f)
    {
        if (container == null) return false;
        bool modified = false;

        foreach (var child in container.ChildElements.ToList())
        {
            if (child is P.GroupShape grp)
            {
                float gx = parentScaleX;
                float gy = parentScaleY;
                var gxfrm = grp.GroupShapeProperties?.TransformGroup;
                if (gxfrm?.Extents?.Cx?.Value > 0 && gxfrm.ChildExtents?.Cx?.Value > 0)
                    gx *= (float)gxfrm.Extents.Cx.Value / gxfrm.ChildExtents.Cx.Value;
                if (gxfrm?.Extents?.Cy?.Value > 0 && gxfrm.ChildExtents?.Cy?.Value > 0)
                    gy *= (float)gxfrm.Extents.Cy.Value / gxfrm.ChildExtents.Cy.Value;

                modified |= PreRasterizeCustomShapesInTree(ownerPart, grp, ctx, skipPlaceholders, gx, gy);
                continue;
            }

            if (child is not P.Shape shape)
                continue;

            if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value == true)
                continue;

            bool isPh = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() != null;
            if (isPh && skipPlaceholders)
                continue;

            // Keep shapes that contain visible text as native vector text boxes
            if (!string.IsNullOrWhiteSpace(ExtractShapeText(shape)))
                continue;

            var spPr = shape.ShapeProperties;
            var xfrm = spPr?.Transform2D;
            if (xfrm?.Offset == null || xfrm.Extents == null)
                continue;

            var custGeom = spPr?.GetFirstChild<A.CustomGeometry>();
            var prstGeom = spPr?.GetFirstChild<A.PresetGeometry>();
            var prst = prstGeom?.Preset?.Value;
            bool hasBlipFill = spPr?.GetFirstChild<A.BlipFill>() != null;

            bool needsRaster = custGeom != null
                               || (hasBlipFill && prst != null && prst != A.ShapeTypeValues.Rectangle)
                               || IsPolygonOrArrowPreset(prst);

            if (!needsRaster)
                continue;

            long rawX = xfrm.Offset.X?.Value ?? 0L;
            long rawY = xfrm.Offset.Y?.Value ?? 0L;
            long rawCx = xfrm.Extents.Cx?.Value ?? 0L;
            long rawCy = xfrm.Extents.Cy?.Value ?? 0L;
            if (rawCx <= 0 || rawCy <= 0) continue;

            float rotDeg = xfrm.Rotation != null ? xfrm.Rotation.Value / 60000f : 0f;
            bool flipH = xfrm.HorizontalFlip?.Value == true;
            bool flipV = xfrm.VerticalFlip?.Value == true;

            // Compute axis-aligned bounding box in child space AFTER rotation so we bake both rotation
            // and flips into the PNG (preserving PowerPoint's exact Rotate -> GroupScale order).
            double rad = rotDeg * Math.PI / 180.0;
            double absCos = Math.Abs(Math.Cos(rad));
            double absSin = Math.Abs(Math.Sin(rad));
            long bboxCx = Math.Max(1L, (long)Math.Round(rawCx * absCos + rawCy * absSin));
            long bboxCy = Math.Max(1L, (long)Math.Round(rawCx * absSin + rawCy * absCos));
            long bboxX = (long)Math.Round(rawX + rawCx * 0.5 - bboxCx * 0.5);
            long bboxY = (long)Math.Round(rawY + rawCy * 0.5 - bboxCy * 0.5);

            float boxW = Math.Abs((float)(bboxCx / EmusPerPoint) * parentScaleX);
            float boxH = Math.Abs((float)(bboxCy / EmusPerPoint) * parentScaleY);
            if (boxW <= 1f || boxH <= 1f) continue;

            try
            {
                float scale = Math.Min(3.0f, 900f / Math.Max(boxW, boxH));
                int outW = Math.Clamp((int)Math.Round(boxW * scale), 24, 1000);
                int outH = Math.Clamp((int)Math.Round(boxH * scale), 24, 1000);

                // Resolve outline stroke width first so we can inset the path slightly to avoid clipping the stroke at edges
                var outline = spPr?.GetFirstChild<A.Outline>();
                SKColor? strokeColor = null;
                float strokeWidthPt = 1f;
                A.PresetLineDashValues? dashPreset = null;

                if (outline != null)
                {
                    if (outline.GetFirstChild<A.NoFill>() == null)
                    {
                        strokeColor = ResolveColorFromContainer(outline.GetFirstChild<A.SolidFill>(), ctx.ThemeColors, null);
                        if (!strokeColor.HasValue && shape.ShapeStyle?.LineReference != null && (shape.ShapeStyle.LineReference.Index?.Value ?? 0U) > 0)
                        {
                            strokeColor = ResolveColorFromContainer(shape.ShapeStyle.LineReference, ctx.ThemeColors, null);
                        }
                        if (outline.Width?.Value != null)
                        {
                            strokeWidthPt = Math.Max(0.75f, (float)(outline.Width.Value / EmusPerPoint));
                        }
                        dashPreset = outline.GetFirstChild<A.PresetDash>()?.Val?.Value;
                    }
                }
                else if (shape.ShapeStyle?.LineReference != null && (shape.ShapeStyle.LineReference.Index?.Value ?? 0U) > 0)
                {
                    strokeColor = ResolveColorFromContainer(shape.ShapeStyle.LineReference, ctx.ThemeColors, null);
                }

                float strokeInset = strokeColor.HasValue ? Math.Max(1.5f, (strokeWidthPt * scale) * 0.6f) : 0.5f;
                var drawRect = new SKRect(strokeInset, strokeInset, outW - strokeInset, outH - strokeInset);

                // Build shape in normalized local coordinates centered at (0, 0), then transform:
                // Flip -> Rotate (in child space) -> Scale to pixel bbox (applying group aspect ratio) -> Translate to center
                float localW = (float)(rawCx / EmusPerPoint);
                float localH = (float)(rawCy / EmusPerPoint);
                float bboxPtW = (float)(bboxCx / EmusPerPoint);
                float bboxPtH = (float)(bboxCy / EmusPerPoint);

                using var path = BuildShapePath(new SKRect(-localW * 0.5f, -localH * 0.5f, localW * 0.5f, localH * 0.5f), custGeom, prstGeom);

                var m = SKMatrix.CreateIdentity();
                if (flipH || flipV)
                    m = SKMatrix.Concat(SKMatrix.CreateScale(flipH ? -1f : 1f, flipV ? -1f : 1f), m);
                if (Math.Abs(rotDeg) > 0.01f)
                    m = SKMatrix.Concat(SKMatrix.CreateRotationDegrees(rotDeg), m);
                m = SKMatrix.Concat(SKMatrix.CreateScale(drawRect.Width / bboxPtW, drawRect.Height / bboxPtH), m);
                m = SKMatrix.Concat(SKMatrix.CreateTranslation(drawRect.MidX, drawRect.MidY), m);
                path.Transform(m);

                using var surface = SKSurface.Create(new SKImageInfo(outW, outH, SKColorType.Rgba8888, SKAlphaType.Premul));
                if (surface == null) continue;

                var canvas = surface.Canvas;
                canvas.Clear(SKColors.Transparent);

                // Draw Fill
                bool hasNoFill = spPr?.GetFirstChild<A.NoFill>() != null;
                bool drewFill = false;
                if (!hasNoFill && spPr != null && HasAnyFillChild(spPr))
                {
                    drewFill = TryDrawFillFromContainer(canvas, spPr, ownerPart, ctx.ThemeColors, drawRect, null, path, out _);
                }

                if (!hasNoFill && !drewFill && shape.ShapeStyle?.FillReference != null && (shape.ShapeStyle.FillReference.Index?.Value ?? 0U) > 0)
                {
                    var refClr = ResolveColorFromContainer(shape.ShapeStyle.FillReference, ctx.ThemeColors, null);
                    if (refClr.HasValue)
                    {
                        using var fPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = refClr.Value, IsAntialias = true };
                        canvas.DrawPath(path, fPaint);
                        drewFill = true;
                    }
                }

                // Draw Stroke (with optional dash style)
                if (strokeColor.HasValue)
                {
                    float swPx = Math.Max(1f, strokeWidthPt * scale);
                    using var dashEffect = CreateDashPathEffect(dashPreset, swPx);
                    using var sPaint = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = strokeColor.Value,
                        StrokeWidth = swPx,
                        StrokeJoin = SKStrokeJoin.Miter,
                        PathEffect = dashEffect,
                        IsAntialias = true
                    };
                    canvas.DrawPath(path, sPaint);
                }

                if (!drewFill && !strokeColor.HasValue)
                    continue;

                using var snap = surface.Snapshot();
                using var encodedPng = snap.Encode(SKEncodedImageFormat.Png, 95);
                if (encodedPng == null) continue;

                ImagePart newImgPart = ownerPart switch
                {
                    SlidePart sp => sp.AddImagePart(ImagePartType.Png),
                    SlideLayoutPart lp => lp.AddImagePart(ImagePartType.Png),
                    SlideMasterPart mp => mp.AddImagePart(ImagePartType.Png),
                    _ => null!
                };
                if (newImgPart == null) continue;

                using (var pngMs = new MemoryStream(encodedPng.ToArray(), writable: false))
                {
                    newImgPart.FeedData(pngMs);
                }

                string imgRelId = ownerPart.GetIdOfPart(newImgPart);
                uint origId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 900U;
                string origName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Shape";

                var clonedXfrm = new A.Transform2D(
                    new A.Offset { X = bboxX, Y = bboxY },
                    new A.Extents { Cx = bboxCx, Cy = bboxCy });

                var replacementPic = new P.Picture(
                    new P.NonVisualPictureProperties(
                        new P.NonVisualDrawingProperties { Id = origId, Name = origName },
                        new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.BlipFill(
                        new A.Blip { Embed = imgRelId },
                        new A.Stretch(new A.FillRectangle())),
                    new P.ShapeProperties(
                        clonedXfrm,
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

                shape.InsertAfterSelf(replacementPic);
                shape.Remove();
                modified = true;
            }
            catch
            {
                // Skip and leave original shape intact
            }
        }

        return modified;
    }

    private static bool IsPolygonOrArrowPreset(A.ShapeTypeValues? prst)
    {
        if (prst == null) return false;
        return prst == A.ShapeTypeValues.Hexagon
               || prst == A.ShapeTypeValues.Diamond
               || prst == A.ShapeTypeValues.Triangle
               || prst == A.ShapeTypeValues.RightTriangle
               || prst == A.ShapeTypeValues.Parallelogram
               || prst == A.ShapeTypeValues.Trapezoid
               || prst == A.ShapeTypeValues.Pentagon
               || prst == A.ShapeTypeValues.Heptagon
               || prst == A.ShapeTypeValues.Octagon
               || prst == A.ShapeTypeValues.Decagon
               || prst == A.ShapeTypeValues.Dodecagon
               || prst == A.ShapeTypeValues.Chevron
               || prst == A.ShapeTypeValues.HomePlate
               || prst == A.ShapeTypeValues.RightArrow
               || prst == A.ShapeTypeValues.LeftArrow
               || prst == A.ShapeTypeValues.UpArrow
               || prst == A.ShapeTypeValues.DownArrow
               || prst == A.ShapeTypeValues.Star4
               || prst == A.ShapeTypeValues.Star5
               || prst == A.ShapeTypeValues.Star6
               || prst == A.ShapeTypeValues.Star8
               || prst == A.ShapeTypeValues.Plus
               || prst == A.ShapeTypeValues.Donut
               || prst == A.ShapeTypeValues.Frame
               || prst == A.ShapeTypeValues.HalfFrame
               || prst == A.ShapeTypeValues.Corner
               || prst == A.ShapeTypeValues.Heart
               || prst == A.ShapeTypeValues.Teardrop;
    }

    private static SKPathEffect? CreateDashPathEffect(A.PresetLineDashValues? dash, float sw)
    {
        if (dash == null || dash == A.PresetLineDashValues.Solid) return null;
        if (dash == A.PresetLineDashValues.Dot || dash == A.PresetLineDashValues.SystemDot)
            return SKPathEffect.CreateDash(new[] { sw, sw * 2f }, 0f);
        if (dash == A.PresetLineDashValues.Dash || dash == A.PresetLineDashValues.SystemDash)
            return SKPathEffect.CreateDash(new[] { sw * 4f, sw * 3f }, 0f);
        if (dash == A.PresetLineDashValues.LargeDash)
            return SKPathEffect.CreateDash(new[] { sw * 8f, sw * 3f }, 0f);
        if (dash == A.PresetLineDashValues.DashDot || dash == A.PresetLineDashValues.SystemDashDot)
            return SKPathEffect.CreateDash(new[] { sw * 4f, sw * 2f, sw, sw * 2f }, 0f);
        return null;
    }

    private static (A.CustomGeometry? CustGeom, A.PresetGeometry? PrstGeom, A.Transform2D? InheritedXfrm) ResolvePictureClipGeometry(
        P.Picture pic,
        List<P.Shape>? layoutShapes,
        List<P.Shape>? masterShapes)
    {
        var ph = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
        P.Shape? layoutMatch = ph != null ? MatchPlaceholderInList(ph, layoutShapes, isMasterLookup: false) : null;
        P.Shape? masterMatch = ph != null ? MatchPlaceholderInList(ph, masterShapes, isMasterLookup: true) : null;
        var inhXfrm = layoutMatch?.ShapeProperties?.Transform2D ?? masterMatch?.ShapeProperties?.Transform2D;

        var directCust = pic.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
        var directPrst = pic.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (directCust != null || (directPrst?.Preset?.Value != null && directPrst.Preset.Value != A.ShapeTypeValues.Rectangle))
        {
            return (directCust, directPrst, inhXfrm);
        }

        if (ph != null)
        {
            var inhCust = layoutMatch?.ShapeProperties?.GetFirstChild<A.CustomGeometry>()
                          ?? masterMatch?.ShapeProperties?.GetFirstChild<A.CustomGeometry>();
            var inhPrst = layoutMatch?.ShapeProperties?.GetFirstChild<A.PresetGeometry>()
                          ?? masterMatch?.ShapeProperties?.GetFirstChild<A.PresetGeometry>();

            return (inhCust, inhPrst, inhXfrm);
        }

        return (directCust, directPrst, inhXfrm);
    }

    /// <summary>
    /// Compresses embedded RGB raster images inside the generated PDF using DCTDecode (JPEG)
    /// and enables best Flate compression, cutting file size dramatically while keeping vector text crisp.
    /// </summary>
    private static byte[] OptimizePdfSizeInMemory(byte[] pdfBytes)
    {
        try
        {
            using var inMs = new MemoryStream(pdfBytes, writable: false);
            using var pdfDoc = PdfSharp.Pdf.IO.PdfReader.Open(inMs, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
            pdfDoc.Options.FlateEncodeMode = PdfSharp.Pdf.PdfFlateEncodeMode.BestCompression;
            pdfDoc.Options.NoCompression = false;

            var visited = new HashSet<PdfSharp.Pdf.PdfDictionary>();

            foreach (var page in pdfDoc.Pages)
            {
                var xObjects = page.Elements.GetDictionary("/Resources")?.Elements.GetDictionary("/XObject");
                if (xObjects == null) continue;

                foreach (var key in xObjects.Elements.Keys.ToList())
                {
                    var xObjRef = xObjects.Elements[key] as PdfSharp.Pdf.Advanced.PdfReference;
                    var xObj = (xObjRef?.Value as PdfSharp.Pdf.PdfDictionary) ?? xObjects.Elements.GetDictionary(key);
                    if (xObj == null || !visited.Add(xObj)) continue;

                    var subtype = xObj.Elements.GetName("/Subtype");
                    if (!string.Equals(subtype, "/Image", StringComparison.Ordinal)) continue;
                    if (xObj.Stream == null || xObj.Stream.Value == null || xObj.Stream.Value.Length < 16_384) continue;

                    // Skip images already JPEG-compressed or indexed/1-bit masks
                    var filter = xObj.Elements.GetName("/Filter");
                    if (string.Equals(filter, "/DCTDecode", StringComparison.Ordinal)) continue;

                    // Keep small/medium transparent PNGs (with /SMask) lossless so vector-like icons & shapes stay sharp
                    if (xObj.Elements.ContainsKey("/SMask") && xObj.Stream.Value.Length < 65_536)
                        continue;

                    int width = xObj.Elements.GetInteger("/Width");
                    int height = xObj.Elements.GetInteger("/Height");
                    int bpc = xObj.Elements.GetInteger("/BitsPerComponent");
                    var colorSpace = xObj.Elements.GetName("/ColorSpace");

                    if (width <= 0 || height <= 0 || bpc != 8 || !string.Equals(colorSpace, "/DeviceRGB", StringComparison.Ordinal))
                        continue;

                    byte[] rawRgb = xObj.Stream.UnfilteredValue;
                    int expectedLen = width * height * 3;
                    if (rawRgb == null || rawRgb.Length < expectedLen) continue;

                    using var bmp = new SKBitmap(width, height, SKColorType.Rgba8888, SKAlphaType.Opaque);
                    var pixels = new uint[width * height];
                    for (int i = 0, p = 0; p < pixels.Length; i += 3, p++)
                    {
                        byte r = rawRgb[i];
                        byte g = rawRgb[i + 1];
                        byte b = rawRgb[i + 2];
                        pixels[p] = (uint)((255 << 24) | (b << 16) | (g << 8) | r);
                    }

                    System.Runtime.InteropServices.Marshal.Copy(Array.ConvertAll(pixels, x => (int)x), 0, bmp.GetPixels(), pixels.Length);

                    using var img = SKImage.FromBitmap(bmp);
                    using var jpgData = img.Encode(SKEncodedImageFormat.Jpeg, 82);
                    if (jpgData == null) continue;

                    byte[] jpgBytes = jpgData.ToArray();
                    if (jpgBytes.Length < xObj.Stream.Value.Length)
                    {
                        xObj.Stream.Value = jpgBytes;
                        xObj.Elements["/Filter"] = new PdfSharp.Pdf.PdfName("/DCTDecode");
                        xObj.Elements["/Length"] = new PdfSharp.Pdf.PdfInteger(jpgBytes.Length);
                        xObj.Elements.Remove("/DecodeParms");
                    }
                }
            }

            using var outMs = new MemoryStream();
            pdfDoc.Save(outMs, false);
            byte[] optimized = outMs.ToArray();
            return optimized.Length > 0 && optimized.Length < pdfBytes.Length ? optimized : pdfBytes;
        }
        catch
        {
            return pdfBytes;
        }
    }

    private void RenderSlideToCanvas(
        SKCanvas canvas,
        SlideRenderContext ctx,
        PowerPointPdfOptions options)
    {
        float slideW = ctx.SlideWidth;
        float slideH = ctx.SlideHeight;
        var slideRect = new SKRect(0, 0, slideW, slideH);

        canvas.Save();
        canvas.ClipRect(slideRect);

        // 1. Render Slide Background (Slide -> SlideLayout -> SlideMaster)
        RenderEffectiveBackground(canvas, ctx, slideRect);

        // 2. Render Non-Placeholder Master & Layout Decorative Shapes/Logos
        bool showLayoutMasterShapes = ctx.SlidePart.Slide?.ShowMasterShapes?.Value ?? true;
        bool showMasterOnLayout = ctx.LayoutPart?.SlideLayout?.ShowMasterShapes?.Value ?? true;

        if (showLayoutMasterShapes && showMasterOnLayout && ctx.MasterPart?.SlideMaster?.CommonSlideData?.ShapeTree != null)
        {
            foreach (var masterChild in ctx.MasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements())
            {
                if (IsPlaceholderElement(masterChild))
                    continue;
                RenderShapeTreeElement(canvas, ctx, ctx.MasterPart, masterChild, isMasterOrLayoutLayer: true, 0f, 0f, 1f, 1f);
            }
        }

        if (showLayoutMasterShapes && ctx.LayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree != null)
        {
            foreach (var layoutChild in ctx.LayoutPart.SlideLayout.CommonSlideData.ShapeTree.Elements())
            {
                if (IsPlaceholderElement(layoutChild))
                    continue;
                RenderShapeTreeElement(canvas, ctx, ctx.LayoutPart, layoutChild, isMasterOrLayoutLayer: true, 0f, 0f, 1f, 1f);
            }
        }

        // 3. Render Slide ShapeTree Elements
        var spTree = ctx.SlidePart.Slide?.CommonSlideData?.ShapeTree;
        if (spTree != null)
        {
            foreach (var child in spTree.Elements())
            {
                RenderShapeTreeElement(canvas, ctx, ctx.SlidePart, child, isMasterOrLayoutLayer: false, 0f, 0f, 1f, 1f);
            }
        }

        // 4. Optional Slide Number Badge
        if (options.IncludeSlideNumbers)
        {
            using var numTf = _fontService.ResolveOsTypeface(ctx.MinorLatinFont, isBold: false, isItalic: false);
            using var numFont = new SKFont(numTf, 9.5f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var numPaint = new SKPaint { IsAntialias = true, Color = new SKColor(100, 116, 139, 200) };
            string label = ctx.SlideNumber.ToString(CultureInfo.InvariantCulture);
            float w = numFont.MeasureText(label);
            canvas.DrawText(label, slideW - w - 18f, slideH - 12f, SKTextAlign.Left, numFont, numPaint);
        }

        canvas.Restore();

        // 5. Optional Speaker Notes Band below slide
        if (options.IncludeSpeakerNotes)
        {
            float notesY = slideH;
            using var bandBg = new SKPaint { Style = SKPaintStyle.Fill, Color = new SKColor(248, 250, 252), IsAntialias = true };
            using var bandBorder = new SKPaint { Style = SKPaintStyle.Stroke, Color = new SKColor(203, 213, 225), StrokeWidth = 1f, IsAntialias = true };
            canvas.DrawRect(new SKRect(0, notesY, slideW, notesY + 86f), bandBg);
            canvas.DrawLine(0, notesY, slideW, notesY, bandBorder);

            string notesText = ctx.SlidePart.NotesSlidePart != null
                ? ExtractNotesText(ctx.SlidePart.NotesSlidePart)
                : string.Empty;

            if (string.IsNullOrWhiteSpace(notesText))
            {
                notesText = $"Slide {ctx.SlideNumber} — No speaker notes";
            }

            bool isRtl = OsFontService.ContainsArabic(notesText);
            using var notesTf = _fontService.ResolveOsTypeface(ctx.MinorLatinFont, isBold: false, isItalic: false, requiresArabicGlyphs: isRtl);
            using var notesFont = new SKFont(notesTf, 10f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
            using var notesShaper = new SKShaper(notesTf);
            using var notesPaint = new SKPaint { IsAntialias = true, Color = new SKColor(51, 65, 85) };

            var lines = WrapPlainTextLines(notesText, notesShaper, notesFont, 10f, slideW - 48f);
            float curY = notesY + 20f;
            foreach (var line in lines.Take(4))
            {
                DrawPlainVisualLine(canvas, line, notesShaper, notesFont, notesPaint, 10f, isRtl, 24f, slideW - 48f, notesFont.MeasureText(" "), curY, A.TextAlignmentTypeValues.Left);
                curY += 15f;
            }
        }
    }

    private static bool IsPlaceholderElement(OpenXmlElement el)
    {
        if (el is P.Shape sp)
        {
            return sp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() != null;
        }
        return false;
    }

    // =========================================================================
    // 3. Background Rendering (Solid, Gradient, Image, & bgRef)
    // =========================================================================

    private static void RenderEffectiveBackground(SKCanvas canvas, SlideRenderContext ctx, SKRect slideRect)
    {
        canvas.Clear(SKColors.White);

        // Find the first defined P.Background in Slide -> SlideLayout -> SlideMaster
        (P.Background? bg, OpenXmlPart? ownerPart) = FindEffectiveBackground(ctx);
        if (bg == null || ownerPart == null)
        {
            // Default to theme bg1 if present
            if (ctx.ThemeColors.TryGetValue("bg1", out var bg1))
            {
                canvas.Clear(bg1);
            }
            return;
        }

        var bgPr = bg.BackgroundProperties;
        if (bgPr != null)
        {
            if (TryDrawFillFromContainer(canvas, bgPr, ownerPart, ctx.ThemeColors, slideRect, null))
                return;
        }

        var bgRef = bg.BackgroundStyleReference;
        if (bgRef != null)
        {
            SKColor? refColor = ResolveColorFromContainer(bgRef, ctx.ThemeColors, null);
            var bgFillList = ctx.ThemePart?.Theme?.ThemeElements?.FormatScheme?.BackgroundFillStyleList;
            uint idx = bgRef.Index?.Value ?? 0U;
            if (bgFillList != null && idx > 0)
            {
                int listIdx = (int)(idx >= 1001 ? idx - 1001 : idx - 1);
                var fillElements = bgFillList.ChildElements.ToList();
                if (listIdx >= 0 && listIdx < fillElements.Count)
                {
                    if (TryDrawFillElement(canvas, fillElements[listIdx], ownerPart, ctx.ThemeColors, slideRect, refColor))
                        return;
                }
            }

            if (refColor.HasValue)
            {
                canvas.Clear(refColor.Value);
            }
        }
    }

    private static (P.Background? Bg, OpenXmlPart? Owner) FindEffectiveBackground(SlideRenderContext ctx)
    {
        var slideBg = ctx.SlidePart.Slide?.CommonSlideData?.Background;
        if (slideBg != null && (slideBg.BackgroundProperties != null || slideBg.BackgroundStyleReference != null))
            return (slideBg, ctx.SlidePart);

        var layoutBg = ctx.LayoutPart?.SlideLayout?.CommonSlideData?.Background;
        if (layoutBg != null && (layoutBg.BackgroundProperties != null || layoutBg.BackgroundStyleReference != null))
            return (layoutBg, ctx.LayoutPart);

        var masterBg = ctx.MasterPart?.SlideMaster?.CommonSlideData?.Background;
        if (masterBg != null && (masterBg.BackgroundProperties != null || masterBg.BackgroundStyleReference != null))
            return (masterBg, ctx.MasterPart);

        return (null, null);
    }

    // =========================================================================
    // 4. Placeholder Inheritance Lookup (Slide -> SlideLayout -> SlideMaster)
    // =========================================================================

    private static (P.Shape? LayoutShape, P.Shape? MasterShape) FindInheritedPlaceholderShapes(P.Shape slideShape, SlidePart slidePart)
    {
        var ph = slideShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
        if (ph == null)
            return (null, null);

        var layoutShapes = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
        var masterShapes = (slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree)
            ?.Descendants<P.Shape>().ToList();

        P.Shape? matchedLayout = MatchPlaceholderInList(ph, layoutShapes, isMasterLookup: false);
        var effectivePh = matchedLayout?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() ?? ph;
        P.Shape? matchedMaster = MatchPlaceholderInList(effectivePh, masterShapes, isMasterLookup: true);

        return (matchedLayout, matchedMaster);
    }

    private static P.Shape? MatchPlaceholderInList(P.PlaceholderShape targetPh, List<P.Shape>? candidates, bool isMasterLookup)
    {
        if (candidates == null || candidates.Count == 0)
            return null;

        uint? idx = targetPh.Index?.Value;
        var type = targetPh.Type?.Value;

        // 1. Match by explicit Index first (when looking up in SlideLayout)
        if (!isMasterLookup && idx.HasValue)
        {
            foreach (var cand in candidates)
            {
                var cPh = cand.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                if (cPh?.Index?.Value == idx.Value)
                    return cand;
            }
        }

        // 2. Match by Placeholder Type
        if (type.HasValue)
        {
            var searchType = type.Value;
            if (isMasterLookup)
            {
                if (searchType == P.PlaceholderValues.CenteredTitle)
                    searchType = P.PlaceholderValues.Title;
                else if (searchType == P.PlaceholderValues.SubTitle || searchType == P.PlaceholderValues.Object)
                    searchType = P.PlaceholderValues.Body;
            }

            foreach (var cand in candidates)
            {
                var cPh = cand.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                if (cPh?.Type?.Value == searchType || cPh?.Type?.Value == type.Value)
                    return cand;
            }
        }

        // 3. Fallback: if targetPh had an index and no type, and we are looking in Master, match any Body placeholder
        if (isMasterLookup && idx.HasValue)
        {
            foreach (var cand in candidates)
            {
                var cPh = cand.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                if (cPh?.Index?.Value == idx.Value || cPh?.Type?.Value == P.PlaceholderValues.Body)
                    return cand;
            }
        }

        return null;
    }

    private static P.PlaceholderValues? GetEffectivePlaceholderType(P.Shape slideShape, P.Shape? layoutShape, P.Shape? masterShape = null)
    {
        return slideShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value
               ?? layoutShape?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value
               ?? masterShape?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value;
    }

    // =========================================================================
    // 5. Shape Tree Dispatch (Shapes, Connectors, Pictures, Tables, Groups)
    // =========================================================================

    private void RenderShapeTreeElement(
        SKCanvas canvas,
        SlideRenderContext ctx,
        OpenXmlPart ownerPart,
        OpenXmlElement element,
        bool isMasterOrLayoutLayer,
        float offsetX,
        float offsetY,
        float scaleX,
        float scaleY)
    {
        if (element is P.Shape shape)
        {
            RenderShape(canvas, ctx, ownerPart, shape, isMasterOrLayoutLayer, offsetX, offsetY, scaleX, scaleY);
        }
        else if (element is P.ConnectionShape connShape)
        {
            RenderConnectionShape(canvas, ctx, connShape, offsetX, offsetY, scaleX, scaleY);
        }
        else if (element is P.Picture pic)
        {
            RenderPicture(canvas, ctx, ownerPart, pic, offsetX, offsetY, scaleX, scaleY);
        }
        else if (element is P.GraphicFrame graphicFrame)
        {
            RenderGraphicFrame(canvas, ctx, graphicFrame, offsetX, offsetY, scaleX, scaleY);
        }
        else if (element is P.GroupShape group)
        {
            var xfrm = group.GroupShapeProperties?.TransformGroup;
            if (xfrm == null)
            {
                foreach (var child in group.Elements())
                {
                    RenderShapeTreeElement(canvas, ctx, ownerPart, child, isMasterOrLayoutLayer, offsetX, offsetY, scaleX, scaleY);
                }
                return;
            }

            float gx = offsetX + (xfrm.Offset?.X != null ? (float)(xfrm.Offset.X.Value / EmusPerPoint) * scaleX : 0f);
            float gy = offsetY + (xfrm.Offset?.Y != null ? (float)(xfrm.Offset.Y.Value / EmusPerPoint) * scaleY : 0f);
            float chX = xfrm.ChildOffset?.X != null ? (float)(xfrm.ChildOffset.X.Value / EmusPerPoint) : 0f;
            float chY = xfrm.ChildOffset?.Y != null ? (float)(xfrm.ChildOffset.Y.Value / EmusPerPoint) : 0f;

            float sx = scaleX;
            float sy = scaleY;
            if (xfrm.Extents?.Cx != null && xfrm.ChildExtents?.Cx != null && xfrm.ChildExtents.Cx.Value > 0)
            {
                sx *= (float)xfrm.Extents.Cx.Value / xfrm.ChildExtents.Cx.Value;
            }
            if (xfrm.Extents?.Cy != null && xfrm.ChildExtents?.Cy != null && xfrm.ChildExtents.Cy.Value > 0)
            {
                sy *= (float)xfrm.Extents.Cy.Value / xfrm.ChildExtents.Cy.Value;
            }

            foreach (var child in group.Elements())
            {
                RenderShapeTreeElement(canvas, ctx, ownerPart, child, isMasterOrLayoutLayer, gx - (chX * sx), gy - (chY * sy), sx, sy);
            }
        }
    }

    // =========================================================================
    // 6. Shape & Rich Text Rendering with Full Master/Layout Inheritance
    // =========================================================================

    private void RenderShape(
        SKCanvas canvas,
        SlideRenderContext ctx,
        OpenXmlPart ownerPart,
        P.Shape shape,
        bool isMasterOrLayoutLayer,
        float offsetX,
        float offsetY,
        float scaleX,
        float scaleY)
    {
        // Check hidden flag
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value == true)
            return;

        P.Shape? layoutShape = null;
        P.Shape? masterShape = null;
        if (!isMasterOrLayoutLayer)
        {
            (layoutShape, masterShape) = FindInheritedPlaceholderShapes(shape, ctx.SlidePart);
        }

        var phType = GetEffectivePlaceholderType(shape, layoutShape, masterShape);

        // Resolve Transform2D (Slide -> SlideLayout -> SlideMaster)
        var xfrm = GetEffectiveTransform2D(shape.ShapeProperties?.Transform2D, layoutShape?.ShapeProperties?.Transform2D, masterShape?.ShapeProperties?.Transform2D);
        var rect = GetTransformRect(xfrm, offsetX, offsetY, scaleX, scaleY);

        // If a placeholder still has no explicit transform (e.g. minimal synthesized presentation), give it a clean default rect
        if ((rect.Width <= 1f || rect.Height <= 1f) && phType.HasValue)
        {
            rect = GetFallbackPlaceholderRect(phType.Value, ctx.SlideWidth, ctx.SlideHeight);
        }

        if (rect.Width <= 1f || rect.Height <= 1f)
            return;

        // Apply rotation / flips if specified on Transform2D
        float rotDeg = xfrm?.Rotation != null ? xfrm.Rotation.Value / 60000f : 0f;
        bool flipH = xfrm?.HorizontalFlip?.Value == true;
        bool flipV = xfrm?.VerticalFlip?.Value == true;

        canvas.Save();
        if (Math.Abs(rotDeg) > 0.05f || flipH || flipV)
        {
            canvas.Translate(rect.MidX, rect.MidY);
            if (Math.Abs(rotDeg) > 0.05f) canvas.RotateDegrees(rotDeg);
            if (flipH || flipV) canvas.Scale(flipH ? -1f : 1f, flipV ? -1f : 1f);
            canvas.Translate(-rect.MidX, -rect.MidY);
        }

        // 1. Resolve Geometry (CustomGeometry or PresetGeometry)
        var spPr = shape.ShapeProperties;
        var layoutSpPr = layoutShape?.ShapeProperties;
        var masterSpPr = masterShape?.ShapeProperties;

        var custGeom = spPr?.GetFirstChild<A.CustomGeometry>()
                       ?? layoutSpPr?.GetFirstChild<A.CustomGeometry>()
                       ?? masterSpPr?.GetFirstChild<A.CustomGeometry>();

        var prstGeom = spPr?.GetFirstChild<A.PresetGeometry>()
                       ?? layoutSpPr?.GetFirstChild<A.PresetGeometry>()
                       ?? masterSpPr?.GetFirstChild<A.PresetGeometry>();

        var shapeStyle = shape.ShapeStyle ?? layoutShape?.ShapeStyle ?? masterShape?.ShapeStyle;

        using var shapePath = BuildShapePath(rect, custGeom, prstGeom);

        // 2. Draw Shape Fill (Slide spPr -> Layout spPr -> Master spPr -> ShapeStyle fillRef)
        bool hasDrawnFill = false;
        SKColor? approxFillColor = null;

        if (spPr?.GetFirstChild<A.NoFill>() != null)
        {
            hasDrawnFill = true; // Explicitly no fill
        }
        else if (spPr != null && HasAnyFillChild(spPr))
        {
            hasDrawnFill = TryDrawFillFromContainer(canvas, spPr, ownerPart, ctx.ThemeColors, rect, null, shapePath, out approxFillColor);
        }
        else if (layoutSpPr?.GetFirstChild<A.NoFill>() != null)
        {
            hasDrawnFill = true;
        }
        else if (layoutSpPr != null && HasAnyFillChild(layoutSpPr))
        {
            hasDrawnFill = TryDrawFillFromContainer(canvas, layoutSpPr, ctx.LayoutPart ?? ownerPart, ctx.ThemeColors, rect, null, shapePath, out approxFillColor);
        }
        else if (masterSpPr != null && HasAnyFillChild(masterSpPr))
        {
            hasDrawnFill = TryDrawFillFromContainer(canvas, masterSpPr, ctx.MasterPart ?? ownerPart, ctx.ThemeColors, rect, null, shapePath, out approxFillColor);
        }

        if (!hasDrawnFill && shapeStyle?.FillReference != null)
        {
            uint fillIdx = shapeStyle.FillReference.Index?.Value ?? 0U;
            if (fillIdx > 0)
            {
                var refColor = ResolveColorFromContainer(shapeStyle.FillReference, ctx.ThemeColors, null);
                if (refColor.HasValue)
                {
                    approxFillColor = refColor.Value;
                    using var fillPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = refColor.Value, IsAntialias = true };
                    canvas.DrawPath(shapePath, fillPaint);
                }
            }
        }

        // 3. Draw Shape Outline (Slide spPr -> Layout spPr -> Master spPr -> ShapeStyle lnRef)
        var outline = spPr?.GetFirstChild<A.Outline>()
                      ?? layoutSpPr?.GetFirstChild<A.Outline>()
                      ?? masterSpPr?.GetFirstChild<A.Outline>();

        SKColor? strokeColor = null;
        float strokeWidth = 1f;

        if (outline != null)
        {
            if (outline.GetFirstChild<A.NoFill>() == null)
            {
                strokeColor = ResolveColorFromContainer(outline.GetFirstChild<A.SolidFill>(), ctx.ThemeColors, null);
                if (!strokeColor.HasValue && shapeStyle?.LineReference != null && (shapeStyle.LineReference.Index?.Value ?? 0U) > 0)
                {
                    strokeColor = ResolveColorFromContainer(shapeStyle.LineReference, ctx.ThemeColors, null);
                }
                if (outline.Width?.Value != null)
                {
                    strokeWidth = Math.Max(0.75f, (float)(outline.Width.Value / EmusPerPoint));
                }
            }
        }
        else if (shapeStyle?.LineReference != null && (shapeStyle.LineReference.Index?.Value ?? 0U) > 0)
        {
            strokeColor = ResolveColorFromContainer(shapeStyle.LineReference, ctx.ThemeColors, null);
        }

        if (strokeColor.HasValue)
        {
            using var strokePaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = strokeColor.Value,
                StrokeWidth = strokeWidth,
                IsAntialias = true
            };
            canvas.DrawPath(shapePath, strokePaint);
        }

        // Undo flip for text if shape was flipped (PowerPoint does not mirror text inside a flipped shape)
        if (flipH || flipV)
        {
            canvas.Translate(rect.MidX, rect.MidY);
            canvas.Scale(flipH ? -1f : 1f, flipV ? -1f : 1f);
            canvas.Translate(-rect.MidX, -rect.MidY);
        }

        // 4. Render TextBody with Full Style Inheritance
        var txBody = shape.TextBody;
        if (txBody != null)
        {
            RenderTextBody(canvas, ctx, shape, layoutShape, masterShape, phType, shapeStyle, rect, approxFillColor);
        }

        canvas.Restore();
    }

    private void RenderConnectionShape(
        SKCanvas canvas,
        SlideRenderContext ctx,
        P.ConnectionShape conn,
        float offsetX,
        float offsetY,
        float scaleX,
        float scaleY)
    {
        var xfrm = conn.ShapeProperties?.Transform2D;
        var rect = GetTransformRect(xfrm, offsetX, offsetY, scaleX, scaleY);
        if (rect.Width < 0.1f && rect.Height < 0.1f)
            return;

        var outline = conn.ShapeProperties?.GetFirstChild<A.Outline>();
        if (outline?.GetFirstChild<A.NoFill>() != null)
            return;

        SKColor? strokeColor = ResolveColorFromContainer(outline?.GetFirstChild<A.SolidFill>(), ctx.ThemeColors, null);
        if (!strokeColor.HasValue && conn.ShapeStyle?.LineReference != null)
        {
            strokeColor = ResolveColorFromContainer(conn.ShapeStyle.LineReference, ctx.ThemeColors, null);
        }
        strokeColor ??= ctx.ThemeColors.GetValueOrDefault("tx1", new SKColor(100, 116, 139));

        float strokeW = outline?.Width?.Value != null
            ? Math.Max(0.75f, (float)(outline.Width.Value / EmusPerPoint))
            : 1.25f;

        bool flipH = xfrm?.HorizontalFlip?.Value == true;
        bool flipV = xfrm?.VerticalFlip?.Value == true;

        float x1 = flipH ? rect.Right : rect.Left;
        float y1 = flipV ? rect.Bottom : rect.Top;
        float x2 = flipH ? rect.Left : rect.Right;
        float y2 = flipV ? rect.Top : rect.Bottom;

        using var paint = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            Color = strokeColor.Value,
            StrokeWidth = strokeW,
            StrokeCap = SKStrokeCap.Round,
            IsAntialias = true
        };
        canvas.DrawLine(x1, y1, x2, y2, paint);
    }

    // =========================================================================
    // 7. Rich TextBody Layout, Paragraph Level Inheritance & Vertical Alignment
    // =========================================================================

    private sealed class StyledRunWord
    {
        public required string Word { get; init; }
        public required float Width { get; init; }
        public required float SpaceWidth { get; init; }
        public required float FontSize { get; init; }
        public required SKTypeface Typeface { get; init; }
        public required SKColor Color { get; init; }
        public bool IsBold { get; init; }
        public bool IsItalic { get; init; }
        public bool IsUnderline { get; init; }
    }

    private sealed class LayoutVisualLine
    {
        public List<StyledRunWord> Words { get; } = new();
        public float MaxFontSize { get; set; }
        public float LineHeight { get; set; }
        public float IndentLeft { get; set; }
        public bool IsRtl { get; set; }
        public A.TextAlignmentTypeValues Alignment { get; set; } = A.TextAlignmentTypeValues.Left;
    }

    private void RenderTextBody(
        SKCanvas canvas,
        SlideRenderContext ctx,
        P.Shape shape,
        P.Shape? layoutShape,
        P.Shape? masterShape,
        P.PlaceholderValues? phType,
        P.ShapeStyle? shapeStyle,
        SKRect rect,
        SKColor? shapeFillColor)
    {
        var txBody = shape.TextBody!;
        var paragraphs = txBody.Elements<A.Paragraph>().ToList();
        if (paragraphs.Count == 0)
            return;

        // Check if all paragraphs are completely empty and have no fields
        if (paragraphs.All(p => string.IsNullOrWhiteSpace(ExtractDrawingParagraphText(p))))
            return;

        var bodyPr = txBody.BodyProperties;
        var layoutBodyPr = layoutShape?.TextBody?.BodyProperties;
        var masterBodyPr = masterShape?.TextBody?.BodyProperties;

        // Insets (default PowerPoint text box insets: L/R = 91440 EMU = 7.2pt, T/B = 45720 EMU = 3.6pt)
        float leftInset = GetEffectiveInset(bodyPr?.LeftInset, layoutBodyPr?.LeftInset, masterBodyPr?.LeftInset, 7.2f);
        float rightInset = GetEffectiveInset(bodyPr?.RightInset, layoutBodyPr?.RightInset, masterBodyPr?.RightInset, 7.2f);
        float topInset = GetEffectiveInset(bodyPr?.TopInset, layoutBodyPr?.TopInset, masterBodyPr?.TopInset, 3.6f);
        float bottomInset = GetEffectiveInset(bodyPr?.BottomInset, layoutBodyPr?.BottomInset, masterBodyPr?.BottomInset, 3.6f);

        var anchor = bodyPr?.Anchor?.Value
                     ?? layoutBodyPr?.Anchor?.Value
                     ?? masterBodyPr?.Anchor?.Value
                     ?? (phType == P.PlaceholderValues.CenteredTitle ? A.TextAnchoringTypeValues.Center : A.TextAnchoringTypeValues.Top);

        bool wordWrap = (bodyPr?.Wrap?.Value ?? layoutBodyPr?.Wrap?.Value ?? masterBodyPr?.Wrap?.Value) != A.TextWrappingValues.None;

        // NormAutofit fontScale (in 1/1000ths of a percent, e.g. 80000 = 80%)
        var normAutofit = bodyPr?.GetFirstChild<A.NormalAutoFit>()
                          ?? layoutBodyPr?.GetFirstChild<A.NormalAutoFit>()
                          ?? masterBodyPr?.GetFirstChild<A.NormalAutoFit>();

        float fontScale = normAutofit?.FontScale?.Value != null
            ? Math.Clamp(normAutofit.FontScale.Value / 100000f, 0.35f, 1.0f)
            : 1.0f;

        float lineSpaceReduction = normAutofit?.LineSpaceReduction?.Value != null
            ? Math.Clamp(normAutofit.LineSpaceReduction.Value / 100000f, 0f, 0.45f)
            : 0f;

        float availW = Math.Max(14f, rect.Width - leftInset - rightInset);
        float availH = Math.Max(12f, rect.Height - topInset - bottomInset);

        // Resolve shapeStyle fontRef color if present
        SKColor? styleFontColor = shapeStyle?.FontReference != null
            ? ResolveColorFromContainer(shapeStyle.FontReference, ctx.ThemeColors, null)
            : null;

        // Build visual lines with two-pass auto-fit scaling if text overflows the box vertically
        var visualLines = LayoutParagraphsToVisualLines(
            ctx,
            paragraphs,
            shape,
            layoutShape,
            masterShape,
            phType,
            styleFontColor,
            shapeFillColor,
            availW,
            wordWrap,
            fontScale,
            lineSpaceReduction,
            out float totalTextHeight);

        if (visualLines.Count == 0)
            return;

        // If text height exceeds the box height by more than 12%, scale down proportionally so nothing clips
        if (totalTextHeight > availH * 1.08f && availH > 20f)
        {
            float fitFactor = Math.Clamp((float)Math.Sqrt(availH / totalTextHeight), 0.52f, 0.92f);
            visualLines = LayoutParagraphsToVisualLines(
                ctx,
                paragraphs,
                shape,
                layoutShape,
                masterShape,
                phType,
                styleFontColor,
                shapeFillColor,
                availW,
                wordWrap,
                fontScale * fitFactor,
                lineSpaceReduction,
                out totalTextHeight);
        }

        // Compute starting Y based on vertical anchor (Top / Center / Bottom)
        float startY = rect.Top + topInset;
        if (anchor == A.TextAnchoringTypeValues.Center)
        {
            startY = rect.Top + topInset + Math.Max(0f, (availH - totalTextHeight) / 2f);
        }
        else if (anchor == A.TextAnchoringTypeValues.Bottom)
        {
            startY = rect.Bottom - bottomInset - totalTextHeight;
        }

        float curY = startY;
        float textLeft = rect.Left + leftInset;

        foreach (var vLine in visualLines)
        {
            if (vLine.Words.Count == 0)
            {
                curY += vLine.LineHeight;
                continue;
            }

            float baselineY = curY + (vLine.MaxFontSize * 0.88f);
            DrawStyledVisualLine(canvas, vLine, textLeft + vLine.IndentLeft, Math.Max(10f, availW - vLine.IndentLeft), baselineY);
            curY += vLine.LineHeight;
        }
    }

    private List<LayoutVisualLine> LayoutParagraphsToVisualLines(
        SlideRenderContext ctx,
        List<A.Paragraph> paragraphs,
        P.Shape shape,
        P.Shape? layoutShape,
        P.Shape? masterShape,
        P.PlaceholderValues? phType,
        SKColor? styleFontColor,
        SKColor? shapeFillColor,
        float availW,
        bool wordWrap,
        float fontScale,
        float lineSpaceReduction,
        out float totalHeight)
    {
        var allLines = new List<LayoutVisualLine>();
        float accHeight = 0f;
        int autoNumCounter = 1;

        for (int pIdx = 0; pIdx < paragraphs.Count; pIdx++)
        {
            var para = paragraphs[pIdx];
            var pPr = para.ParagraphProperties;
            int level = Math.Clamp(pPr?.Level?.Value ?? 0, 0, 8);

            // Gather inherited paragraph properties from Shape -> Layout -> Master -> Master.TextStyles -> Pres.DefaultTextStyle
            var inheritedPPrs = GetInheritedParagraphPropertyChain(ctx, shape, layoutShape, masterShape, phType, level, pPr);

            // Paragraph alignment
            var align = ResolveInheritedAlignment(inheritedPPrs, phType);
            bool paraRtl = ResolveInheritedRtl(inheritedPPrs) || OsFontService.ContainsArabic(ExtractDrawingParagraphText(para));
            if (paraRtl && align == null)
            {
                align = A.TextAlignmentTypeValues.Right;
            }
            var finalAlign = align ?? A.TextAlignmentTypeValues.Left;

            // Indentation / Margin
            float leftMarginPt = ResolveInheritedMarginLeft(inheritedPPrs, level);

            // SpaceBefore & SpaceAfter
            float spaceBeforePt = ResolveInheritedSpacingPoints(inheritedPPrs, isBefore: true, defaultFontSize: 16f);
            float spaceAfterPt = ResolveInheritedSpacingPoints(inheritedPPrs, isBefore: false, defaultFontSize: 16f);
            float lineSpacingMult = ResolveInheritedLineSpacingMultiplier(inheritedPPrs) * (1f - lineSpaceReduction);

            if (pIdx > 0 && spaceBeforePt > 0.5f)
            {
                accHeight += spaceBeforePt * fontScale;
                var spacer = new LayoutVisualLine { LineHeight = spaceBeforePt * fontScale };
                allLines.Add(spacer);
            }

            // Collect styled words from runs in this paragraph
            var paraWords = new List<(StyledRunWord Word, bool IsLineBreak)>();

            // Check for Bullet / Numbering on non-empty paragraphs
            string plainParaText = ExtractDrawingParagraphText(para);
            if (!string.IsNullOrWhiteSpace(plainParaText))
            {
                string? bulletPrefix = ResolveInheritedBulletPrefix(inheritedPPrs, autoNumCounter);
                if (!string.IsNullOrEmpty(bulletPrefix))
                {
                    autoNumCounter++;
                    var bulletStyle = ResolveRunStyling(ctx, null, inheritedPPrs, phType, styleFontColor, shapeFillColor, fontScale);
                    var bTf = _fontService.ResolveOsTypeface(bulletStyle.FontFamily, bulletStyle.IsBold, false, requiresArabicGlyphs: OsFontService.ContainsArabic(bulletPrefix));
                    using var bFont = new SKFont(bTf, bulletStyle.FontSize);
                    using var bShaper = new SKShaper(bTf);
                    float bW = MeasureWordWithEmojis(bulletPrefix, bShaper, bFont, bulletStyle.FontSize);
                    paraWords.Add((new StyledRunWord
                    {
                        Word = bulletPrefix,
                        Width = bW,
                        SpaceWidth = bFont.MeasureText(" ") * 1.3f,
                        FontSize = bulletStyle.FontSize,
                        Typeface = bTf,
                        Color = bulletStyle.Color,
                        IsBold = bulletStyle.IsBold
                    }, false));
                }
            }
            else
            {
                autoNumCounter = 1;
            }

            foreach (var child in para.ChildElements)
            {
                if (child is A.Break)
                {
                    paraWords.Add((null!, true));
                }
                else if (child is A.Run run)
                {
                    string runText = run.Text?.Text ?? string.Empty;
                    if (string.IsNullOrEmpty(runText)) continue;
                    AppendRunWords(ctx, runText, run.RunProperties, inheritedPPrs, phType, styleFontColor, shapeFillColor, fontScale, paraWords);
                }
                else if (child is A.Field fld)
                {
                    string fldText = fld.Text?.Text ?? string.Empty;
                    if (fld.Type?.Value != null && fld.Type.Value.Contains("slidenum", StringComparison.OrdinalIgnoreCase))
                    {
                        fldText = ctx.SlideNumber.ToString(CultureInfo.InvariantCulture);
                    }
                    if (string.IsNullOrEmpty(fldText)) continue;
                    AppendRunWords(ctx, fldText, fld.RunProperties, inheritedPPrs, phType, styleFontColor, shapeFillColor, fontScale, paraWords);
                }
            }

            if (paraWords.Count == 0)
            {
                var emptyStyle = ResolveRunStyling(ctx, null, inheritedPPrs, phType, styleFontColor, shapeFillColor, fontScale);
                float emptyH = emptyStyle.FontSize * 0.75f;
                allLines.Add(new LayoutVisualLine { LineHeight = emptyH });
                accHeight += emptyH;
                continue;
            }

            // Wrap paraWords into visual lines within (availW - leftMarginPt)
            float lineMaxW = wordWrap ? Math.Max(20f, availW - leftMarginPt) : 100000f;
            var currentLine = new LayoutVisualLine
            {
                IndentLeft = paraRtl ? 0f : leftMarginPt,
                IsRtl = paraRtl,
                Alignment = finalAlign
            };
            float currentLineW = 0f;

            void FlushCurrentLine()
            {
                if (currentLine.Words.Count == 0) return;
                float maxFs = currentLine.Words.Max(w => w.FontSize);
                currentLine.MaxFontSize = maxFs;
                currentLine.LineHeight = maxFs * lineSpacingMult;
                allLines.Add(currentLine);
                accHeight += currentLine.LineHeight;

                currentLine = new LayoutVisualLine
                {
                    IndentLeft = paraRtl ? 0f : leftMarginPt,
                    IsRtl = paraRtl,
                    Alignment = finalAlign
                };
                currentLineW = 0f;
            }

            foreach (var (wordObj, isBreak) in paraWords)
            {
                if (isBreak)
                {
                    FlushCurrentLine();
                    continue;
                }

                float spW = currentLine.Words.Count == 0 ? 0f : currentLine.Words[^1].SpaceWidth;
                float projected = currentLineW + spW + wordObj.Width;

                if (projected > lineMaxW && currentLine.Words.Count > 0)
                {
                    FlushCurrentLine();
                    currentLine.Words.Add(wordObj);
                    currentLineW = wordObj.Width;
                }
                else
                {
                    currentLine.Words.Add(wordObj);
                    currentLineW = projected;
                }
            }

            FlushCurrentLine();

            if (spaceAfterPt > 0.5f)
            {
                accHeight += spaceAfterPt * fontScale;
                allLines.Add(new LayoutVisualLine { LineHeight = spaceAfterPt * fontScale });
            }
        }

        totalHeight = accHeight;
        return allLines;
    }

    private void AppendRunWords(
        SlideRenderContext ctx,
        string runText,
        A.RunProperties? runPr,
        List<A.TextParagraphPropertiesType> inheritedPPrs,
        P.PlaceholderValues? phType,
        SKColor? styleFontColor,
        SKColor? shapeFillColor,
        float fontScale,
        List<(StyledRunWord Word, bool IsLineBreak)> targetList)
    {
        var style = ResolveRunStyling(ctx, runPr, inheritedPPrs, phType, styleFontColor, shapeFillColor, fontScale);
        bool hasAr = OsFontService.ContainsArabic(runText);
        var tf = _fontService.ResolveOsTypeface(style.FontFamily, style.IsBold, style.IsItalic, requiresArabicGlyphs: hasAr);

        using var font = new SKFont(tf, style.FontSize)
        {
            Subpixel = true,
            Edging = SKFontEdging.SubpixelAntialias,
            Embolden = style.IsBold && !tf.IsBold,
            SkewX = (style.IsItalic && !tf.IsItalic) ? -0.18f : 0f
        };
        using var shaper = new SKShaper(tf);
        float spaceW = font.MeasureText(" ");

        var lines = runText.Replace("\r\n", "\n").Split('\n');
        for (int l = 0; l < lines.Length; l++)
        {
            if (l > 0)
            {
                targetList.Add((null!, true));
            }

            var words = lines[l].Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (var w in words)
            {
                float wWidth = MeasureWordWithEmojis(w, shaper, font, style.FontSize);
                targetList.Add((new StyledRunWord
                {
                    Word = w,
                    Width = wWidth,
                    SpaceWidth = spaceW,
                    FontSize = style.FontSize,
                    Typeface = tf,
                    Color = style.Color,
                    IsBold = style.IsBold,
                    IsItalic = style.IsItalic,
                    IsUnderline = style.IsUnderline
                }, false));
            }
        }
    }

    private static void DrawStyledVisualLine(
        SKCanvas canvas,
        LayoutVisualLine line,
        float startX,
        float maxWidth,
        float baselineY)
    {
        if (line.Words.Count == 0) return;

        float totalLineWidth = 0f;
        for (int i = 0; i < line.Words.Count; i++)
        {
            totalLineWidth += line.Words[i].Width;
            if (i < line.Words.Count - 1)
                totalLineWidth += line.Words[i].SpaceWidth;
        }

        if (line.IsRtl)
        {
            float cursorRight = line.Alignment == A.TextAlignmentTypeValues.Center
                ? startX + (maxWidth + totalLineWidth) / 2f
                : line.Alignment == A.TextAlignmentTypeValues.Left
                    ? startX + Math.Min(maxWidth, totalLineWidth)
                    : startX + maxWidth;

            foreach (var sw in line.Words)
            {
                float wordLeft = cursorRight - sw.Width;
                DrawSingleStyledWord(canvas, sw, isRtl: true, wordLeft, baselineY);
                cursorRight = wordLeft - sw.SpaceWidth;
            }
        }
        else
        {
            float cursorLeft = line.Alignment == A.TextAlignmentTypeValues.Center
                ? startX + Math.Max(0f, (maxWidth - totalLineWidth) / 2f)
                : line.Alignment == A.TextAlignmentTypeValues.Right
                    ? startX + Math.Max(0f, maxWidth - totalLineWidth)
                    : startX;

            foreach (var sw in line.Words)
            {
                DrawSingleStyledWord(canvas, sw, isRtl: false, cursorLeft, baselineY);
                cursorLeft += sw.Width + sw.SpaceWidth;
            }
        }
    }

    private static void DrawSingleStyledWord(
        SKCanvas canvas,
        StyledRunWord sw,
        bool isRtl,
        float wordLeft,
        float baselineY)
    {
        using var font = new SKFont(sw.Typeface, sw.FontSize)
        {
            Subpixel = true,
            Edging = SKFontEdging.SubpixelAntialias,
            Embolden = sw.IsBold && !sw.Typeface.IsBold,
            SkewX = (sw.IsItalic && !sw.Typeface.IsItalic) ? -0.18f : 0f
        };
        using var shaper = new SKShaper(sw.Typeface);
        using var paint = new SKPaint
        {
            IsAntialias = true,
            Color = sw.Color
        };

        DrawWordTokens(canvas, sw.Word, shaper, font, paint, sw.FontSize, isRtl, wordLeft, sw.Width, baselineY);

        if (sw.IsUnderline)
        {
            using var uPaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = sw.Color,
                StrokeWidth = Math.Max(0.8f, sw.FontSize * 0.055f),
                IsAntialias = true
            };
            float uY = baselineY + (sw.FontSize * 0.11f);
            canvas.DrawLine(wordLeft, uY, wordLeft + sw.Width, uY, uPaint);
        }
    }

    // =========================================================================
    // 8. Paragraph & Run Inheritance Chain Resolvers
    // =========================================================================

    private static List<A.TextParagraphPropertiesType> GetInheritedParagraphPropertyChain(
        SlideRenderContext ctx,
        P.Shape shape,
        P.Shape? layoutShape,
        P.Shape? masterShape,
        P.PlaceholderValues? phType,
        int level,
        A.ParagraphProperties? directPPr)
    {
        var chain = new List<A.TextParagraphPropertiesType>(6);
        if (directPPr != null) chain.Add(directPPr);

        AddLevelPPrFromListStyle(chain, shape.TextBody?.ListStyle, level);
        AddLevelPPrFromListStyle(chain, layoutShape?.TextBody?.ListStyle, level);
        AddLevelPPrFromListStyle(chain, masterShape?.TextBody?.ListStyle, level);

        var txStyles = ctx.MasterPart?.SlideMaster?.TextStyles;
        if (txStyles != null)
        {
            OpenXmlCompositeElement? masterStyle = (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle)
                ? txStyles.TitleStyle
                : (phType == P.PlaceholderValues.Body || phType == P.PlaceholderValues.SubTitle || phType == P.PlaceholderValues.Object)
                    ? txStyles.BodyStyle
                    : txStyles.OtherStyle;

            AddLevelPPrFromListStyle(chain, masterStyle, level);
        }

        AddLevelPPrFromListStyle(chain, ctx.PresentationPart.Presentation?.DefaultTextStyle, level);
        return chain;
    }

    private static void AddLevelPPrFromListStyle(List<A.TextParagraphPropertiesType> chain, OpenXmlCompositeElement? listStyle, int level)
    {
        if (listStyle == null) return;

        A.TextParagraphPropertiesType? lvlPPr = level switch
        {
            0 => listStyle.GetFirstChild<A.Level1ParagraphProperties>(),
            1 => listStyle.GetFirstChild<A.Level2ParagraphProperties>(),
            2 => listStyle.GetFirstChild<A.Level3ParagraphProperties>(),
            3 => listStyle.GetFirstChild<A.Level4ParagraphProperties>(),
            4 => listStyle.GetFirstChild<A.Level5ParagraphProperties>(),
            5 => listStyle.GetFirstChild<A.Level6ParagraphProperties>(),
            6 => listStyle.GetFirstChild<A.Level7ParagraphProperties>(),
            7 => listStyle.GetFirstChild<A.Level8ParagraphProperties>(),
            8 => listStyle.GetFirstChild<A.Level9ParagraphProperties>(),
            _ => null
        };

        lvlPPr ??= (A.TextParagraphPropertiesType?)listStyle.GetFirstChild<A.Level1ParagraphProperties>()
                  ?? listStyle.GetFirstChild<A.DefaultParagraphProperties>();

        if (lvlPPr != null)
            chain.Add(lvlPPr);
    }

    private static A.TextAlignmentTypeValues? ResolveInheritedAlignment(List<A.TextParagraphPropertiesType> chain, P.PlaceholderValues? phType)
    {
        foreach (var pPr in chain)
        {
            if (pPr.Alignment?.Value != null)
                return pPr.Alignment.Value;
        }
        if (phType == P.PlaceholderValues.CenteredTitle || phType == P.PlaceholderValues.SubTitle)
            return A.TextAlignmentTypeValues.Center;
        return null;
    }

    private static bool ResolveInheritedRtl(List<A.TextParagraphPropertiesType> chain)
    {
        foreach (var pPr in chain)
        {
            if (pPr.RightToLeft?.Value != null)
                return pPr.RightToLeft.Value;
        }
        return false;
    }

    private static float ResolveInheritedMarginLeft(List<A.TextParagraphPropertiesType> chain, int level)
    {
        foreach (var pPr in chain)
        {
            if (pPr.LeftMargin?.Value != null)
            {
                float m = (float)(pPr.LeftMargin.Value / EmusPerPoint);
                float indent = pPr.Indent?.Value != null ? (float)(pPr.Indent.Value / EmusPerPoint) : 0f;
                return Math.Max(0f, m + indent);
            }
        }
        return level * 18f;
    }

    private static float ResolveInheritedSpacingPoints(List<A.TextParagraphPropertiesType> chain, bool isBefore, float defaultFontSize)
    {
        foreach (var pPr in chain)
        {
            OpenXmlCompositeElement? spc = isBefore
                ? pPr.GetFirstChild<A.SpaceBefore>()
                : pPr.GetFirstChild<A.SpaceAfter>();

            if (spc != null)
            {
                var pts = spc.GetFirstChild<A.SpacingPoints>();
                if (pts?.Val?.Value != null)
                    return Math.Clamp(pts.Val.Value / 100f, 0f, 48f);

                var pct = spc.GetFirstChild<A.SpacingPercent>();
                if (pct?.Val?.Value != null)
                    return Math.Clamp((pct.Val.Value / 100000f) * defaultFontSize, 0f, 48f);
            }
        }
        return isBefore ? 0f : 3f;
    }

    private static float ResolveInheritedLineSpacingMultiplier(List<A.TextParagraphPropertiesType> chain)
    {
        foreach (var pPr in chain)
        {
            var lnSpc = pPr.GetFirstChild<A.LineSpacing>();
            var pct = lnSpc?.GetFirstChild<A.SpacingPercent>();
            if (pct?.Val?.Value != null)
            {
                return Math.Clamp(pct.Val.Value / 100000f, 0.9f, 2.2f) * 1.22f;
            }
        }
        return 1.28f;
    }

    private static string? ResolveInheritedBulletPrefix(List<A.TextParagraphPropertiesType> chain, int autoNumIndex)
    {
        foreach (var pPr in chain)
        {
            if (pPr.GetFirstChild<A.NoBullet>() != null)
                return null;

            var charBu = pPr.GetFirstChild<A.CharacterBullet>();
            if (charBu?.Char?.Value != null)
            {
                string ch = charBu.Char.Value;
                // Map common Wingdings/Private Use bullet codepoints to standard clean Unicode bullets
                if (ch.Length == 1 && (ch[0] >= '\uF000' || ch[0] == '§' || ch[0] == 'q' || ch[0] == 'n' || ch[0] == 'l'))
                    return "•";
                return ch;
            }

            if (pPr.GetFirstChild<A.AutoNumberedBullet>() != null)
                return $"{autoNumIndex}.";
        }
        return null;
    }

    private (string FontFamily, float FontSize, bool IsBold, bool IsItalic, bool IsUnderline, SKColor Color) ResolveRunStyling(
        SlideRenderContext ctx,
        A.RunProperties? runPr,
        List<A.TextParagraphPropertiesType> inheritedPPrs,
        P.PlaceholderValues? phType,
        SKColor? styleFontColor,
        SKColor? shapeFillColor,
        float fontScale)
    {
        // Collect DefaultRunProperties from the paragraph inheritance chain
        var defRPrs = new List<A.TextCharacterPropertiesType>(inheritedPPrs.Count + 1);
        if (runPr != null) defRPrs.Add(runPr);
        foreach (var pPr in inheritedPPrs)
        {
            var defRPr = pPr.GetFirstChild<A.DefaultRunProperties>();
            if (defRPr != null) defRPrs.Add(defRPr);
        }

        // 1. Font Size
        float? fontSize = null;
        foreach (var rPr in defRPrs)
        {
            if (rPr.FontSize?.Value != null)
            {
                fontSize = Math.Clamp(rPr.FontSize.Value / 100f, 6f, 96f);
                break;
            }
        }
        fontSize ??= (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle)
            ? 34f
            : phType == P.PlaceholderValues.SubTitle
                ? 20f
                : (phType == P.PlaceholderValues.Body || phType == P.PlaceholderValues.Object)
                    ? 18f
                    : 15f;

        float scaledFontSize = Math.Max(6.5f, fontSize.Value * fontScale);

        // 2. Bold / Italic / Underline
        bool? isBold = null;
        bool? isItalic = null;
        bool isUnderline = runPr?.Underline?.Value != null && runPr.Underline.Value != A.TextUnderlineValues.None;

        foreach (var rPr in defRPrs)
        {
            if (!isBold.HasValue && rPr.Bold?.Value != null)
                isBold = rPr.Bold.Value;
            if (!isItalic.HasValue && rPr.Italic?.Value != null)
                isItalic = rPr.Italic.Value;
        }

        // 3. Font Family (Resolve +mj-lt / +mn-lt / +mj-cs / +mn-cs theme font tokens)
        string? rawTypeface = null;
        foreach (var rPr in defRPrs)
        {
            rawTypeface = rPr.GetFirstChild<A.LatinFont>()?.Typeface?.Value
                          ?? rPr.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(rawTypeface))
                break;
        }

        string resolvedFont = ResolveThemeFontAlias(
            rawTypeface,
            ctx,
            isHeading: phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle);

        // 4. Text Color
        SKColor? textColor = null;
        foreach (var rPr in defRPrs)
        {
            var solid = rPr.GetFirstChild<A.SolidFill>();
            if (solid != null)
            {
                textColor = ResolveColorFromContainer(solid, ctx.ThemeColors, null);
                if (textColor.HasValue) break;
            }
        }

        textColor ??= styleFontColor;
        if (!textColor.HasValue)
        {
            if (shapeFillColor.HasValue)
            {
                textColor = IsDarkColor(shapeFillColor.Value) ? SKColors.White : new SKColor(15, 23, 42);
            }
            else if (ctx.ThemeColors.TryGetValue("tx1", out var tx1))
            {
                textColor = tx1;
            }
            else
            {
                textColor = new SKColor(15, 23, 42);
            }
        }

        return (resolvedFont, scaledFontSize, isBold ?? false, isItalic ?? false, isUnderline, textColor.Value);
    }

    private static string ResolveThemeFontAlias(string? typefaceToken, SlideRenderContext ctx, bool isHeading)
    {
        if (string.IsNullOrWhiteSpace(typefaceToken))
            return isHeading ? ctx.MajorLatinFont : ctx.MinorLatinFont;

        string t = typefaceToken.Trim();
        if (t.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase))
            return t.EndsWith("cs", StringComparison.OrdinalIgnoreCase) ? ctx.MajorComplexFont : ctx.MajorLatinFont;
        if (t.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase))
            return t.EndsWith("cs", StringComparison.OrdinalIgnoreCase) ? ctx.MinorComplexFont : ctx.MinorLatinFont;

        return t;
    }

    // =========================================================================
    // 9. Pictures (with Cropping & Rotation) & Tables
    // =========================================================================

    private static void RenderPicture(
        SKCanvas canvas,
        SlideRenderContext ctx,
        OpenXmlPart ownerPart,
        P.Picture pic,
        float offsetX,
        float offsetY,
        float scaleX,
        float scaleY)
    {
        if (pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Hidden?.Value == true)
            return;

        var xfrm = pic.ShapeProperties?.Transform2D;
        var rect = GetTransformRect(xfrm, offsetX, offsetY, scaleX, scaleY);
        if (rect.Width <= 1f || rect.Height <= 1f)
            return;

        string? embedId = pic.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId))
            return;

        try
        {
            if (ownerPart.GetPartById(embedId) is not ImagePart imgPart)
                return;

            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            using var skImg = SKImage.FromEncodedData(ms.ToArray());
            if (skImg == null || skImg.Width <= 0 || skImg.Height <= 0)
                return;

            float rotDeg = xfrm?.Rotation != null ? xfrm.Rotation.Value / 60000f : 0f;
            bool flipH = xfrm?.HorizontalFlip?.Value == true;
            bool flipV = xfrm?.VerticalFlip?.Value == true;

            canvas.Save();
            if (Math.Abs(rotDeg) > 0.05f || flipH || flipV)
            {
                canvas.Translate(rect.MidX, rect.MidY);
                if (Math.Abs(rotDeg) > 0.05f) canvas.RotateDegrees(rotDeg);
                if (flipH || flipV) canvas.Scale(flipH ? -1f : 1f, flipV ? -1f : 1f);
                canvas.Translate(-rect.MidX, -rect.MidY);
            }

            // Clip to direct or layout-inherited preset/custom shape if picture has an ellipse, hexagon, roundRect, or custom polygon
            var layoutShapes = ctx.LayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
            var masterShapes = ctx.MasterPart?.SlideMaster?.CommonSlideData?.ShapeTree?.Descendants<P.Shape>().ToList();
            var (custGeom, prstGeom, _) = ResolvePictureClipGeometry(pic, layoutShapes, masterShapes);
            if ((prstGeom?.Preset?.Value != null && prstGeom.Preset.Value != A.ShapeTypeValues.Rectangle) || custGeom != null)
            {
                using var clipPath = BuildShapePath(rect, custGeom, prstGeom);
                canvas.ClipPath(clipPath, SKClipOperation.Intersect, antialias: true);
            }

            // Check SourceRectangle crop (in 1/1000ths of a percent: 100000 = 100%)
            var srcRectEl = pic.BlipFill?.SourceRectangle;
            if (srcRectEl != null &&
                (srcRectEl.Left?.Value != null || srcRectEl.Top?.Value != null || srcRectEl.Right?.Value != null || srcRectEl.Bottom?.Value != null))
            {
                float cL = (srcRectEl.Left?.Value ?? 0) / 100000f;
                float cT = (srcRectEl.Top?.Value ?? 0) / 100000f;
                float cR = (srcRectEl.Right?.Value ?? 0) / 100000f;
                float cB = (srcRectEl.Bottom?.Value ?? 0) / 100000f;

                float sx1 = Math.Clamp(cL * skImg.Width, 0f, skImg.Width - 1f);
                float sy1 = Math.Clamp(cT * skImg.Height, 0f, skImg.Height - 1f);
                float sx2 = Math.Clamp((1f - cR) * skImg.Width, sx1 + 1f, skImg.Width);
                float sy2 = Math.Clamp((1f - cB) * skImg.Height, sy1 + 1f, skImg.Height);

                canvas.DrawImage(skImg, new SKRect(sx1, sy1, sx2, sy2), rect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
            }
            else
            {
                canvas.DrawImage(skImg, rect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
            }

            canvas.Restore();
        }
        catch
        {
            // Ignore unreadable embedded image part
        }
    }

    private void RenderGraphicFrame(
        SKCanvas canvas,
        SlideRenderContext ctx,
        P.GraphicFrame frame,
        float offsetX,
        float offsetY,
        float scaleX,
        float scaleY)
    {
        var rect = GetTransformRect(frame.Transform, offsetX, offsetY, scaleX, scaleY);
        if (rect.Width <= 1f || rect.Height <= 1f)
            return;

        var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();
        if (table == null)
            return;

        var rows = table.Elements<A.TableRow>().ToList();
        if (rows.Count == 0)
            return;

        var gridCols = table.TableGrid?.Elements<A.GridColumn>().ToList() ?? new List<A.GridColumn>();
        float[] colWidths;
        if (gridCols.Count > 0)
        {
            double totalGridEmu = gridCols.Sum(c => c.Width?.Value ?? 0);
            colWidths = gridCols.Select(c => totalGridEmu > 0
                ? (float)((c.Width?.Value ?? 0) / totalGridEmu) * rect.Width
                : rect.Width / gridCols.Count).ToArray();
        }
        else
        {
            int cols = rows[0].Elements<A.TableCell>().Count();
            colWidths = Enumerable.Repeat(rect.Width / Math.Max(1, cols), Math.Max(1, cols)).ToArray();
        }

        bool firstRowHeader = table.TableProperties?.FirstRow?.Value ?? true;
        bool bandRow = table.TableProperties?.BandRow?.Value ?? true;
        SKColor accentHeader = ctx.ThemeColors.GetValueOrDefault("accent1", new SKColor(37, 99, 235));

        double totalRowEmu = rows.Sum(r => r.Height?.Value ?? 370840);
        float curY = rect.Top;

        using var borderPaint = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            Color = new SKColor(203, 213, 225),
            StrokeWidth = 0.8f,
            IsAntialias = true
        };

        for (int rIdx = 0; rIdx < rows.Count; rIdx++)
        {
            var row = rows[rIdx];
            float rowH = totalRowEmu > 0
                ? (float)((row.Height?.Value ?? 370840) / totalRowEmu) * rect.Height
                : rect.Height / rows.Count;

            var cells = row.Elements<A.TableCell>().ToList();
            float curX = rect.Left;

            for (int cIdx = 0; cIdx < Math.Min(cells.Count, colWidths.Length); cIdx++)
            {
                float cellW = colWidths[cIdx];
                var cell = cells[cIdx];
                var cellRect = new SKRect(curX, curY, curX + cellW, curY + rowH);

                var explicitFill = ResolveColorFromContainer(cell.TableCellProperties?.GetFirstChild<A.SolidFill>(), ctx.ThemeColors, null);
                SKColor cellBg = explicitFill ?? (rIdx == 0 && firstRowHeader
                    ? accentHeader
                    : (bandRow && rIdx % 2 == 1 ? new SKColor(241, 245, 249) : SKColors.White));

                using var bgPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = cellBg, IsAntialias = true };
                canvas.DrawRect(cellRect, bgPaint);
                canvas.DrawRect(cellRect, borderPaint);

                string cellText = cell.TextBody != null
                    ? string.Join("\n", cell.TextBody.Elements<A.Paragraph>().Select(ExtractDrawingParagraphText)).Trim()
                    : string.Empty;

                if (!string.IsNullOrEmpty(cellText))
                {
                    bool isRtl = OsFontService.ContainsArabic(cellText);
                    var firstRun = cell.TextBody?.Descendants<A.Run>().FirstOrDefault();
                    SKColor? runClr = ResolveColorFromContainer(firstRun?.RunProperties?.GetFirstChild<A.SolidFill>(), ctx.ThemeColors, null);
                    SKColor txtColor = runClr ?? (IsDarkColor(cellBg) ? SKColors.White : new SKColor(15, 23, 42));

                    float fontSize = firstRun?.RunProperties?.FontSize?.Value != null
                        ? Math.Clamp(firstRun.RunProperties.FontSize.Value / 100f, 7f, 24f)
                        : 11f;
                    bool isBold = (rIdx == 0 && firstRowHeader) || (firstRun?.RunProperties?.Bold?.Value == true);

                    using var tf = _fontService.ResolveOsTypeface(ctx.MinorLatinFont, isBold, isItalic: false, requiresArabicGlyphs: isRtl);
                    using var font = new SKFont(tf, fontSize) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias, Embolden = isBold && !tf.IsBold };
                    using var shaper = new SKShaper(tf);
                    using var txtPaint = new SKPaint { IsAntialias = true, Color = txtColor };

                    var lines = WrapPlainTextLines(cellText, shaper, font, fontSize, Math.Max(10f, cellW - 10f));
                    float lineH = fontSize * 1.28f;
                    float totalTxtH = lines.Count * lineH;
                    float startLineY = curY + Math.Max(4f, (rowH - totalTxtH) / 2f) + (fontSize * 0.88f);

                    var cellAlign = cell.TextBody?.Elements<A.Paragraph>().FirstOrDefault()?.ParagraphProperties?.Alignment?.Value
                                    ?? (isRtl ? A.TextAlignmentTypeValues.Right : A.TextAlignmentTypeValues.Left);

                    foreach (var line in lines)
                    {
                        if (startLineY > curY + rowH + 2f) break;
                        DrawPlainVisualLine(canvas, line, shaper, font, txtPaint, fontSize, isRtl, curX + 5f, cellW - 10f, font.MeasureText(" "), startLineY, cellAlign);
                        startLineY += lineH;
                    }
                }

                curX += cellW;
            }

            curY += rowH;
        }
    }

    // =========================================================================
    // 10. Color Resolution, Modifiers (lumMod, lumOff, tint, shade, alpha) & Gradients
    // =========================================================================

    private static bool HasAnyFillChild(OpenXmlCompositeElement container)
    {
        return container.GetFirstChild<A.SolidFill>() != null
               || container.GetFirstChild<A.GradientFill>() != null
               || container.GetFirstChild<A.BlipFill>() != null;
    }

    private static bool TryDrawFillFromContainer(
        SKCanvas canvas,
        OpenXmlCompositeElement container,
        OpenXmlPart ownerPart,
        Dictionary<string, SKColor> themeColors,
        SKRect rect,
        SKColor? phColor)
    {
        using var rectPath = new SKPath();
        rectPath.AddRect(rect);
        return TryDrawFillFromContainer(canvas, container, ownerPart, themeColors, rect, phColor, rectPath, out _);
    }

    private static bool TryDrawFillFromContainer(
        SKCanvas canvas,
        OpenXmlCompositeElement container,
        OpenXmlPart ownerPart,
        Dictionary<string, SKColor> themeColors,
        SKRect rect,
        SKColor? phColor,
        SKPath targetPath,
        out SKColor? approxColor)
    {
        approxColor = null;
        foreach (var child in container.ChildElements)
        {
            if (child is A.SolidFill || child is A.GradientFill || child is A.BlipFill)
            {
                if (TryDrawFillElement(canvas, child, ownerPart, themeColors, rect, phColor, targetPath, out approxColor))
                    return true;
            }
        }
        return false;
    }

    private static bool TryDrawFillElement(
        SKCanvas canvas,
        OpenXmlElement fillEl,
        OpenXmlPart ownerPart,
        Dictionary<string, SKColor> themeColors,
        SKRect rect,
        SKColor? phColor)
    {
        using var rectPath = new SKPath();
        rectPath.AddRect(rect);
        return TryDrawFillElement(canvas, fillEl, ownerPart, themeColors, rect, phColor, rectPath, out _);
    }

    private static bool TryDrawFillElement(
        SKCanvas canvas,
        OpenXmlElement fillEl,
        OpenXmlPart ownerPart,
        Dictionary<string, SKColor> themeColors,
        SKRect rect,
        SKColor? phColor,
        SKPath targetPath,
        out SKColor? approxColor)
    {
        approxColor = null;

        if (fillEl is A.SolidFill solid)
        {
            var c = ResolveColorFromContainer(solid, themeColors, phColor);
            if (!c.HasValue) return false;
            approxColor = c.Value;
            using var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = c.Value, IsAntialias = true };
            canvas.DrawPath(targetPath, paint);
            return true;
        }

        if (fillEl is A.GradientFill grad)
        {
            var stops = grad.GradientStopList?.Elements<A.GradientStop>().ToList();
            if (stops == null || stops.Count == 0) return false;

            var skColors = new List<SKColor>();
            var skPositions = new List<float>();

            foreach (var gs in stops.OrderBy(s => s.Position?.Value ?? 0))
            {
                var c = ResolveColorFromContainer(gs, themeColors, phColor);
                if (c.HasValue)
                {
                    skColors.Add(c.Value);
                    skPositions.Add(Math.Clamp((gs.Position?.Value ?? 0) / 100000f, 0f, 1f));
                }
            }

            if (skColors.Count == 0) return false;
            if (skColors.Count == 1)
            {
                approxColor = skColors[0];
                using var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = skColors[0], IsAntialias = true };
                canvas.DrawPath(targetPath, paint);
                return true;
            }

            approxColor = skColors[skColors.Count / 2];

            var lin = grad.GetFirstChild<A.LinearGradientFill>();
            float angleDeg = lin?.Angle?.Value != null ? lin.Angle.Value / 60000f : 90f;
            float rad = angleDeg * (float)(Math.PI / 180.0);

            float cx = rect.MidX;
            float cy = rect.MidY;
            float dx = (float)Math.Cos(rad) * (rect.Width / 2f);
            float dy = (float)Math.Sin(rad) * (rect.Height / 2f);

            using var shader = SKShader.CreateLinearGradient(
                new SKPoint(cx - dx, cy - dy),
                new SKPoint(cx + dx, cy + dy),
                skColors.ToArray(),
                skPositions.ToArray(),
                SKShaderTileMode.Clamp);

            using var gPaint = new SKPaint { Style = SKPaintStyle.Fill, Shader = shader, IsAntialias = true };
            canvas.DrawPath(targetPath, gPaint);
            return true;
        }

        if (fillEl is A.BlipFill blipFill)
        {
            string? embedId = blipFill.Blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) return false;

            try
            {
                if (ownerPart.GetPartById(embedId) is ImagePart imgPart)
                {
                    using var stream = imgPart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    using var skImg = SKImage.FromEncodedData(ms.ToArray());
                    if (skImg != null)
                    {
                        canvas.Save();
                        canvas.ClipPath(targetPath, SKClipOperation.Intersect, antialias: true);
                        canvas.DrawImage(skImg, rect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
                        canvas.Restore();
                        return true;
                    }
                }
            }
            catch
            {
                // Ignore broken image fill
            }
        }

        return false;
    }

    private static SKColor? ResolveColorFromContainer(
        OpenXmlCompositeElement? container,
        Dictionary<string, SKColor> themeColors,
        SKColor? phColor)
    {
        if (container == null) return null;

        SKColor? baseColor = null;
        OpenXmlCompositeElement? colorElement = null;

        var srgb = container.GetFirstChild<A.RgbColorModelHex>();
        if (srgb?.Val?.Value != null && SKColor.TryParse("#" + srgb.Val.Value, out var parsedRgb))
        {
            baseColor = parsedRgb;
            colorElement = srgb;
        }

        if (!baseColor.HasValue)
        {
            var scheme = container.GetFirstChild<A.SchemeColor>();
            if (scheme?.Val?.Value != null)
            {
                colorElement = scheme;
                string key = ((IEnumValue)scheme.Val.Value).Value ?? string.Empty;
                if (scheme.Val.Value == A.SchemeColorValues.PhColor && phColor.HasValue)
                {
                    baseColor = phColor.Value;
                }
                else if (themeColors.TryGetValue(key, out var tc))
                {
                    baseColor = tc;
                }
            }
        }

        if (!baseColor.HasValue)
        {
            var sys = container.GetFirstChild<A.SystemColor>();
            if (sys?.LastColor?.Value != null && SKColor.TryParse("#" + sys.LastColor.Value, out var sysColor))
            {
                baseColor = sysColor;
                colorElement = sys;
            }
        }

        if (!baseColor.HasValue)
        {
            var prst = container.GetFirstChild<A.PresetColor>();
            if (prst?.Val?.Value != null)
            {
                string prstName = ((IEnumValue)prst.Val.Value).Value ?? string.Empty;
                if (SKColor.TryParse(prstName, out var prstColor))
                {
                    baseColor = prstColor;
                    colorElement = prst;
                }
            }
        }

        if (!baseColor.HasValue || colorElement == null)
            return baseColor;

        return ApplyColorTransforms(baseColor.Value, colorElement);
    }

    private static SKColor ApplyColorTransforms(SKColor color, OpenXmlCompositeElement colorElement)
    {
        color.ToHsl(out float h, out float s, out float l);
        float alpha = color.Alpha / 255f;
        bool hslModified = false;

        foreach (var mod in colorElement.ChildElements)
        {
            if (mod is A.LuminanceModulation lumMod && lumMod.Val?.Value != null)
            {
                l *= (lumMod.Val.Value / 100000f);
                hslModified = true;
            }
            else if (mod is A.LuminanceOffset lumOff && lumOff.Val?.Value != null)
            {
                l += (lumOff.Val.Value / 1000f); // OpenXML 100000 = 100% on 0..100 scale
                hslModified = true;
            }
            else if (mod is A.SaturationModulation satMod && satMod.Val?.Value != null)
            {
                s *= (satMod.Val.Value / 100000f);
                hslModified = true;
            }
            else if (mod is A.Alpha alphaMod && alphaMod.Val?.Value != null)
            {
                alpha = Math.Clamp(alphaMod.Val.Value / 100000f, 0f, 1f);
            }
        }

        SKColor current = hslModified
            ? SKColor.FromHsl(Math.Clamp(h, 0f, 360f), Math.Clamp(s, 0f, 100f), Math.Clamp(l, 0f, 100f))
            : color;

        foreach (var mod in colorElement.ChildElements)
        {
            if (mod is A.Tint tint && tint.Val?.Value != null)
            {
                // OpenXML tint: val=100000 means 100% original color; smaller val blends toward white
                float t = Math.Clamp(tint.Val.Value / 100000f, 0f, 1f);
                byte r = (byte)Math.Clamp((int)Math.Round(255 - (255 - current.Red) * t), 0, 255);
                byte g = (byte)Math.Clamp((int)Math.Round(255 - (255 - current.Green) * t), 0, 255);
                byte b = (byte)Math.Clamp((int)Math.Round(255 - (255 - current.Blue) * t), 0, 255);
                current = new SKColor(r, g, b);
            }
            else if (mod is A.Shade shade && shade.Val?.Value != null)
            {
                float sh = Math.Clamp(shade.Val.Value / 100000f, 0f, 1f);
                byte r = (byte)Math.Clamp((int)Math.Round(current.Red * sh), 0, 255);
                byte g = (byte)Math.Clamp((int)Math.Round(current.Green * sh), 0, 255);
                byte b = (byte)Math.Clamp((int)Math.Round(current.Blue * sh), 0, 255);
                current = new SKColor(r, g, b);
            }
        }

        return current.WithAlpha((byte)Math.Clamp((int)Math.Round(alpha * 255f), 0, 255));
    }

    // =========================================================================
    // 11. Custom & Preset Shape Vector Path Builder (SKPath)
    // =========================================================================

    private static SKPath BuildShapePath(SKRect rect, A.CustomGeometry? custGeom, A.PresetGeometry? prstGeom)
    {
        var path = new SKPath();

        // 1. CustomGeometry (<a:custGeom><a:pathLst><a:path ...>) — supports multi-subpath & ArcTo
        var custPaths = custGeom?.PathList?.Elements<A.Path>().ToList();
        if (custPaths != null && custPaths.Count > 0)
        {
            foreach (var custPath in custPaths)
            {
                float pw = custPath.Width?.Value != null && custPath.Width.Value > 0 ? custPath.Width.Value : rect.Width * (float)EmusPerPoint;
                float ph = custPath.Height?.Value != null && custPath.Height.Value > 0 ? custPath.Height.Value : rect.Height * (float)EmusPerPoint;

                float MapX(string? rawX) => long.TryParse(rawX, out long vx) ? rect.Left + (vx / pw) * rect.Width : rect.Left;
                float MapY(string? rawY) => long.TryParse(rawY, out long vy) ? rect.Top + (vy / ph) * rect.Height : rect.Top;

                float curX = rect.Left, curY = rect.Top;

                foreach (var cmd in custPath.ChildElements)
                {
                    if (cmd is A.MoveTo mv && mv.Point != null)
                    {
                        curX = MapX(mv.Point.X?.Value);
                        curY = MapY(mv.Point.Y?.Value);
                        path.MoveTo(curX, curY);
                    }
                    else if (cmd is A.LineTo ln && ln.Point != null)
                    {
                        curX = MapX(ln.Point.X?.Value);
                        curY = MapY(ln.Point.Y?.Value);
                        path.LineTo(curX, curY);
                    }
                    else if (cmd is A.CubicBezierCurveTo cb)
                    {
                        var pts = cb.Elements<A.Point>().ToList();
                        if (pts.Count == 3)
                        {
                            curX = MapX(pts[2].X?.Value);
                            curY = MapY(pts[2].Y?.Value);
                            path.CubicTo(
                                MapX(pts[0].X?.Value), MapY(pts[0].Y?.Value),
                                MapX(pts[1].X?.Value), MapY(pts[1].Y?.Value),
                                curX, curY);
                        }
                    }
                    else if (cmd is A.QuadraticBezierCurveTo qb)
                    {
                        var pts = qb.Elements<A.Point>().ToList();
                        if (pts.Count == 2)
                        {
                            curX = MapX(pts[1].X?.Value);
                            curY = MapY(pts[1].Y?.Value);
                            path.QuadTo(
                                MapX(pts[0].X?.Value), MapY(pts[0].Y?.Value),
                                curX, curY);
                        }
                    }
                    else if (cmd is A.ArcTo arc)
                    {
                        if (long.TryParse(arc.WidthRadius?.Value, out long wrRaw) &&
                            long.TryParse(arc.HeightRadius?.Value, out long hrRaw) &&
                            long.TryParse(arc.StartAngle?.Value, out long stAngRaw) &&
                            long.TryParse(arc.SwingAngle?.Value, out long swAngRaw) &&
                            wrRaw > 0 && hrRaw > 0)
                        {
                            float rx = (wrRaw / pw) * rect.Width;
                            float ry = (hrRaw / ph) * rect.Height;
                            float startDeg = stAngRaw / 60000f;
                            float sweepDeg = swAngRaw / 60000f;
                            double startRad = startDeg * Math.PI / 180.0;
                            float cx = curX - rx * (float)Math.Cos(startRad);
                            float cy = curY - ry * (float)Math.Sin(startRad);
                            var oval = new SKRect(cx - rx, cy - ry, cx + rx, cy + ry);
                            path.ArcTo(oval, startDeg, sweepDeg, forceMoveTo: false);
                            double endRad = (startDeg + sweepDeg) * Math.PI / 180.0;
                            curX = cx + rx * (float)Math.Cos(endRad);
                            curY = cy + ry * (float)Math.Sin(endRad);
                        }
                    }
                    else if (cmd is A.CloseShapePath)
                    {
                        path.Close();
                    }
                }
            }

            if (path.PointCount > 1)
                return path;
        }

        // 2. PresetGeometry (<a:prstGeom prst="...">)
        var preset = prstGeom?.Preset?.Value;
        if (preset == A.ShapeTypeValues.Ellipse)
        {
            path.AddOval(rect);
            return path;
        }

        if (preset == A.ShapeTypeValues.RoundRectangle)
        {
            float minSide = Math.Min(rect.Width, rect.Height);
            float adjFactor = 0.16667f;
            var gd = prstGeom?.AdjustValueList?.Elements<A.ShapeGuide>().FirstOrDefault();
            if (gd?.Formula?.Value != null && gd.Formula.Value.StartsWith("val ", StringComparison.OrdinalIgnoreCase))
            {
                if (float.TryParse(gd.Formula.Value[4..].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out float rawAdj))
                {
                    adjFactor = Math.Clamp(rawAdj / 100000f, 0f, 0.5f);
                }
            }
            float r = minSide * adjFactor;
            path.AddRoundRect(rect, r, r);
            return path;
        }

        if (preset == A.ShapeTypeValues.Triangle)
        {
            path.MoveTo(rect.MidX, rect.Top);
            path.LineTo(rect.Right, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.RightTriangle)
        {
            path.MoveTo(rect.Left, rect.Top);
            path.LineTo(rect.Right, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Parallelogram)
        {
            float dx = Math.Min(rect.Width, rect.Height) * 0.25f;
            path.MoveTo(rect.Left + dx, rect.Top);
            path.LineTo(rect.Right, rect.Top);
            path.LineTo(rect.Right - dx, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Trapezoid)
        {
            float dx = Math.Min(rect.Width, rect.Height) * 0.25f;
            path.MoveTo(rect.Left + dx, rect.Top);
            path.LineTo(rect.Right - dx, rect.Top);
            path.LineTo(rect.Right, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Diamond)
        {
            path.MoveTo(rect.MidX, rect.Top);
            path.LineTo(rect.Right, rect.MidY);
            path.LineTo(rect.MidX, rect.Bottom);
            path.LineTo(rect.Left, rect.MidY);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Hexagon)
        {
            float adjFactor = 0.25f;
            var gd = prstGeom?.AdjustValueList?.Elements<A.ShapeGuide>()
                .FirstOrDefault(g => string.Equals(g.Name?.Value, "adj", StringComparison.OrdinalIgnoreCase));
            if (gd?.Formula?.Value != null && gd.Formula.Value.StartsWith("val ", StringComparison.OrdinalIgnoreCase))
            {
                if (float.TryParse(gd.Formula.Value[4..].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out float rawAdj))
                {
                    adjFactor = Math.Clamp(rawAdj / 100000f, 0.05f, 0.48f);
                }
            }
            float dx = Math.Min(rect.Width, rect.Height) * adjFactor;
            path.MoveTo(rect.Left + dx, rect.Top);
            path.LineTo(rect.Right - dx, rect.Top);
            path.LineTo(rect.Right, rect.MidY);
            path.LineTo(rect.Right - dx, rect.Bottom);
            path.LineTo(rect.Left + dx, rect.Bottom);
            path.LineTo(rect.Left, rect.MidY);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Pentagon || preset == A.ShapeTypeValues.Heptagon ||
            preset == A.ShapeTypeValues.Octagon || preset == A.ShapeTypeValues.Decagon ||
            preset == A.ShapeTypeValues.Dodecagon)
        {
            int sides = preset == A.ShapeTypeValues.Pentagon ? 5
                : preset == A.ShapeTypeValues.Heptagon ? 7
                : preset == A.ShapeTypeValues.Octagon ? 8
                : preset == A.ShapeTypeValues.Decagon ? 10 : 12;
            AddRegularPolygon(path, rect, sides, startAngleDeg: -90f + (sides % 2 == 0 ? 180f / sides : 0f));
            return path;
        }

        if (preset == A.ShapeTypeValues.Star4 || preset == A.ShapeTypeValues.Star5 ||
            preset == A.ShapeTypeValues.Star6 || preset == A.ShapeTypeValues.Star8)
        {
            int points = preset == A.ShapeTypeValues.Star4 ? 4
                : preset == A.ShapeTypeValues.Star5 ? 5
                : preset == A.ShapeTypeValues.Star6 ? 6 : 8;
            AddStarPolygon(path, rect, points, innerRatio: 0.42f);
            return path;
        }

        if (preset == A.ShapeTypeValues.Plus)
        {
            float arm = Math.Min(rect.Width, rect.Height) * 0.25f;
            float x1 = rect.Left + arm, x2 = rect.Right - arm;
            float y1 = rect.Top + arm, y2 = rect.Bottom - arm;
            path.MoveTo(x1, rect.Top);
            path.LineTo(x2, rect.Top);
            path.LineTo(x2, y1);
            path.LineTo(rect.Right, y1);
            path.LineTo(rect.Right, y2);
            path.LineTo(x2, y2);
            path.LineTo(x2, rect.Bottom);
            path.LineTo(x1, rect.Bottom);
            path.LineTo(x1, y2);
            path.LineTo(rect.Left, y2);
            path.LineTo(rect.Left, y1);
            path.LineTo(x1, y1);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.RightArrow || preset == A.ShapeTypeValues.LeftArrow)
        {
            float headW = Math.Min(rect.Width * 0.35f, rect.Height * 0.6f);
            float shaftH = rect.Height * 0.5f;
            float sy1 = rect.MidY - shaftH * 0.5f, sy2 = rect.MidY + shaftH * 0.5f;
            if (preset == A.ShapeTypeValues.RightArrow)
            {
                path.MoveTo(rect.Left, sy1);
                path.LineTo(rect.Right - headW, sy1);
                path.LineTo(rect.Right - headW, rect.Top);
                path.LineTo(rect.Right, rect.MidY);
                path.LineTo(rect.Right - headW, rect.Bottom);
                path.LineTo(rect.Right - headW, sy2);
                path.LineTo(rect.Left, sy2);
            }
            else
            {
                path.MoveTo(rect.Right, sy1);
                path.LineTo(rect.Left + headW, sy1);
                path.LineTo(rect.Left + headW, rect.Top);
                path.LineTo(rect.Left, rect.MidY);
                path.LineTo(rect.Left + headW, rect.Bottom);
                path.LineTo(rect.Left + headW, sy2);
                path.LineTo(rect.Right, sy2);
            }
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.UpArrow || preset == A.ShapeTypeValues.DownArrow)
        {
            float headH = Math.Min(rect.Height * 0.35f, rect.Width * 0.6f);
            float shaftW = rect.Width * 0.5f;
            float sx1 = rect.MidX - shaftW * 0.5f, sx2 = rect.MidX + shaftW * 0.5f;
            if (preset == A.ShapeTypeValues.UpArrow)
            {
                path.MoveTo(rect.MidX, rect.Top);
                path.LineTo(rect.Right, rect.Top + headH);
                path.LineTo(sx2, rect.Top + headH);
                path.LineTo(sx2, rect.Bottom);
                path.LineTo(sx1, rect.Bottom);
                path.LineTo(sx1, rect.Top + headH);
                path.LineTo(rect.Left, rect.Top + headH);
            }
            else
            {
                path.MoveTo(sx1, rect.Top);
                path.LineTo(sx2, rect.Top);
                path.LineTo(sx2, rect.Bottom - headH);
                path.LineTo(rect.Right, rect.Bottom - headH);
                path.LineTo(rect.MidX, rect.Bottom);
                path.LineTo(rect.Left, rect.Bottom - headH);
                path.LineTo(sx1, rect.Bottom - headH);
            }
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Chevron || preset == A.ShapeTypeValues.HomePlate)
        {
            float notch = Math.Min(rect.Width * 0.22f, rect.Height * 0.5f);
            path.MoveTo(rect.Left, rect.Top);
            path.LineTo(rect.Right - notch, rect.Top);
            path.LineTo(rect.Right, rect.MidY);
            path.LineTo(rect.Right - notch, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            if (preset == A.ShapeTypeValues.Chevron)
                path.LineTo(rect.Left + notch, rect.MidY);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Donut)
        {
            float adjFactor = ReadShapeAdjust(prstGeom, "adj", 0.25f, 0.01f, 0.49f);
            float thickness = Math.Min(rect.Width, rect.Height) * adjFactor;
            path.FillType = SKPathFillType.EvenOdd;
            path.AddOval(rect);
            var inner = new SKRect(rect.Left + thickness, rect.Top + thickness, rect.Right - thickness, rect.Bottom - thickness);
            if (inner.Width > 0.5f && inner.Height > 0.5f)
                path.AddOval(inner);
            return path;
        }

        if (preset == A.ShapeTypeValues.Frame)
        {
            float adjFactor = ReadShapeAdjust(prstGeom, "adj1", 0.125f, 0.005f, 0.49f);
            float thickness = Math.Min(rect.Width, rect.Height) * adjFactor;
            path.FillType = SKPathFillType.EvenOdd;
            path.AddRect(rect);
            var inner = new SKRect(rect.Left + thickness, rect.Top + thickness, rect.Right - thickness, rect.Bottom - thickness);
            if (inner.Width > 0.5f && inner.Height > 0.5f)
                path.AddRect(inner);
            return path;
        }

        if (preset == A.ShapeTypeValues.HalfFrame || preset == A.ShapeTypeValues.Corner)
        {
            float minSide = Math.Min(rect.Width, rect.Height);
            float tY = minSide * ReadShapeAdjust(prstGeom, "adj1", 0.33333f, 0.02f, 0.95f);
            float tX = minSide * ReadShapeAdjust(prstGeom, "adj2", 0.33333f, 0.02f, 0.95f);
            if (preset == A.ShapeTypeValues.HalfFrame)
            {
                path.MoveTo(rect.Left, rect.Top);
                path.LineTo(rect.Right, rect.Top);
                path.LineTo(rect.Right - tX, rect.Top + tY);
                path.LineTo(rect.Left + tX, rect.Top + tY);
                path.LineTo(rect.Left + tX, rect.Bottom - tY);
                path.LineTo(rect.Left, rect.Bottom);
            }
            else
            {
                path.MoveTo(rect.Left, rect.Top);
                path.LineTo(rect.Left + tX, rect.Top);
                path.LineTo(rect.Left + tX, rect.Bottom - tY);
                path.LineTo(rect.Right, rect.Bottom - tY);
                path.LineTo(rect.Right, rect.Bottom);
                path.LineTo(rect.Left, rect.Bottom);
            }
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Heart)
        {
            float w = rect.Width, h = rect.Height;
            float l = rect.Left, t = rect.Top;
            path.MoveTo(l + w * 0.5f, t + h * 0.28f);
            path.CubicTo(l + w * 0.5f, t + h * 0.05f, l + w * 0.15f, t - h * 0.02f, l + w * 0.04f, t + h * 0.24f);
            path.CubicTo(l - w * 0.05f, t + h * 0.48f, l + w * 0.16f, t + h * 0.72f, l + w * 0.5f, t + h);
            path.CubicTo(l + w * 0.84f, t + h * 0.72f, l + w * 1.05f, t + h * 0.48f, l + w * 0.96f, t + h * 0.24f);
            path.CubicTo(l + w * 0.85f, t - h * 0.02f, l + w * 0.5f, t + h * 0.05f, l + w * 0.5f, t + h * 0.28f);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Teardrop)
        {
            float r = Math.Min(rect.Width, rect.Height) * 0.5f;
            path.AddOval(rect);
            path.MoveTo(rect.MidX, rect.Top);
            path.LineTo(rect.Right, rect.Top);
            path.LineTo(rect.Right, rect.MidY);
            path.Close();
            return path;
        }

        if (preset == A.ShapeTypeValues.Line || preset == A.ShapeTypeValues.StraightConnector1)
        {
            path.MoveTo(rect.Left, rect.Top);
            path.LineTo(rect.Right, rect.Bottom);
            return path;
        }

        path.AddRect(rect);
        return path;
    }

    private static float ReadShapeAdjust(A.PresetGeometry? prstGeom, string name, float defaultVal, float minVal, float maxVal)
    {
        var gd = prstGeom?.AdjustValueList?.Elements<A.ShapeGuide>()
            .FirstOrDefault(g => string.Equals(g.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
            ?? prstGeom?.AdjustValueList?.Elements<A.ShapeGuide>().FirstOrDefault();
        if (gd?.Formula?.Value != null && gd.Formula.Value.StartsWith("val ", StringComparison.OrdinalIgnoreCase))
        {
            if (float.TryParse(gd.Formula.Value[4..].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out float rawAdj))
            {
                return Math.Clamp(rawAdj / 100000f, minVal, maxVal);
            }
        }
        return defaultVal;
    }

    private static void AddRegularPolygon(SKPath path, SKRect rect, int sides, float startAngleDeg)
    {
        float rx = rect.Width * 0.5f;
        float ry = rect.Height * 0.5f;
        float cx = rect.MidX;
        float cy = rect.MidY;
        for (int i = 0; i < sides; i++)
        {
            double rad = (startAngleDeg + i * (360.0 / sides)) * Math.PI / 180.0;
            float x = cx + rx * (float)Math.Cos(rad);
            float y = cy + ry * (float)Math.Sin(rad);
            if (i == 0) path.MoveTo(x, y);
            else path.LineTo(x, y);
        }
        path.Close();
    }

    private static void AddStarPolygon(SKPath path, SKRect rect, int points, float innerRatio)
    {
        float rx = rect.Width * 0.5f;
        float ry = rect.Height * 0.5f;
        float cx = rect.MidX;
        float cy = rect.MidY;
        int total = points * 2;
        for (int i = 0; i < total; i++)
        {
            float rScale = (i % 2 == 0) ? 1f : innerRatio;
            double rad = (-90.0 + i * (360.0 / total)) * Math.PI / 180.0;
            float x = cx + rx * rScale * (float)Math.Cos(rad);
            float y = cy + ry * rScale * (float)Math.Sin(rad);
            if (i == 0) path.MoveTo(x, y);
            else path.LineTo(x, y);
        }
        path.Close();
    }

    // =========================================================================
    // 12. Convert PDF Pages / Images -> PowerPoint (.pptx)
    // =========================================================================

    public async Task<ExecutionResult> ConvertSlidesToPptxAsync(
        string originalFileName,
        IReadOnlyList<PdfPageSlideImage> slides,
        PowerPointPdfOptions? options = null,
        IProgress<(int Current, int Total, string Stage)>? progress = null)
    {
        if (slides == null || slides.Count == 0)
        {
            return ExecutionResult.Fail("No PDF pages or slide images were provided.");
        }

        options ??= new PowerPointPdfOptions { Mode = PowerPointConversionMode.PdfToPptx };
        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            try
            {
                using var outStream = new MemoryStream();
                using (var presDoc = PresentationDocument.Create(outStream, PresentationDocumentType.Presentation, true))
                {
                    var presPart = presDoc.AddPresentationPart();
                    presPart.Presentation = new P.Presentation();

                    long slideCxEmu;
                    long slideCyEmu;
                    switch (options.SlideSizePreset)
                    {
                        case PptxSlideSizePreset.Widescreen16x9:
                            slideCxEmu = 12192000L; // 960 pt
                            slideCyEmu = 6858000L;  // 540 pt
                            break;
                        case PptxSlideSizePreset.Standard4x3:
                            slideCxEmu = 9144000L;  // 720 pt
                            slideCyEmu = 6858000L;  // 540 pt
                            break;
                        case PptxSlideSizePreset.MatchSource:
                        default:
                            var first = slides[0];
                            double wPt = first.WidthPoints > 10 ? first.WidthPoints : 960.0;
                            double hPt = first.HeightPoints > 10 ? first.HeightPoints : 540.0;
                            slideCxEmu = (long)Math.Round(wPt * EmusPerPoint);
                            slideCyEmu = (long)Math.Round(hPt * EmusPerPoint);
                            break;
                    }

                    var slideMasterPart = CreateMinimalSlideMasterPart(presPart);
                    var slideLayoutPart = slideMasterPart.SlideLayoutParts.First();
                    string masterRelId = presPart.GetIdOfPart(slideMasterPart);

                    presPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(
                        new P.SlideMasterId { Id = 2147483648U, RelationshipId = masterRelId }
                    );

                    var slideIdList = new P.SlideIdList();
                    presPart.Presentation.SlideIdList = slideIdList;
                    presPart.Presentation.SlideSize = new P.SlideSize
                    {
                        Cx = (int)Math.Clamp(slideCxEmu, 914400L, 51206400L),
                        Cy = (int)Math.Clamp(slideCyEmu, 914400L, 51206400L),
                        Type = options.SlideSizePreset == PptxSlideSizePreset.Widescreen16x9
                            ? P.SlideSizeValues.Screen16x9
                            : options.SlideSizePreset == PptxSlideSizePreset.Standard4x3
                                ? P.SlideSizeValues.Screen4x3
                                : P.SlideSizeValues.Custom
                    };
                    presPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000L, Cy = 9144000L };

                    uint slideIdCounter = 256U;
                    int createdCount = 0;

                    foreach (var pageSlide in slides)
                    {
                        if (pageSlide.ImageBytes == null || pageSlide.ImageBytes.Length == 0)
                            continue;

                        var slidePart = presPart.AddNewPart<SlidePart>();
                        slidePart.AddPart(slideLayoutPart);

                        bool isPng = pageSlide.ImageBytes.Length >= 4 &&
                                     pageSlide.ImageBytes[0] == 0x89 &&
                                     pageSlide.ImageBytes[1] == 0x50;

                        var imgPart = slidePart.AddImagePart(isPng ? ImagePartType.Png : ImagePartType.Jpeg);
                        using (var imgStream = new MemoryStream(pageSlide.ImageBytes, writable: false))
                        {
                            imgPart.FeedData(imgStream);
                        }
                        string imgRelId = slidePart.GetIdOfPart(imgPart);

                        double imgW = pageSlide.WidthPoints > 1 ? pageSlide.WidthPoints : slideCxEmu / EmusPerPoint;
                        double imgH = pageSlide.HeightPoints > 1 ? pageSlide.HeightPoints : slideCyEmu / EmusPerPoint;
                        if (pageSlide.WidthPoints <= 1 || pageSlide.HeightPoints <= 1)
                        {
                            using var skBmp = SKBitmap.Decode(pageSlide.ImageBytes);
                            if (skBmp != null && skBmp.Width > 0 && skBmp.Height > 0)
                            {
                                imgW = skBmp.Width;
                                imgH = skBmp.Height;
                            }
                        }

                        double scale = Math.Min((double)slideCxEmu / imgW, (double)slideCyEmu / imgH);
                        long drawCx = (long)Math.Round(imgW * scale);
                        long drawCy = (long)Math.Round(imgH * scale);
                        long offX = Math.Max(0L, (slideCxEmu - drawCx) / 2L);
                        long offY = Math.Max(0L, (slideCyEmu - drawCy) / 2L);

                        var picture = new P.Picture(
                            new P.NonVisualPictureProperties(
                                new P.NonVisualDrawingProperties { Id = 2U, Name = $"PDF Page {pageSlide.PageNumber}" },
                                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                                new P.ApplicationNonVisualDrawingProperties()),
                            new P.BlipFill(
                                new A.Blip { Embed = imgRelId },
                                new A.Stretch(new A.FillRectangle())),
                            new P.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = offX, Y = offY },
                                    new A.Extents { Cx = drawCx, Cy = drawCy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                        );

                        slidePart.Slide = new P.Slide(
                            new P.CommonSlideData(
                                new P.ShapeTree(
                                    new P.NonVisualGroupShapeProperties(
                                        new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                        new P.NonVisualGroupShapeDrawingProperties(),
                                        new P.ApplicationNonVisualDrawingProperties()),
                                    new P.GroupShapeProperties(new A.TransformGroup(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = 0L, Cy = 0L },
                                        new A.ChildOffset { X = 0L, Y = 0L },
                                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                                    picture)),
                            new P.ColorMapOverride(new A.MasterColorMapping())
                        );
                        slidePart.Slide.Save();

                        slideIdList.AppendChild(new P.SlideId
                        {
                            Id = slideIdCounter++,
                            RelationshipId = presPart.GetIdOfPart(slidePart)
                        });
                        createdCount++;
                        progress?.Report((createdCount, slides.Count, "packaging"));
                    }

                    if (createdCount == 0)
                    {
                        return ExecutionResult.Fail("No valid PDF page images could be embedded into PowerPoint.");
                    }

                    presPart.Presentation.Save();
                }

                sw.Stop();
                string baseName = Path.GetFileNameWithoutExtension(string.IsNullOrWhiteSpace(originalFileName) ? "Presentation.pdf" : originalFileName);
                string fileName = $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                string summary = $"Converted {slides.Count} PDF page(s) to PowerPoint (.pptx) in {sw.ElapsedMilliseconds} ms";

                return ExecutionResult.Ok(
                    outStream.ToArray(),
                    fileName,
                    summary,
                    sw.ElapsedMilliseconds,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation");
            }
            catch (Exception ex)
            {
                sw.Stop();
                return ExecutionResult.Fail($"PDF to PowerPoint error: {ex.Message}", sw.ElapsedMilliseconds);
            }
        });
    }

    private static SlideMasterPart CreateMinimalSlideMasterPart(PresentationPart presPart)
    {
        var slideMasterPart = presPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        // CRITICAL: PowerPoint requires the SlideLayout -> SlideMaster back-relationship!
        slideLayoutPart.AddPart(slideMasterPart);

        // Create ThemePart under PresentationPart so it lives at standard /ppt/theme/theme1.xml
        var themePart = presPart.AddNewPart<ThemePart>();
        slideMasterPart.AddPart(themePart);

        // Standard auxiliary parts expected by Microsoft PowerPoint
        var presPropsPart = presPart.AddNewPart<PresentationPropertiesPart>();
        presPropsPart.PresentationProperties = new P.PresentationProperties();
        presPropsPart.PresentationProperties.Save();

        var viewPropsPart = presPart.AddNewPart<ViewPropertiesPart>();
        viewPropsPart.ViewProperties = new P.ViewProperties();
        viewPropsPart.ViewProperties.Save();

        var tableStylesPart = presPart.AddNewPart<TableStylesPart>();
        tableStylesPart.TableStyleList = new A.TableStyleList { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
        tableStylesPart.TableStyleList.Save();

        themePart.Theme = new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1E293B" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "F8FAFC" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "059669" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "4F46E5" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "D97706" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "DC2626" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "0F172A" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "7C3AED" }))
                { Name = "MnshorTheme" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Segoe UI" }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = "Segoe UI" }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Segoe UI" }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = "Segoe UI" }))
                { Name = "MnshorFonts" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 6350 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 12700 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 19050 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
                { Name = "MnshorFormat" }))
        { Name = "MnshorTheme" };
        themePart.Theme.Save();

        slideLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMapOverride(new A.MasterColorMapping())
        )
        { Type = P.SlideLayoutValues.Blank, Preserve = true };
        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
            )
        );
        slideMasterPart.SlideMaster.Save();

        return slideMasterPart;
    }

    // =========================================================================
    // 13. OpenXML & Skia Utility Methods
    // =========================================================================

    private static (float WidthPt, float HeightPt) GetSlideDimensionsPoints(P.Presentation pres)
    {
        if (pres.SlideSize?.Cx != null && pres.SlideSize?.Cy != null)
        {
            float w = (float)(pres.SlideSize.Cx.Value / EmusPerPoint);
            float h = (float)(pres.SlideSize.Cy.Value / EmusPerPoint);
            if (w > 50f && h > 50f)
                return (w, h);
        }
        return (960f, 540f);
    }

    private static A.Transform2D? GetEffectiveTransform2D(A.Transform2D? slideXfrm, A.Transform2D? layoutXfrm, A.Transform2D? masterXfrm)
    {
        if (slideXfrm?.Offset != null && slideXfrm.Extents != null)
            return slideXfrm;
        if (layoutXfrm?.Offset != null && layoutXfrm.Extents != null)
            return layoutXfrm;
        return masterXfrm ?? slideXfrm ?? layoutXfrm;
    }

    private static SKRect GetFallbackPlaceholderRect(P.PlaceholderValues phType, float slideW, float slideH)
    {
        if (phType == P.PlaceholderValues.CenteredTitle)
            return new SKRect(slideW * 0.08f, slideH * 0.28f, slideW * 0.92f, slideH * 0.54f);
        if (phType == P.PlaceholderValues.Title)
            return new SKRect(slideW * 0.06f, slideH * 0.06f, slideW * 0.94f, slideH * 0.24f);
        if (phType == P.PlaceholderValues.SubTitle)
            return new SKRect(slideW * 0.12f, slideH * 0.56f, slideW * 0.88f, slideH * 0.74f);
        if (phType == P.PlaceholderValues.Footer || phType == P.PlaceholderValues.DateAndTime || phType == P.PlaceholderValues.SlideNumber)
            return new SKRect(slideW * 0.06f, slideH * 0.91f, slideW * 0.94f, slideH * 0.97f);

        return new SKRect(slideW * 0.06f, slideH * 0.26f, slideW * 0.94f, slideH * 0.88f);
    }

    private static float GetEffectiveInset(Int32Value? slideVal, Int32Value? layoutVal, Int32Value? masterVal, float defaultPt)
    {
        int? emu = slideVal?.Value ?? layoutVal?.Value ?? masterVal?.Value;
        return emu.HasValue ? Math.Clamp((float)(emu.Value / EmusPerPoint), 0f, 120f) : defaultPt;
    }

    private static SKRect GetTransformRect(A.Transform2D? xfrm, float offsetX, float offsetY, float scaleX, float scaleY)
    {
        if (xfrm?.Offset == null || xfrm.Extents == null)
            return SKRect.Empty;

        float x = offsetX + (float)((xfrm.Offset.X?.Value ?? 0L) / EmusPerPoint) * scaleX;
        float y = offsetY + (float)((xfrm.Offset.Y?.Value ?? 0L) / EmusPerPoint) * scaleY;
        float w = (float)((xfrm.Extents.Cx?.Value ?? 0L) / EmusPerPoint) * scaleX;
        float h = (float)((xfrm.Extents.Cy?.Value ?? 0L) / EmusPerPoint) * scaleY;

        return new SKRect(x, y, x + w, y + h);
    }

    private static SKRect GetTransformRect(P.Transform? xfrm, float offsetX, float offsetY, float scaleX, float scaleY)
    {
        if (xfrm?.Offset == null || xfrm.Extents == null)
            return SKRect.Empty;

        float x = offsetX + (float)((xfrm.Offset.X?.Value ?? 0L) / EmusPerPoint) * scaleX;
        float y = offsetY + (float)((xfrm.Offset.Y?.Value ?? 0L) / EmusPerPoint) * scaleY;
        float w = (float)((xfrm.Extents.Cx?.Value ?? 0L) / EmusPerPoint) * scaleX;
        float h = (float)((xfrm.Extents.Cy?.Value ?? 0L) / EmusPerPoint) * scaleY;

        return new SKRect(x, y, x + w, y + h);
    }

    private static Dictionary<string, string> ExtractColorMap(
        P.ColorMap? masterMap,
        P.ColorMapOverride? layoutOvr,
        P.ColorMapOverride? slideOvr)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["bg1"] = "lt1",
            ["tx1"] = "dk1",
            ["bg2"] = "lt2",
            ["tx2"] = "dk2",
            ["accent1"] = "accent1",
            ["accent2"] = "accent2",
            ["accent3"] = "accent3",
            ["accent4"] = "accent4",
            ["accent5"] = "accent5",
            ["accent6"] = "accent6",
            ["hlink"] = "hlink",
            ["folHlink"] = "folHlink"
        };

        void ApplyMappingAttributes(OpenXmlElement? el)
        {
            if (el == null) return;
            foreach (var attr in el.GetAttributes())
            {
                if (!string.IsNullOrEmpty(attr.LocalName) && !string.IsNullOrEmpty(attr.Value))
                {
                    map[attr.LocalName] = attr.Value;
                }
            }
        }

        ApplyMappingAttributes(masterMap);
        ApplyMappingAttributes(layoutOvr?.GetFirstChild<A.OverrideColorMapping>());
        ApplyMappingAttributes(slideOvr?.GetFirstChild<A.OverrideColorMapping>());

        return map;
    }

    private static Dictionary<string, SKColor> ExtractThemeColors(ThemePart? themePart)
    {
        var map = new Dictionary<string, SKColor>(StringComparer.OrdinalIgnoreCase)
        {
            ["dk1"] = SKColors.Black,
            ["lt1"] = SKColors.White,
            ["dk2"] = new SKColor(30, 41, 59),
            ["lt2"] = new SKColor(248, 250, 252),
            ["accent1"] = new SKColor(37, 99, 235),
            ["accent2"] = new SKColor(5, 150, 105),
            ["accent3"] = new SKColor(79, 70, 229),
            ["accent4"] = new SKColor(217, 119, 6),
            ["accent5"] = new SKColor(220, 38, 38),
            ["accent6"] = new SKColor(15, 23, 42),
            ["hlink"] = new SKColor(37, 99, 235),
            ["folHlink"] = new SKColor(124, 58, 237),
            ["tx1"] = SKColors.Black,
            ["bg1"] = SKColors.White,
            ["tx2"] = new SKColor(30, 41, 59),
            ["bg2"] = new SKColor(248, 250, 252)
        };

        var cs = themePart?.Theme?.ThemeElements?.ColorScheme;
        if (cs == null)
            return map;

        TrySetThemeSlotColor(map, "dk1", cs.Dark1Color);
        TrySetThemeSlotColor(map, "lt1", cs.Light1Color);
        TrySetThemeSlotColor(map, "dk2", cs.Dark2Color);
        TrySetThemeSlotColor(map, "lt2", cs.Light2Color);
        TrySetThemeSlotColor(map, "accent1", cs.Accent1Color);
        TrySetThemeSlotColor(map, "accent2", cs.Accent2Color);
        TrySetThemeSlotColor(map, "accent3", cs.Accent3Color);
        TrySetThemeSlotColor(map, "accent4", cs.Accent4Color);
        TrySetThemeSlotColor(map, "accent5", cs.Accent5Color);
        TrySetThemeSlotColor(map, "accent6", cs.Accent6Color);
        TrySetThemeSlotColor(map, "hlink", cs.Hyperlink);
        TrySetThemeSlotColor(map, "folHlink", cs.FollowedHyperlinkColor);

        if (map.TryGetValue("dk1", out var dk1)) map["tx1"] = dk1;
        if (map.TryGetValue("lt1", out var lt1)) map["bg1"] = lt1;
        if (map.TryGetValue("dk2", out var dk2)) map["tx2"] = dk2;
        if (map.TryGetValue("lt2", out var lt2)) map["bg2"] = lt2;

        return map;
    }

    private static void TrySetThemeSlotColor(Dictionary<string, SKColor> map, string key, A.Color2Type? colorNode)
    {
        if (colorNode == null) return;
        string? hex = colorNode.RgbColorModelHex?.Val?.Value ?? colorNode.SystemColor?.LastColor?.Value;
        if (!string.IsNullOrEmpty(hex) && SKColor.TryParse("#" + hex, out var parsed))
        {
            map[key] = parsed;
        }
    }

    private static bool IsDarkColor(SKColor c)
    {
        int lum = (c.Red * 299 + c.Green * 587 + c.Blue * 114) / 1000;
        return lum < 135;
    }

    private static string ExtractShapeText(P.Shape shape)
    {
        if (shape.TextBody == null) return string.Empty;
        return string.Join("\n", shape.TextBody.Elements<A.Paragraph>().Select(ExtractDrawingParagraphText));
    }

    private static string ExtractDrawingParagraphText(A.Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is A.Run run && run.Text != null)
            {
                sb.Append(run.Text.Text);
            }
            else if (child is A.Field field && field.Text != null)
            {
                sb.Append(field.Text.Text);
            }
            else if (child is A.Break)
            {
                sb.Append('\n');
            }
        }
        return sb.ToString();
    }

    private static string ExtractNotesText(NotesSlidePart notesPart)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return string.Empty;

        var parts = new List<string>();
        foreach (var shape in spTree.Elements<P.Shape>())
        {
            var phType = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value;

            if (phType == P.PlaceholderValues.SlideNumber || phType == P.PlaceholderValues.Header || phType == P.PlaceholderValues.Footer)
                continue;

            string txt = ExtractShapeText(shape).Trim();
            if (!string.IsNullOrWhiteSpace(txt))
                parts.Add(txt);
        }
        return string.Join(" ", parts);
    }

    private static List<List<(string Word, float Width)>> WrapPlainTextLines(
        string text,
        SKShaper shaper,
        SKFont font,
        float fontSize,
        float maxWidth)
    {
        var result = new List<List<(string Word, float Width)>>();
        if (string.IsNullOrWhiteSpace(text)) return result;

        float spaceW = font.MeasureText(" ");
        foreach (var logical in text.Replace("\r\n", "\n").Split('\n'))
        {
            var words = logical.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) continue;

            var currentLine = new List<(string Word, float Width)>();
            float currentWidth = 0f;

            foreach (var word in words)
            {
                float wWidth = MeasureWordWithEmojis(word, shaper, font, fontSize);
                float projected = currentLine.Count == 0 ? wWidth : currentWidth + spaceW + wWidth;
                if (projected > maxWidth && currentLine.Count > 0)
                {
                    result.Add(currentLine);
                    currentLine = new List<(string Word, float Width)> { (word, wWidth) };
                    currentWidth = wWidth;
                }
                else
                {
                    currentLine.Add((word, wWidth));
                    currentWidth = projected;
                }
            }

            if (currentLine.Count > 0)
                result.Add(currentLine);
        }

        return result;
    }

    private static float MeasureWordWithEmojis(string word, SKShaper shaper, SKFont font, float fontSize)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float width = 0f;
        float emojiAdvance = fontSize * 1.15f;

        foreach (var (tokenText, isEmoji) in tokens)
        {
            width += isEmoji ? emojiAdvance : shaper.Shape(tokenText, font).Width;
        }
        return width;
    }

    private static void DrawPlainVisualLine(
        SKCanvas canvas,
        List<(string Word, float Width)> words,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float fontSize,
        bool isRtl,
        float startX,
        float maxWidth,
        float spaceWidth,
        float baselineY,
        A.TextAlignmentTypeValues alignment)
    {
        float lineWidth = words.Sum(w => w.Width) + (Math.Max(0, words.Count - 1) * spaceWidth);

        if (isRtl)
        {
            float cursorRight = alignment == A.TextAlignmentTypeValues.Center
                ? startX + (maxWidth + lineWidth) / 2f
                : alignment == A.TextAlignmentTypeValues.Left
                    ? startX + lineWidth
                    : startX + maxWidth;

            foreach (var (word, wordWidth) in words)
            {
                float wordLeft = cursorRight - wordWidth;
                DrawWordTokens(canvas, word, shaper, font, paint, fontSize, isRtl: true, wordLeft, wordWidth, baselineY);
                cursorRight = wordLeft - spaceWidth;
            }
        }
        else
        {
            float cursorLeft = alignment == A.TextAlignmentTypeValues.Center
                ? startX + Math.Max(0f, (maxWidth - lineWidth) / 2f)
                : alignment == A.TextAlignmentTypeValues.Right
                    ? startX + Math.Max(0f, maxWidth - lineWidth)
                    : startX;

            foreach (var (word, wordWidth) in words)
            {
                DrawWordTokens(canvas, word, shaper, font, paint, fontSize, isRtl: false, cursorLeft, wordWidth, baselineY);
                cursorLeft += wordWidth + spaceWidth;
            }
        }
    }

    private static void DrawWordTokens(
        SKCanvas canvas,
        string word,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float fontSize,
        bool isRtl,
        float wordLeft,
        float wordWidth,
        float baselineY)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float emojiBoxSize = fontSize * 1.12f;
        float emojiAdvance = fontSize * 1.15f;

        if (isRtl && tokens.Count > 1)
        {
            float tokenRight = wordLeft + wordWidth;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    float x = tokenRight - emojiAdvance;
                    DrawEmojiImage(canvas, tokenText, x, baselineY - (fontSize * 0.85f), emojiBoxSize);
                    tokenRight -= emojiAdvance;
                }
                else
                {
                    var shaped = shaper.Shape(tokenText, font);
                    float x = tokenRight - shaped.Width;
                    canvas.DrawShapedText(shaper, tokenText, x, baselineY, SKTextAlign.Left, font, paint);
                    tokenRight -= shaped.Width;
                }
            }
        }
        else
        {
            float tokenLeft = wordLeft;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    DrawEmojiImage(canvas, tokenText, tokenLeft, baselineY - (fontSize * 0.85f), emojiBoxSize);
                    tokenLeft += emojiAdvance;
                }
                else
                {
                    var shaped = shaper.Shape(tokenText, font);
                    canvas.DrawShapedText(shaper, tokenText, tokenLeft, baselineY, SKTextAlign.Left, font, paint);
                    tokenLeft += shaped.Width;
                }
            }
        }
    }

    private static void DrawEmojiImage(SKCanvas canvas, string emojiCluster, float x, float y, float size)
    {
        var pngBytes = EmojiRenderer.RenderEmojiToPng(emojiCluster, 96);
        if (pngBytes == null || pngBytes.Length == 0) return;

        using var skImage = SKImage.FromEncodedData(pngBytes);
        if (skImage == null) return;

        canvas.DrawImage(skImage, new SKRect(x, y, x + size, y + size), new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
    }
}
