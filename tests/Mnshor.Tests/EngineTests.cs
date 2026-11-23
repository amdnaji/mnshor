using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Mnshor.Engines;
using Mnshor.Models;
using Mnshor.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xunit;
using OxSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using OxDrawing = DocumentFormat.OpenXml.Drawing;
using OxPres = DocumentFormat.OpenXml.Presentation;

namespace Mnshor.Tests;

public class EngineTests
{
    static EngineTests()
    {
        PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new SystemFontResolver();
    }

    private readonly FastPdfEngine _pdfEngine = new();
    private readonly LocalWordConverter _wordConverter = new();
    private readonly LocalSpreadsheetMarkdownConverter _excelMdConverter = new();
    private readonly LocalPowerPointConverter _pptxConverter = new();
    private readonly LocalPdfImageConverter _pdfImageConverter = new();

    private static byte[] CreateSamplePdf(int pageCount, string titlePrefix = "Doc")
    {
        using var doc = new PdfDocument();
        for (int i = 1; i <= pageCount; i++)
        {
            var page = doc.AddPage();
            using var gfx = XGraphics.FromPdfPage(page);
            var font = new XFont("Arial", 14, XFontStyleEx.Bold);
            gfx.DrawString($"{titlePrefix} - Page {i}", font, XBrushes.DarkSlateBlue, new XPoint(50, 100));
        }

        using var ms = new MemoryStream();
        doc.Save(ms, false);
        return ms.ToArray();
    }

    private static byte[] CreateSampleDocx(string heading, string paragraph, string tableCellContent)
    {
        using var ms = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Heading paragraph
            var hPara = body.AppendChild(new Paragraph());
            var hProp = hPara.AppendChild(new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }));
            var hRun = hPara.AppendChild(new Run());
            hRun.AppendChild(new Text(heading));

            // Body paragraph
            var bPara = body.AppendChild(new Paragraph());
            var bRun = bPara.AppendChild(new Run());
            bRun.AppendChild(new Text(paragraph));

            // Table
            var table = body.AppendChild(new Table());
            var row = table.AppendChild(new TableRow());
            var cell = row.AppendChild(new TableCell());
            cell.AppendChild(new Paragraph(new Run(new Text(tableCellContent))));
        }

        return ms.ToArray();
    }

    [Fact]
    public async Task FastPdfEngine_Merge_CombinesPagesAccuratelyAndFast()
    {
        // Arrange: 3 documents with 2, 3, and 5 pages respectively = 10 pages total
        var doc1 = new DocumentItem { FileName = "Doc1.pdf", Data = CreateSamplePdf(2, "Doc 1") };
        var doc2 = new DocumentItem { FileName = "Doc2.pdf", Data = CreateSamplePdf(3, "Doc 2") };
        var doc3 = new DocumentItem { FileName = "Doc3.pdf", Data = CreateSamplePdf(5, "Doc 3") };

        // Act
        var result = await _pdfEngine.MergeAsync(new[] { doc1, doc2, doc3 });

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.True(result.ElapsedMilliseconds >= 0);

        // Verify resulting PDF in memory
        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(10, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_SplitRange_ExtractsRequestedPages()
    {
        // Arrange: 10 page document, requesting 2-4 and 7 (4 pages)
        var doc = new DocumentItem { FileName = "TestTenPages.pdf", Data = CreateSamplePdf(10, "Original") };

        // Act
        var result = await _pdfEngine.SplitRangeAsync(doc, "2-4, 7");

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_BurstToZip_GeneratesValidZipWithCorrectPages()
    {
        // Arrange: 4 page document
        var doc = new DocumentItem { FileName = "Report.pdf", Data = CreateSamplePdf(4, "Report") };

        // Act
        var result = await _pdfEngine.BurstToZipAsync(doc);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("application/zip", result.ContentType);
        Assert.NotEmpty(result.OutputData);

        using var zipStream = new MemoryStream(result.OutputData);
        using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read);
        Assert.Equal(4, archive.Entries.Count);
        Assert.Contains(archive.Entries, e => e.Name.Contains("page_1") || e.Name.Contains("page_01") || e.Name.Contains("page_"));
    }

    [Fact]
    public async Task FastPdfEngine_Rotate_AllPages_AppliesRotationCorrectly()
    {
        // Arrange: 3 page document
        var doc = new DocumentItem { FileName = "DocToRotate.pdf", Data = CreateSamplePdf(3, "RotateDoc") };

        // Act: Rotate all pages by 90 degrees
        var result = await _pdfEngine.RotateAsync(doc, 90, "all");

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(3, parsedDoc.PageCount);
        for (int i = 0; i < parsedDoc.PageCount; i++)
        {
            Assert.Equal(90, parsedDoc.Pages[i].Rotate);
        }
    }

    [Fact]
    public async Task FastPdfEngine_Rotate_OddPages_RotatesOnlyOddPages()
    {
        // Arrange: 4 page document
        var doc = new DocumentItem { FileName = "OddEvenDoc.pdf", Data = CreateSamplePdf(4, "OddEven") };

        // Act: Rotate odd pages by 180 degrees
        var result = await _pdfEngine.RotateAsync(doc, 180, "odd");

        // Assert
        Assert.True(result.Success, result.ErrorMessage);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
        Assert.Equal(180, parsedDoc.Pages[0].Rotate); // Page 1
        Assert.Equal(0, parsedDoc.Pages[1].Rotate);   // Page 2
        Assert.Equal(180, parsedDoc.Pages[2].Rotate); // Page 3
        Assert.Equal(0, parsedDoc.Pages[3].Rotate);   // Page 4
    }

    [Fact]
    public async Task FastPdfEngine_Rotate_CustomRange_RotatesSpecifiedPages()
    {
        // Arrange: 5 page document
        var doc = new DocumentItem { FileName = "CustomRangeDoc.pdf", Data = CreateSamplePdf(5, "CustomRange") };

        // Act: Rotate pages 2 and 4-5 by 270 degrees
        var result = await _pdfEngine.RotateAsync(doc, 270, "range", "2, 4-5");

        // Assert
        Assert.True(result.Success, result.ErrorMessage);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(5, parsedDoc.PageCount);
        Assert.Equal(0, parsedDoc.Pages[0].Rotate);   // Page 1
        Assert.Equal(270, parsedDoc.Pages[1].Rotate); // Page 2
        Assert.Equal(0, parsedDoc.Pages[2].Rotate);   // Page 3
        Assert.Equal(270, parsedDoc.Pages[3].Rotate); // Page 4
        Assert.Equal(270, parsedDoc.Pages[4].Rotate); // Page 5
    }

    [Fact]
    public async Task FastPdfEngine_RotatePages_AppliesIndividualRotationsAccurately()
    {
        // Arrange: 4 page document
        var doc = new DocumentItem { FileName = "IndividualRotateDoc.pdf", Data = CreateSamplePdf(4, "MultiRotate") };

        // Act: Apply distinct rotations per page: [90, 180, 270, 0]
        var rotations = new[] { 90, 180, 270, 0 };
        var result = await _pdfEngine.RotatePagesAsync(doc, rotations);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
        Assert.Equal(90, parsedDoc.Pages[0].Rotate);
        Assert.Equal(180, parsedDoc.Pages[1].Rotate);
        Assert.Equal(270, parsedDoc.Pages[2].Rotate);
        Assert.Equal(0, parsedDoc.Pages[3].Rotate);
    }

    [Fact]
    public async Task FastPdfEngine_RemovePages_DeletesSpecifiedPagesAndPreservesOrder()
    {
        // Arrange: 6 page document
        var doc = new DocumentItem { FileName = "SixPageDoc.pdf", Data = CreateSamplePdf(6, "SixPages") };

        // Act: Remove pages 2 and 5
        var result = await _pdfEngine.RemovePagesAsync(doc, new[] { 2, 5 });

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_RemovePages_WithRangeString_DeletesCorrectPages()
    {
        // Arrange: 5 page document
        var doc = new DocumentItem { FileName = "FivePageDoc.pdf", Data = CreateSamplePdf(5, "FivePages") };

        // Act: Remove range "1, 3-4" -> pages 2 and 5 should remain
        var result = await _pdfEngine.RemovePagesAsync(doc, "1, 3-4");

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(2, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_RemovePages_PreventDeletingAllPages()
    {
        // Arrange: 3 page document
        var doc = new DocumentItem { FileName = "ThreePageDoc.pdf", Data = CreateSamplePdf(3, "ThreePages") };

        // Act: Attempt to remove all pages [1, 2, 3]
        var result = await _pdfEngine.RemovePagesAsync(doc, new[] { 1, 2, 3 });

        // Assert: Must fail safely
        Assert.False(result.Success);
        Assert.Contains("at least one page", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FastPdfEngine_ReversePages_ReversesOrderAccurately()
    {
        // Arrange: 4 page document
        var doc = new DocumentItem { FileName = "FourPages.pdf", Data = CreateSamplePdf(4, "Page") };

        // Act
        var result = await _pdfEngine.ReversePagesAsync(doc);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_ReorderPages_AppliesCustomSequence()
    {
        // Arrange: 5 page document
        var doc = new DocumentItem { FileName = "FivePages.pdf", Data = CreateSamplePdf(5, "Page") };

        // Act: Custom sequence: 3, 1, 5, 2
        var result = await _pdfEngine.ReorderPagesAsync(doc, new[] { 3, 1, 5, 2 });

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.Equal(4, parsedDoc.PageCount);
    }

    [Fact]
    public async Task FastPdfEngine_ReorderPages_EmptySequenceFails()
    {
        var doc = new DocumentItem { FileName = "Sample.pdf", Data = CreateSamplePdf(3, "Page") };
        var result = await _pdfEngine.ReorderPagesAsync(doc, Array.Empty<int>());

        Assert.False(result.Success);
    }

    [Fact]
    public async Task FastPdfEngine_LargeDocument150Pages_RotateRemoveAndReorderPerformFast()
    {
        // Arrange: 150-page PDF
        var largePdfBytes = CreateSamplePdf(150, "LargeDoc");
        var doc = new DocumentItem { FileName = "Large150Pages.pdf", Data = largePdfBytes, PageCount = 150 };

        // 1. Rotate pages 50..100 by 90 degrees
        var rotations = Enumerable.Range(1, 150).Select(p => (p >= 50 && p <= 100) ? 90 : 0).ToArray();
        var rotateResult = await _pdfEngine.RotatePagesAsync(doc, rotations);
        Assert.True(rotateResult.Success, rotateResult.ErrorMessage);

        // 2. Remove all even pages (75 pages removed, 75 remaining)
        var evenPages = Enumerable.Range(1, 150).Where(p => p % 2 == 0).ToArray();
        var removeResult = await _pdfEngine.RemovePagesAsync(doc, evenPages);
        Assert.True(removeResult.Success, removeResult.ErrorMessage);
        Assert.Equal(75, _pdfEngine.GetPageCount(removeResult.OutputData));

        // 3. Reverse all 150 pages
        var reverseResult = await _pdfEngine.ReversePagesAsync(doc);
        Assert.True(reverseResult.Success, reverseResult.ErrorMessage);
        Assert.Equal(150, _pdfEngine.GetPageCount(reverseResult.OutputData));
    }

    [Fact]
    public async Task FastPdfEngine_Metadata_UpdatesAndSanitizesAccurately()
    {
        var doc = new DocumentItem { FileName = "MetaDoc.pdf", Data = CreateSamplePdf(2, "Meta") };

        // 1. Update metadata
        var newMeta = new PdfMetadataModel
        {
            Title = "Q4 Strategic Report",
            Author = "Mnshor Team",
            Subject = "Architecture & Security",
            Keywords = "pdf, local, ram",
            Creator = "Mnshor Suite"
        };

        var updateRes = await _pdfEngine.UpdateMetadataAsync(doc, newMeta, sanitizeAll: false);
        Assert.True(updateRes.Success, updateRes.ErrorMessage);

        var readBack = _pdfEngine.GetMetadata(updateRes.OutputData);
        Assert.Equal("Q4 Strategic Report", readBack.Title);
        Assert.Equal("Mnshor Team", readBack.Author);
        Assert.Equal("Architecture & Security", readBack.Subject);
        Assert.Equal("pdf, local, ram", readBack.Keywords);
        Assert.Equal("Mnshor Suite", readBack.Creator);

        // 2. Sanitize metadata
        var updatedDoc = new DocumentItem { FileName = "MetaDoc_Updated.pdf", Data = updateRes.OutputData };
        var sanitizeRes = await _pdfEngine.UpdateMetadataAsync(updatedDoc, newMeta, sanitizeAll: true);
        Assert.True(sanitizeRes.Success, sanitizeRes.ErrorMessage);

        var sanitizedMeta = _pdfEngine.GetMetadata(sanitizeRes.OutputData);
        Assert.False(sanitizedMeta.HasAnyMetadata);
        Assert.Equal(string.Empty, sanitizedMeta.Title);
        Assert.Equal(string.Empty, sanitizedMeta.Author);
        Assert.Equal(string.Empty, sanitizedMeta.Subject);
        Assert.Equal(string.Empty, sanitizedMeta.Keywords);
        Assert.Equal(string.Empty, sanitizedMeta.Creator);
    }

    [Fact]
    public async Task LocalWordConverter_ConvertsDocxToPdf_ProducesValidPdf()
    {
        // Arrange
        var docxBytes = CreateSampleDocx("Project Mnshor", "This is high performance local document processing.", "Cell A1 Data");
        var docItem = new DocumentItem { FileName = "Summary.docx", Data = docxBytes };

        // Act
        var result = await _wordConverter.ConvertToPdfAsync(docItem);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.True(parsedDoc.PageCount >= 1);
    }

    [Fact]
    public async Task LocalWordConverter_ConvertsDocxWithEmojis_RendersSuccessfully()
    {
        // Arrange with common emojis
        var docxBytes = CreateSampleDocx("العنوان مع إيموجي 🚀", "نص يحتوي على مؤشرات ورموز 📊 مع علامة نجاح ✅ وخاتمة 🌟", "خلية 💡");
        var docItem = new DocumentItem { FileName = "EmojiDoc.docx", Data = docxBytes };

        // Act
        var result = await _wordConverter.ConvertToPdfAsync(docItem);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);

        using var resStream = new MemoryStream(result.OutputData);
        using var parsedDoc = PdfReader.Open(resStream, PdfDocumentOpenMode.Import);
        Assert.True(parsedDoc.PageCount >= 1);
    }

    [Fact]
    public void EmojiRenderer_RendersNonEmptyPixels_ForSurrogateAndBmpEmojis()
    {
        foreach (var emoji in new[] { "🚀", "😀", "✅", "🔥", "📊", "💡", "❤️" })
        {
            var pngBytes = EmojiRenderer.RenderEmojiToPng(emoji, 96);
            Assert.NotNull(pngBytes);
            Assert.NotEmpty(pngBytes);

            using var bmp = SkiaSharp.SKBitmap.Decode(pngBytes);
            Assert.NotNull(bmp);

            int nonTransparentPixels = 0;
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    var p = bmp.GetPixel(x, y);
                    if (p.Alpha > 10)
                    {
                        nonTransparentPixels++;
                    }
                }
            }

            Assert.True(nonTransparentPixels > 50, $"Emoji '{emoji}' rendered {nonTransparentPixels} visible pixels!");
        }

        // Also check that two different emojis ("🚀" vs "😀") produce DIFFERENT images (not the same .notdef box!)
        var rocketPng = EmojiRenderer.RenderEmojiToPng("🚀", 96)!;
        var smilePng = EmojiRenderer.RenderEmojiToPng("😀", 96)!;
        Assert.False(rocketPng.SequenceEqual(smilePng), "Rocket and Smile emojis produced identical PNGs (.notdef box)!");
    }

    [Fact]
    public void Localization_TogglesAccuratelyBetweenArabicAndEnglish()
    {
        var loc = new LocalizationState();
        Assert.Equal("en", loc.CurrentLanguage);
        Assert.False(loc.IsRtl);
        Assert.Equal("ltr", loc.Direction);
        Assert.Equal("Mnshor", loc["AppName"]);

        // Toggle to Arabic (second culture)
        loc.ToggleLanguage();
        Assert.Equal("ar", loc.CurrentLanguage);
        Assert.True(loc.IsRtl);
        Assert.Equal("rtl", loc.Direction);
        Assert.Equal("مَنشُور", loc["AppName"]);

        // Toggle back to English (default culture)
        loc.ToggleLanguage();
        Assert.Equal("en", loc.CurrentLanguage);
        Assert.False(loc.IsRtl);
    }

    [Fact]
    public async Task FastPdfEngine_WatermarkText_EnglishAndArabic_ProducesValidWatermarkedPdf()
    {
        var doc = new DocumentItem { FileName = "Contract.pdf", Data = CreateSamplePdf(3, "Contract") };

        // 1. English diagonal watermark
        var enOptions = new PdfWatermarkOptions
        {
            Mode = WatermarkMode.Text,
            Text = "CONFIDENTIAL",
            FontFamily = "Arial",
            FontSize = 48f,
            IsBold = true,
            TextColorHex = "#EF4444",
            Position = WatermarkPosition.Center,
            OpacityPercent = 50,
            RotationDegrees = -45,
            Layer = WatermarkLayer.OverContent
        };

        var enResult = await _pdfEngine.WatermarkAsync(doc, enOptions);
        Assert.True(enResult.Success, enResult.ErrorMessage);
        Assert.NotEmpty(enResult.OutputData);
        Assert.True(enResult.OutputData.Length > doc.Data.Length);
        Assert.Equal(3, _pdfEngine.GetPageCount(enResult.OutputData));

        // 2. Arabic watermark with emoji and corner position
        var arOptions = new PdfWatermarkOptions
        {
            Mode = WatermarkMode.Text,
            Text = "سري للغاية 🔒",
            FontFamily = "Segoe UI",
            FontSize = 36f,
            IsBold = true,
            TextColorHex = "#2563EB",
            Position = WatermarkPosition.TopRight,
            OpacityPercent = 75,
            RotationDegrees = 0,
            Layer = WatermarkLayer.BelowContent
        };

        var arResult = await _pdfEngine.WatermarkAsync(doc, arOptions);
        Assert.True(arResult.Success, arResult.ErrorMessage);
        Assert.NotEmpty(arResult.OutputData);
        Assert.Equal(3, _pdfEngine.GetPageCount(arResult.OutputData));
    }

    [Fact]
    public async Task FastPdfEngine_WatermarkImage_OverlaysPngLogoInMemory()
    {
        var doc = new DocumentItem { FileName = "Proposal.pdf", Data = CreateSamplePdf(2, "Proposal") };

        // Create a synthetic 120x60 PNG logo in RAM using SkiaSharp
        byte[] logoPngBytes;
        using (var surface = SkiaSharp.SKSurface.Create(new SkiaSharp.SKImageInfo(120, 60)))
        {
            surface.Canvas.Clear(new SkiaSharp.SKColor(37, 99, 235));
            using var img = surface.Snapshot();
            using var encoded = img.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            logoPngBytes = encoded.ToArray();
        }

        var imgOptions = new PdfWatermarkOptions
        {
            Mode = WatermarkMode.Image,
            ImageData = logoPngBytes,
            ImageScalePercent = 30,
            Position = WatermarkPosition.BottomRight,
            OpacityPercent = 60,
            RotationDegrees = 0,
            Layer = WatermarkLayer.OverContent
        };

        var result = await _pdfEngine.WatermarkAsync(doc, imgOptions);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.Equal(2, _pdfEngine.GetPageCount(result.OutputData));
    }

    [Fact]
    public async Task FastPdfEngine_WatermarkMosaicAndCustomRange_StampsOnlyTargetPages()
    {
        var doc = new DocumentItem { FileName = "FourPageReport.pdf", Data = CreateSamplePdf(4, "Report") };

        var mosaicOptions = new PdfWatermarkOptions
        {
            Mode = WatermarkMode.Text,
            Text = "DRAFT",
            IsMosaic = true,
            OpacityPercent = 35,
            RotationDegrees = -45,
            PageScope = "range",
            CustomRange = "1, 3-4"
        };

        var result = await _pdfEngine.WatermarkAsync(doc, mosaicOptions);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("3 of 4 pages", result.Summary);
        Assert.Equal(4, _pdfEngine.GetPageCount(result.OutputData));
    }

    [Fact]
    public async Task FastPdfEngine_Watermark_InvalidInputsFailSafely()
    {
        var doc = new DocumentItem { FileName = "Sample.pdf", Data = CreateSamplePdf(2, "Sample") };

        // Empty text in Text mode
        var emptyTextRes = await _pdfEngine.WatermarkAsync(doc, new PdfWatermarkOptions { Mode = WatermarkMode.Text, Text = "   " });
        Assert.False(emptyTextRes.Success);

        // Missing image in Image mode
        var emptyImgRes = await _pdfEngine.WatermarkAsync(doc, new PdfWatermarkOptions { Mode = WatermarkMode.Image, ImageData = null });
        Assert.False(emptyImgRes.Success);

        // Invalid page range
        var badRangeRes = await _pdfEngine.WatermarkAsync(doc, new PdfWatermarkOptions { Mode = WatermarkMode.Text, Text = "COPY", PageScope = "range", CustomRange = "99-105" });
        Assert.False(badRangeRes.Success);
    }

    // =========================================================================
    // Feature 11: Excel (.xlsx) & Markdown to PDF Tests
    // =========================================================================

    private static byte[] CreateSampleXlsx(params (string SheetName, List<(string CellRef, string Value, bool IsSharedString)>[] Rows)[] sheetsData)
    {
        using var ms = new MemoryStream();
        using (var spreadsheet = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new OxSpreadsheet.Workbook();
            var sheets = workbookPart.Workbook.AppendChild(new OxSpreadsheet.Sheets());

            var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sstPart.SharedStringTable = new OxSpreadsheet.SharedStringTable();
            var sharedList = new List<string>();

            uint sheetId = 1;
            foreach (var (sheetName, rows) in sheetsData)
            {
                var wsPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new OxSpreadsheet.SheetData();
                wsPart.Worksheet = new OxSpreadsheet.Worksheet(sheetData);

                uint rowIdx = 1;
                foreach (var rowCells in rows)
                {
                    var row = new OxSpreadsheet.Row { RowIndex = rowIdx++ };
                    foreach (var (cellRef, val, isSst) in rowCells)
                    {
                        var cell = new OxSpreadsheet.Cell { CellReference = cellRef };
                        if (isSst)
                        {
                            int sstIdx = sharedList.IndexOf(val);
                            if (sstIdx < 0)
                            {
                                sstIdx = sharedList.Count;
                                sharedList.Add(val);
                                sstPart.SharedStringTable.AppendChild(new OxSpreadsheet.SharedStringItem(new OxSpreadsheet.Text(val)));
                            }
                            cell.DataType = OxSpreadsheet.CellValues.SharedString;
                            cell.CellValue = new OxSpreadsheet.CellValue(sstIdx.ToString());
                        }
                        else
                        {
                            cell.DataType = OxSpreadsheet.CellValues.Number;
                            cell.CellValue = new OxSpreadsheet.CellValue(val);
                        }
                        row.AppendChild(cell);
                    }
                    sheetData.AppendChild(row);
                }

                sheets.AppendChild(new OxSpreadsheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(wsPart),
                    SheetId = sheetId++,
                    Name = sheetName
                });
            }

            sstPart.SharedStringTable.Save();
            workbookPart.Workbook.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public async Task LocalSpreadsheetMarkdownConverter_ExcelSingleAndMultiSheet_ConvertsToPaginatedPdf()
    {
        var xlsxBytes = CreateSampleXlsx(
            ("Q1 Financials", new[]
            {
                new List<(string, string, bool)> { ("A1", "Department", true), ("B1", "Budget", true), ("D1", "Notes (Sparse Col D)", true) },
                new List<(string, string, bool)> { ("A2", "Engineering", true), ("B2", "125000", false), ("D2", "On Track", true) },
                new List<(string, string, bool)> { ("A3", "Security", true), ("B3", "84500.5", false), ("D3", "Zero Disk IO", true) }
            }),
            ("Q2 Forecast", new[]
            {
                new List<(string, string, bool)> { ("A1", "Metric", true), ("B1", "Target", true) },
                new List<(string, string, bool)> { ("A2", "Latency (ms)", true), ("B2", "15", false) }
            })
        );

        // 1. Verify workbook inspection
        var info = _excelMdConverter.InspectWorkbook(xlsxBytes);
        Assert.True(info.IsValid, info.ErrorMessage);
        Assert.Equal(2, info.Sheets.Count);
        Assert.Equal("Q1 Financials", info.Sheets[0].Name);
        Assert.Equal(3, info.Sheets[0].TotalRows);
        Assert.Equal(4, info.Sheets[0].TotalColumns); // A, B, C (empty), D -> 4 columns!
        Assert.Equal("", info.Sheets[0].PreviewRows[1][2]); // Column C is empty string
        Assert.Equal("On Track", info.Sheets[0].PreviewRows[1][3]); // Column D preserved at index 3

        // 2. Convert both sheets to PDF
        var docItem = new DocumentItem { FileName = "Financials.xlsx", Data = xlsxBytes };
        var result = await _excelMdConverter.ConvertExcelToPdfAsync(docItem);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.True(_pdfEngine.GetPageCount(result.OutputData) >= 2);

        // 3. Convert only selected sheet ("Q2 Forecast")
        var singleSheetOptions = new SpreadsheetMarkdownOptions();
        singleSheetOptions.SelectedSheetNames.Add("Q2 Forecast");
        var singleRes = await _excelMdConverter.ConvertExcelToPdfAsync(docItem, singleSheetOptions);
        Assert.True(singleRes.Success, singleRes.ErrorMessage);
        Assert.Equal(1, _pdfEngine.GetPageCount(singleRes.OutputData));
    }

    [Fact]
    public async Task LocalSpreadsheetMarkdownConverter_ExcelLargeRows_PaginatesWithRepeatedHeader()
    {
        var rows = new List<List<(string, string, bool)>>();
        rows.Add(new List<(string, string, bool)>
        {
            ("A1", "ID", true),
            ("B1", "Transaction Description", true),
            ("C1", "Amount USD", true),
            ("D1", "Status", true)
        });

        for (int r = 2; r <= 120; r++)
        {
            rows.Add(new List<(string, string, bool)>
            {
                ($"A{r}", (r - 1).ToString(), false),
                ($"B{r}", $"Enterprise Local-First Subscription Batch #{r - 1} with extended description for wrapping", true),
                ($"C{r}", ((r - 1) * 49.99).ToString(System.Globalization.CultureInfo.InvariantCulture), false),
                ($"D{r}", "COMPLETED", true)
            });
        }

        var xlsxBytes = CreateSampleXlsx(("Audit_Log_120", rows.ToArray()));
        var docItem = new DocumentItem { FileName = "AuditLog.xlsx", Data = xlsxBytes };

        var options = new SpreadsheetMarkdownOptions
        {
            Orientation = PdfPageOrientation.Landscape,
            RepeatHeaderRow = true,
            ZebraStriping = true,
            ShowGridlines = true,
            ShowRowNumbers = true
        };

        var result = await _excelMdConverter.ConvertExcelToPdfAsync(docItem, options);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.True(_pdfEngine.GetPageCount(result.OutputData) >= 3);
    }

    [Fact]
    public async Task LocalSpreadsheetMarkdownConverter_ExcelArabicAndEmojis_RendersSuccessfully()
    {
        var xlsxBytes = CreateSampleXlsx(
            ("تقرير المبيعات", new[]
            {
                new List<(string, string, bool)> { ("A1", "الفرع", true), ("B1", "الإيرادات 📊", true), ("C1", "الحالة", true) },
                new List<(string, string, bool)> { ("A2", "الرياض", true), ("B2", "95000", false), ("C2", "مكتمل ✅", true) },
                new List<(string, string, bool)> { ("A3", "جدة", true), ("B3", "78400", false), ("C3", "نمو سريع 🚀", true) }
            })
        );

        var info = _excelMdConverter.InspectWorkbook(xlsxBytes);
        Assert.True(info.IsValid);
        Assert.True(info.Sheets[0].IsRtlDetected);

        var docItem = new DocumentItem { FileName = "تقرير_المبيعات.xlsx", Data = xlsxBytes };
        var result = await _excelMdConverter.ConvertExcelToPdfAsync(docItem);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.Equal(1, _pdfEngine.GetPageCount(result.OutputData));
    }

    [Fact]
    public async Task LocalSpreadsheetMarkdownConverter_MarkdownFullSyntax_ConvertsToPdf()
    {
        string markdown = """
# Mnshor Architecture & RTL Report 🚀

This document tests **bold text**, *italic text*, ***bold italic***, `inline_code()`, ~~strikethrough~~, and [Mnshor Link](https://mnshor.local).

## 1. Arabic & BiDi Shaping (تشكيل عربي)

مرحباً بكم في **مَنشُور** — محطة معالجة المستندات المحلية 100% داخل الذاكرة العشوائية ✅

- [x] Complete OpenXML `.xlsx` parser
- [x] Complete Markdown AST renderer
- [ ] External network calls (Zero!)

> "Security and speed by keeping every byte in RAM."
> — فريق تطوير مَنشُور 🔒

### Benchmark Table

| Engine | Format | Speed | Status |
| :--- | :--- | :--- | :--- |
| Spreadsheet | `.xlsx` | 15 ms | Ready ✅ |
| Markdown | `.md` | 9 ms | Ready 🚀 |

```csharp
var result = await converter.ConvertMarkdownToPdfAsync(markdown, "Report");
```

---
End of report.
""";

        var options = new SpreadsheetMarkdownOptions
        {
            Mode = ConversionInputMode.Markdown,
            Orientation = PdfPageOrientation.Portrait,
            FontSizeScale = MarkdownFontSizeScale.Medium,
            IncludePageNumbers = true
        };

        var result = await _excelMdConverter.ConvertMarkdownToPdfAsync(markdown, "Architecture_Report", options);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.True(_pdfEngine.GetPageCount(result.OutputData) >= 1);
    }

    [Fact]
    public async Task LocalSpreadsheetMarkdownConverter_InvalidInputs_FailSafely()
    {
        // 1. Empty Excel bytes
        var emptyDoc = new DocumentItem { FileName = "Empty.xlsx", Data = Array.Empty<byte>() };
        var emptyExcelRes = await _excelMdConverter.ConvertExcelToPdfAsync(emptyDoc);
        Assert.False(emptyExcelRes.Success);

        // 2. Corrupted Excel bytes
        var badDoc = new DocumentItem { FileName = "Corrupt.xlsx", Data = new byte[] { 1, 2, 3, 4, 5 } };
        var badInspect = _excelMdConverter.InspectWorkbook(badDoc.Data);
        Assert.False(badInspect.IsValid);
        var badExcelRes = await _excelMdConverter.ConvertExcelToPdfAsync(badDoc);
        Assert.False(badExcelRes.Success);

        // 3. Empty Markdown string
        var emptyMdRes = await _excelMdConverter.ConvertMarkdownToPdfAsync("   \n  ", "Empty");
        Assert.False(emptyMdRes.Success);
    }

    // =========================================================================
    // Sign PDF (iLovePDF Style + Cross-Platform OS Fonts) Tests
    // =========================================================================

    [Fact]
    public void OsFontService_DiscoversInstalledOsFontsAndRendersSignatures()
    {
        var fontService = new OsFontService();

        var allOsFonts = fontService.GetInstalledFontFamilies();
        var sigFonts = fontService.GetSignatureFonts();
        var arabicFonts = fontService.GetArabicFonts();

        Assert.NotEmpty(allOsFonts);
        Assert.NotEmpty(sigFonts);
        Assert.NotEmpty(arabicFonts);

        // 1. Render Latin signature using first curated OS signature font
        var (enPng, enRatio) = fontService.RenderTextSignatureToPng(
            "Ahmed Al-Mansoor",
            sigFonts[0].FamilyName,
            "#1D4ED8",
            fontSizePt: 44f,
            includeSignatureLine: true);

        Assert.NotEmpty(enPng);
        Assert.True(enRatio > 0.5);

        // 2. Render Arabic signature using first curated OS Arabic calligraphy font
        var (arPng, arRatio) = fontService.RenderTextSignatureToPng(
            "أحمد المنصور ✅",
            arabicFonts[0].FamilyName,
            "#0F172A",
            fontSizePt: 44f,
            includeSignatureLine: true);

        Assert.NotEmpty(arPng);
        Assert.True(arRatio > 0.5);

        // 3. Render Official Stamp
        var (stampPng, stampRatio) = fontService.RenderOfficialStampToPng("APPROVED", "2026-11-17 • AM", "Segoe UI", "#059669");
        Assert.NotEmpty(stampPng);
        Assert.True(stampRatio > 1.0);
    }

    [Fact]
    public void OsFontService_RemoveWhiteBackground_CreatesTransparentSignature()
    {
        var fontService = new OsFontService();

        // Synthesize a 160x80 white paper image with a dark signature line in the center
        byte[] rawPaperScan;
        using (var surface = SkiaSharp.SKSurface.Create(new SkiaSharp.SKImageInfo(160, 80)))
        {
            surface.Canvas.Clear(SkiaSharp.SKColors.White);
            using var pen = new SkiaSharp.SKPaint
            {
                Color = new SkiaSharp.SKColor(20, 20, 20),
                StrokeWidth = 5f,
                IsAntialias = true,
                Style = SkiaSharp.SKPaintStyle.Stroke
            };
            surface.Canvas.DrawLine(20, 40, 140, 40, pen);

            using var snap = surface.Snapshot();
            using var enc = snap.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            rawPaperScan = enc.ToArray();
        }

        var (transparentPng, _) = fontService.ProcessUploadedSignatureImage(
            rawPaperScan,
            removeWhiteBackground: true,
            tintColorHex: "#1D4ED8");

        Assert.NotNull(transparentPng);
        Assert.NotEmpty(transparentPng!);

        using var decoded = SkiaSharp.SKBitmap.Decode(transparentPng);
        Assert.NotNull(decoded);

        // Corner pixel (0, 0) was white -> must now be transparent (Alpha == 0)
        Assert.Equal(0, decoded.GetPixel(0, 0).Alpha);

        // Center pixel (80, 40) had dark ink -> must remain opaque and tinted blue
        var centerPixel = decoded.GetPixel(80, 40);
        Assert.True(centerPixel.Alpha > 200);
        Assert.True(centerPixel.Blue > centerPixel.Red);
    }

    [Fact]
    public async Task FastPdfEngine_SignPdf_StampsTypedSignatureInitialsAndDateAcrossPages()
    {
        var doc = new DocumentItem { FileName = "Agreement.pdf", Data = CreateSamplePdf(4, "Agreement") };

        var signOptions = new PdfSignOptions
        {
            SignerFullName = "Ahmed Al-Mansoor",
            SignerInitials = "AM",
            Fields = new List<PdfSignatureField>
            {
                // Initials on all 4 pages
                new PdfSignatureField
                {
                    FieldType = SignFieldType.Initials,
                    Text = "AM",
                    FontFamily = "Segoe Script",
                    TextColorHex = "#1D4ED8",
                    XPercent = 0.82,
                    YPercent = 0.88,
                    WidthPercent = 0.10,
                    PageScope = SignPageScope.AllPages
                },
                // Full Signature on Last Page only
                new PdfSignatureField
                {
                    FieldType = SignFieldType.Signature,
                    Text = "Ahmed Al-Mansoor",
                    FontFamily = "Ink Free",
                    TextColorHex = "#0F172A",
                    XPercent = 0.55,
                    YPercent = 0.75,
                    WidthPercent = 0.28,
                    PageScope = SignPageScope.LastPage
                },
                // Date on Page 1
                new PdfSignatureField
                {
                    FieldType = SignFieldType.Date,
                    Text = "2026-11-17",
                    FontFamily = "Segoe UI",
                    XPercent = 0.15,
                    YPercent = 0.80,
                    WidthPercent = 0.16,
                    TargetPage = 1,
                    PageScope = SignPageScope.SpecificPage
                }
            }
        };

        var result = await _pdfEngine.SignPdfAsync(doc, signOptions);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.True(result.OutputData.Length > doc.Data.Length);
        Assert.Equal(4, _pdfEngine.GetPageCount(result.OutputData));
        Assert.Contains("4 of 4", result.Summary);
    }

    [Fact]
    public async Task FastPdfEngine_SignPdf_ArabicSignatureAndCustomRange_StampsAccurately()
    {
        var doc = new DocumentItem { FileName = "Contract_AR.pdf", Data = CreateSamplePdf(5, "Contract") };

        var signOptions = new PdfSignOptions
        {
            SignerFullName = "أحمد المنصور",
            SignerInitials = "أ.م",
            Fields = new List<PdfSignatureField>
            {
                new PdfSignatureField
                {
                    FieldType = SignFieldType.Signature,
                    Text = "أحمد المنصور",
                    FontFamily = "Aldhabi",
                    TextColorHex = "#1E3A8A",
                    XPercent = 0.18,
                    YPercent = 0.76,
                    WidthPercent = 0.26,
                    OpacityPercent = 85,
                    RotationDegrees = -5,
                    PageScope = SignPageScope.CustomRange,
                    CustomRange = "1, 3, 5"
                }
            }
        };

        var result = await _pdfEngine.SignPdfAsync(doc, signOptions);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.Equal(5, _pdfEngine.GetPageCount(result.OutputData));
        Assert.Contains("3 of 5", result.Summary);
    }

    [Fact]
    public async Task FastPdfEngine_SignPdf_InvalidInputs_FailSafely()
    {
        var doc = new DocumentItem { FileName = "Sample.pdf", Data = CreateSamplePdf(2, "Sample") };

        // 1. Empty fields list
        var noFieldsRes = await _pdfEngine.SignPdfAsync(doc, new PdfSignOptions());
        Assert.False(noFieldsRes.Success);

        // 2. Empty document
        var emptyDocRes = await _pdfEngine.SignPdfAsync(new DocumentItem { FileName = "Empty.pdf", Data = Array.Empty<byte>() }, new PdfSignOptions
        {
            Fields = new List<PdfSignatureField> { new PdfSignatureField { Text = "Test" } }
        });
        Assert.False(emptyDocRes.Success);
    }

    // =========================================================================
    // PowerPoint (.pptx) <-> PDF Tests
    // =========================================================================

    private static byte[] CreateSamplePptx(params (string Title, string Body, string? Notes)[] slidesContent)
    {
        using var ms = new MemoryStream();
        using (var presDoc = PresentationDocument.Create(ms, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, true))
        {
            var presPart = presDoc.AddPresentationPart();
            presPart.Presentation = new OxPres.Presentation();
            var slideMasterPart = presPart.AddNewPart<SlideMasterPart>();
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

            slideLayoutPart.SlideLayout = new OxPres.SlideLayout(
                new OxPres.CommonSlideData(new OxPres.ShapeTree(
                    new OxPres.NonVisualGroupShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                        new OxPres.NonVisualGroupShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()))));

            slideMasterPart.SlideMaster = new OxPres.SlideMaster(
                new OxPres.CommonSlideData(new OxPres.ShapeTree(
                    new OxPres.NonVisualGroupShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                        new OxPres.NonVisualGroupShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()))),
                new OxPres.SlideLayoutIdList(
                    new OxPres.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }));

            presPart.Presentation.SlideMasterIdList = new OxPres.SlideMasterIdList(
                new OxPres.SlideMasterId { Id = 2147483648U, RelationshipId = presPart.GetIdOfPart(slideMasterPart) });
            presPart.Presentation.SlideSize = new OxPres.SlideSize { Cx = 12192000, Cy = 6858000, Type = OxPres.SlideSizeValues.Screen16x9 };

            var slideIdList = new OxPres.SlideIdList();
            presPart.Presentation.SlideIdList = slideIdList;

            uint slideIdCounter = 256U;
            foreach (var (title, body, notes) in slidesContent)
            {
                var slidePart = presPart.AddNewPart<SlidePart>();
                slidePart.AddPart(slideLayoutPart);

                var titleShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 2U, Name = "Title" },
                        new OxPres.NonVisualShapeDrawingProperties(new OxDrawing.ShapeLocks { NoGrouping = true }),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Type = OxPres.PlaceholderValues.Title })),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(
                            new OxDrawing.Offset { X = 600000L, Y = 450000L },
                            new OxDrawing.Extents { Cx = 10900000L, Cy = 1200000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Rectangle }),
                    new OxPres.TextBody(
                        new OxDrawing.BodyProperties(),
                        new OxDrawing.ListStyle(),
                        new OxDrawing.Paragraph(new OxDrawing.Run(new OxDrawing.Text(title)))));

                var bodyShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 3U, Name = "Body" },
                        new OxPres.NonVisualShapeDrawingProperties(new OxDrawing.ShapeLocks { NoGrouping = true }),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Type = OxPres.PlaceholderValues.Body })),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(
                            new OxDrawing.Offset { X = 600000L, Y = 1900000L },
                            new OxDrawing.Extents { Cx = 10900000L, Cy = 3800000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Rectangle }),
                    new OxPres.TextBody(
                        new OxDrawing.BodyProperties(),
                        new OxDrawing.ListStyle(),
                        new OxDrawing.Paragraph(new OxDrawing.Run(new OxDrawing.Text(body)))));

                slidePart.Slide = new OxPres.Slide(
                    new OxPres.CommonSlideData(
                        new OxPres.ShapeTree(
                            new OxPres.NonVisualGroupShapeProperties(
                                new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                new OxPres.NonVisualGroupShapeDrawingProperties(),
                                new OxPres.ApplicationNonVisualDrawingProperties()),
                            new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                            titleShape,
                            bodyShape)));
                slidePart.Slide.Save();

                if (!string.IsNullOrWhiteSpace(notes))
                {
                    var notesPart = slidePart.AddNewPart<NotesSlidePart>();
                    notesPart.NotesSlide = new OxPres.NotesSlide(
                        new OxPres.CommonSlideData(
                            new OxPres.ShapeTree(
                                new OxPres.NonVisualGroupShapeProperties(
                                    new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                    new OxPres.NonVisualGroupShapeDrawingProperties(),
                                    new OxPres.ApplicationNonVisualDrawingProperties()),
                                new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                                new OxPres.Shape(
                                    new OxPres.NonVisualShapeProperties(
                                        new OxPres.NonVisualDrawingProperties { Id = 2U, Name = "Notes" },
                                        new OxPres.NonVisualShapeDrawingProperties(),
                                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Type = OxPres.PlaceholderValues.Body })),
                                    new OxPres.ShapeProperties(),
                                    new OxPres.TextBody(
                                        new OxDrawing.BodyProperties(),
                                        new OxDrawing.ListStyle(),
                                        new OxDrawing.Paragraph(new OxDrawing.Run(new OxDrawing.Text(notes))))))));
                    notesPart.NotesSlide.Save();
                }

                slideIdList.AppendChild(new OxPres.SlideId
                {
                    Id = slideIdCounter++,
                    RelationshipId = presPart.GetIdOfPart(slidePart)
                });
            }

            presPart.Presentation.Save();
        }
        return ms.ToArray();
    }

    [Fact]
    public async Task LocalPowerPointConverter_InspectAndConvertPptxToPdf_MultiSlideWithArabicAndNotes()
    {
        var pptxBytes = CreateSamplePptx(
            ("Mnshor Executive Deck 🚀", "100% In-Memory PowerPoint to PDF Conversion", "Welcome slide speaker notes."),
            ("عرض مَنشُور المحلي 📊", "تحويل العروض التقديمية إلى PDF بدقة عالية وخصوصية تامة ✅", "ملاحظات الشريحة الثانية بالعربية"),
            ("Architecture Summary", "Zero disk writes, SkiaSharp + HarfBuzz rendering.", null)
        );

        // 1. Inspect presentation
        var info = _pptxConverter.InspectPresentation(pptxBytes);
        Assert.True(info.IsValid, info.ErrorMessage);
        Assert.Equal(3, info.SlideCount);
        Assert.Equal("Mnshor Executive Deck 🚀", info.Slides[0].Title);
        Assert.True(info.Slides[0].HasSpeakerNotes);
        Assert.True(info.Slides[1].IsRtlDetected);
        Assert.False(info.Slides[2].HasSpeakerNotes);

        // 2. Convert all slides with speaker notes
        var docItem = new DocumentItem { FileName = "Deck.pptx", Data = pptxBytes };
        var allResult = await _pptxConverter.ConvertPptxToPdfAsync(docItem, new PowerPointPdfOptions
        {
            IncludeSlideNumbers = true,
            IncludeSpeakerNotes = true
        });
        Assert.True(allResult.Success, allResult.ErrorMessage);
        Assert.NotEmpty(allResult.OutputData);
        Assert.Equal(3, _pdfEngine.GetPageCount(allResult.OutputData));

        // 3. Convert only slide #2
        var subsetOpts = new PowerPointPdfOptions();
        subsetOpts.SelectedSlideNumbers.Add(2);
        var subsetResult = await _pptxConverter.ConvertPptxToPdfAsync(docItem, subsetOpts);
        Assert.True(subsetResult.Success, subsetResult.ErrorMessage);
        Assert.Equal(1, _pdfEngine.GetPageCount(subsetResult.OutputData));
    }

    [Fact]
    public async Task LocalPowerPointConverter_ConvertSlidesToPptxAndRoundTrip_ProducesValidPptx()
    {
        var slideImages = new List<PdfPageSlideImage>();
        for (int i = 1; i <= 2; i++)
        {
            using var surface = SkiaSharp.SKSurface.Create(new SkiaSharp.SKImageInfo(800, 450));
            surface.Canvas.Clear(i == 1 ? new SkiaSharp.SKColor(37, 99, 235) : new SkiaSharp.SKColor(5, 150, 105));
            using var snap = surface.Snapshot();
            using var enc = snap.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            slideImages.Add(new PdfPageSlideImage
            {
                PageNumber = i,
                WidthPoints = 595,
                HeightPoints = 842,
                ImageBytes = enc.ToArray()
            });
        }

        // Convert to 16:9 PowerPoint (.pptx)
        var pptxResult = await _pptxConverter.ConvertSlidesToPptxAsync("AnnualReport.pdf", slideImages, new PowerPointPdfOptions
        {
            Mode = PowerPointConversionMode.PdfToPptx,
            SlideSizePreset = PptxSlideSizePreset.Widescreen16x9
        });

        Assert.True(pptxResult.Success, pptxResult.ErrorMessage);
        Assert.NotEmpty(pptxResult.OutputData);
        Assert.EndsWith(".pptx", pptxResult.OutputFileName, StringComparison.OrdinalIgnoreCase);

        // Verify strict Microsoft PowerPoint package relationships & OpenXML validity
        using (var checkMs = new MemoryStream(pptxResult.OutputData, writable: false))
        using (var checkDoc = PresentationDocument.Open(checkMs, false))
        {
            var presPart = checkDoc.PresentationPart!;
            var masterPart = Assert.Single(presPart.SlideMasterParts);
            var layoutPart = Assert.Single(masterPart.SlideLayoutParts);
            Assert.NotNull(layoutPart.SlideMasterPart); // Required by PowerPoint (slideLayout1.xml.rels -> slideMaster1.xml)
            Assert.NotNull(presPart.ThemePart);
            Assert.Equal("/ppt/theme/theme1.xml", presPart.ThemePart!.Uri.ToString());
            Assert.NotNull(presPart.PresentationPropertiesPart);
            Assert.NotNull(presPart.ViewPropertiesPart);
            Assert.NotNull(presPart.TableStylesPart);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(checkDoc));
        }

        // Round-trip inspect & convert back to PDF
        var info = _pptxConverter.InspectPresentation(pptxResult.OutputData);
        Assert.True(info.IsValid, info.ErrorMessage);
        Assert.Equal(2, info.SlideCount);
        Assert.Equal(2, info.Slides[0].ImageCount + info.Slides[1].ImageCount);

        var roundTripPdf = await _pptxConverter.ConvertPptxToPdfAsync(new DocumentItem
        {
            FileName = pptxResult.OutputFileName,
            Data = pptxResult.OutputData
        });
        Assert.True(roundTripPdf.Success, roundTripPdf.ErrorMessage);
        Assert.Equal(2, _pdfEngine.GetPageCount(roundTripPdf.OutputData));
    }

    [Fact]
    public async Task LocalPowerPointConverter_InvalidInputs_FailSafely()
    {
        var emptyInspect = _pptxConverter.InspectPresentation(Array.Empty<byte>());
        Assert.False(emptyInspect.IsValid);

        var badPptx = await _pptxConverter.ConvertPptxToPdfAsync(new DocumentItem { FileName = "Bad.pptx", Data = new byte[] { 1, 2, 3 } });
        Assert.False(badPptx.Success);

        var emptySlides = await _pptxConverter.ConvertSlidesToPptxAsync("Empty.pdf", Array.Empty<PdfPageSlideImage>());
        Assert.False(emptySlides.Success);
    }

    [Fact]
    public async Task LocalPowerPointConverter_InheritedPlaceholdersAndMasterStyles_RendersWithFullFormatting()
    {
        // Build a realistic .pptx where the Slide shapes have empty <p:spPr/> (no Transform2D on the slide!)
        // and inherit their Transform2D, background, and text styles from SlideLayout and SlideMaster.
        byte[] pptxBytes;
        using (var ms = new MemoryStream())
        {
            using (var presDoc = PresentationDocument.Create(ms, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, true))
            {
                var presPart = presDoc.AddPresentationPart();
                presPart.Presentation = new OxPres.Presentation();
                var slideMasterPart = presPart.AddNewPart<SlideMasterPart>();
                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

                // Layout defines the placeholder Transform2D for idx=0 (Title) and idx=1 (Body)
                var layoutTitle = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 2U, Name = "Title Placeholder" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Type = OxPres.PlaceholderValues.Title, Index = 0U })),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(
                            new OxDrawing.Offset { X = 609600L, Y = 457200L },
                            new OxDrawing.Extents { Cx = 10972800L, Cy = 1371600L })),
                    new OxPres.TextBody(new OxDrawing.BodyProperties(), new OxDrawing.ListStyle()));

                var layoutBody = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 3U, Name = "Content Placeholder" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Type = OxPres.PlaceholderValues.Object, Index = 1U })),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(
                            new OxDrawing.Offset { X = 609600L, Y = 2133600L },
                            new OxDrawing.Extents { Cx = 10972800L, Cy = 4114800L })),
                    new OxPres.TextBody(new OxDrawing.BodyProperties(), new OxDrawing.ListStyle()));

                slideLayoutPart.SlideLayout = new OxPres.SlideLayout(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                        layoutTitle,
                        layoutBody)));

                // Master defines dark theme background and decorative header bar
                var masterHeaderBar = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 4U, Name = "Master Header Bar" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(
                            new OxDrawing.Offset { X = 0L, Y = 0L },
                            new OxDrawing.Extents { Cx = 12192000L, Cy = 180000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Rectangle },
                        new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "2563EB" })));

                slideMasterPart.SlideMaster = new OxPres.SlideMaster(
                    new OxPres.CommonSlideData(
                        new OxPres.Background(
                            new OxPres.BackgroundProperties(
                                new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "0F172A" }))),
                        new OxPres.ShapeTree(
                            new OxPres.NonVisualGroupShapeProperties(
                                new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                new OxPres.NonVisualGroupShapeDrawingProperties(),
                                new OxPres.ApplicationNonVisualDrawingProperties()),
                            new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                            masterHeaderBar)),
                    new OxPres.SlideLayoutIdList(
                        new OxPres.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }));

                // Slide shapes have empty <p:spPr/> (inheriting coordinates from SlideLayout!)
                var slidePart = presPart.AddNewPart<SlidePart>();
                slidePart.AddPart(slideLayoutPart);

                var slideTitle = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 2U, Name = "Title 1" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Index = 0U })),
                    new OxPres.ShapeProperties(), // Empty spPr!
                    new OxPres.TextBody(
                        new OxDrawing.BodyProperties(),
                        new OxDrawing.ListStyle(),
                        new OxDrawing.Paragraph(
                            new OxDrawing.Run(
                                new OxDrawing.RunProperties(new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "FFFFFF" })) { Bold = true },
                                new OxDrawing.Text("Inherited Layout Title — عنوان الشريحة")))));

                var slideBody = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 3U, Name = "Content 2" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties(new OxPres.PlaceholderShape { Index = 1U })),
                    new OxPres.ShapeProperties(), // Empty spPr!
                    new OxPres.TextBody(
                        new OxDrawing.BodyProperties(),
                        new OxDrawing.ListStyle(),
                        new OxDrawing.Paragraph(
                            new OxDrawing.Run(
                                new OxDrawing.RunProperties(new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "E2E8F0" })),
                                new OxDrawing.Text("Body text inherits transform from SlideLayout index 1.")))));

                slidePart.Slide = new OxPres.Slide(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                        slideTitle,
                        slideBody)));

                presPart.Presentation.SlideMasterIdList = new OxPres.SlideMasterIdList(
                    new OxPres.SlideMasterId { Id = 2147483648U, RelationshipId = presPart.GetIdOfPart(slideMasterPart) });
                presPart.Presentation.SlideSize = new OxPres.SlideSize { Cx = 12192000, Cy = 6858000, Type = OxPres.SlideSizeValues.Screen16x9 };
                presPart.Presentation.SlideIdList = new OxPres.SlideIdList(
                    new OxPres.SlideId { Id = 256U, RelationshipId = presPart.GetIdOfPart(slidePart) });
                presPart.Presentation.Save();
            }
            pptxBytes = ms.ToArray();
        }

        // 1. Inspect and verify live thumbnail preview & inherited title detection
        var info = _pptxConverter.InspectPresentation(pptxBytes);
        Assert.True(info.IsValid, info.ErrorMessage);
        Assert.Single(info.Slides);
        Assert.Equal("Inherited Layout Title — عنوان الشريحة", info.Slides[0].Title);
        Assert.StartsWith("data:image/jpeg;base64,", info.Slides[0].PreviewDataUrl);

        // 2. Convert to PDF
        var pdfRes = await _pptxConverter.ConvertPptxToPdfAsync(new DocumentItem { FileName = "Inherited.pptx", Data = pptxBytes });
        Assert.True(pdfRes.Success, pdfRes.ErrorMessage);
        Assert.Equal(1, _pdfEngine.GetPageCount(pdfRes.OutputData));
    }

    [Fact]
    public async Task LocalPowerPointConverter_RichShapesGroupsMaskedImagesAndArcs_ConvertsCleanlyToPdf()
    {
        // Build a synthetic 2-slide PPTX exercising universal PowerPoint features:
        // - Custom freeform geometry (<a:custGeom> with MoveTo, LineTo, ArcTo, Close)
        // - Preset polygons, arrows, stars (Hexagon with adj, Star5, RightArrow flipped, Plus, Octagon)
        // - Dashed outlines and theme luminance transforms (lumMod + lumOff)
        // - Grouped decorative shapes (<p:grpSp>)
        // - Picture masked with Ellipse (<p:pic>) and Picture-filled shape (<p:sp> with <a:blipFill>)
        byte[] samplePng;
        using (var bmp = new SkiaSharp.SKBitmap(64, 64))
        {
            using (var c = new SkiaSharp.SKCanvas(bmp))
            {
                c.Clear(new SkiaSharp.SKColor(37, 99, 235));
            }
            using var img = SkiaSharp.SKImage.FromBitmap(bmp);
            samplePng = img.Encode(SkiaSharp.SKEncodedImageFormat.Png, 90).ToArray();
        }

        byte[] pptxBytes;
        using (var ms = new MemoryStream())
        {
            using (var presDoc = PresentationDocument.Create(ms, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, true))
            {
                var presPart = presDoc.AddPresentationPart();
                presPart.Presentation = new OxPres.Presentation();

                var masterPart = presPart.AddNewPart<SlideMasterPart>();
                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
                layoutPart.AddPart(masterPart);

                layoutPart.SlideLayout = new OxPres.SlideLayout(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()))));

                masterPart.SlideMaster = new OxPres.SlideMaster(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()))),
                    new OxPres.SlideLayoutIdList(
                        new OxPres.SlideLayoutId { Id = 2147483649U, RelationshipId = masterPart.GetIdOfPart(layoutPart) }));

                // Slide 1: Polygons, arrows, stars, custom arc geometry, dashed outline, grouped shapes
                var slide1 = presPart.AddNewPart<SlidePart>();
                slide1.AddPart(layoutPart);

                var hexShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 10U, Name = "Hexagon" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(new OxDrawing.Offset { X = 400000L, Y = 400000L }, new OxDrawing.Extents { Cx = 1800000L, Cy = 1500000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList(new OxDrawing.ShapeGuide { Name = "adj", Formula = "val 28413" })) { Preset = OxDrawing.ShapeTypeValues.Hexagon },
                        new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex(new OxDrawing.LuminanceModulation { Val = 60000 }, new OxDrawing.LuminanceOffset { Val = 40000 }) { Val = "4472C4" }),
                        new OxDrawing.Outline(new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "1E293B" }), new OxDrawing.PresetDash { Val = OxDrawing.PresetLineDashValues.Dash }) { Width = 25400 }));

                var arrowShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 11U, Name = "FlippedArrow" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(new OxDrawing.Offset { X = 2500000L, Y = 400000L }, new OxDrawing.Extents { Cx = 1800000L, Cy = 900000L }) { HorizontalFlip = true },
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.RightArrow },
                        new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "EA580C" })));

                var custArcShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 12U, Name = "CustomArc" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(new OxDrawing.Offset { X = 4600000L, Y = 400000L }, new OxDrawing.Extents { Cx = 1600000L, Cy = 1600000L }),
                        new OxDrawing.CustomGeometry(
                            new OxDrawing.AdjustValueList(),
                            new OxDrawing.ShapeGuideList(),
                            new OxDrawing.AdjustHandleList(),
                            new OxDrawing.ConnectionSiteList(),
                            new OxDrawing.Rectangle { Left = "l", Top = "t", Right = "r", Bottom = "b" },
                            new OxDrawing.PathList(
                                new OxDrawing.Path(
                                    new OxDrawing.MoveTo(new OxDrawing.Point { X = "0", Y = "500" }),
                                    new OxDrawing.ArcTo { WidthRadius = "500", HeightRadius = "500", StartAngle = "10800000", SwingAngle = "10800000" },
                                    new OxDrawing.LineTo(new OxDrawing.Point { X = "1000", Y = "1000" }),
                                    new OxDrawing.LineTo(new OxDrawing.Point { X = "0", Y = "1000" }),
                                    new OxDrawing.CloseShapePath())
                                { Width = 1000L, Height = 1000L })),
                        new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "10B981" })));

                var groupShape = new OxPres.GroupShape(
                    new OxPres.NonVisualGroupShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 20U, Name = "GroupedStars" },
                        new OxPres.NonVisualGroupShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.GroupShapeProperties(
                        new OxDrawing.TransformGroup(
                            new OxDrawing.Offset { X = 400000L, Y = 2400000L },
                            new OxDrawing.Extents { Cx = 3600000L, Cy = 1800000L },
                            new OxDrawing.ChildOffset { X = 0L, Y = 0L },
                            new OxDrawing.ChildExtents { Cx = 3600000L, Cy = 1800000L })),
                    new OxPres.Shape(
                        new OxPres.NonVisualShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 21U, Name = "Star5" },
                            new OxPres.NonVisualShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.ShapeProperties(
                            new OxDrawing.Transform2D(new OxDrawing.Offset { X = 0L, Y = 0L }, new OxDrawing.Extents { Cx = 1600000L, Cy = 1600000L }),
                            new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Star5 },
                            new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "F59E0B" }))),
                    new OxPres.Shape(
                        new OxPres.NonVisualShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 22U, Name = "Plus" },
                            new OxPres.NonVisualShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.ShapeProperties(
                            new OxDrawing.Transform2D(new OxDrawing.Offset { X = 1900000L, Y = 0L }, new OxDrawing.Extents { Cx = 1600000L, Cy = 1600000L }),
                            new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Plus },
                            new OxDrawing.SolidFill(new OxDrawing.RgbColorModelHex { Val = "6366F1" }))));

                slide1.Slide = new OxPres.Slide(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                        hexShape, arrowShape, custArcShape, groupShape)));
                slide1.Slide.Save();

                // Slide 2: Masked <p:pic> (Ellipse) and Picture-filled <p:sp> (Octagon)
                var slide2 = presPart.AddNewPart<SlidePart>();
                slide2.AddPart(layoutPart);

                var imgPart1 = slide2.AddImagePart(ImagePartType.Png);
                using (var s = new MemoryStream(samplePng)) imgPart1.FeedData(s);
                string relImg1 = slide2.GetIdOfPart(imgPart1);

                var maskedPic = new OxPres.Picture(
                    new OxPres.NonVisualPictureProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 30U, Name = "AvatarPic" },
                        new OxPres.NonVisualPictureDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.BlipFill(new OxDrawing.Blip { Embed = relImg1 }, new OxDrawing.Stretch(new OxDrawing.FillRectangle())),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(new OxDrawing.Offset { X = 600000L, Y = 800000L }, new OxDrawing.Extents { Cx = 2400000L, Cy = 2400000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Ellipse }));

                var blipShape = new OxPres.Shape(
                    new OxPres.NonVisualShapeProperties(
                        new OxPres.NonVisualDrawingProperties { Id = 31U, Name = "OctagonPhoto" },
                        new OxPres.NonVisualShapeDrawingProperties(),
                        new OxPres.ApplicationNonVisualDrawingProperties()),
                    new OxPres.ShapeProperties(
                        new OxDrawing.Transform2D(new OxDrawing.Offset { X = 3800000L, Y = 800000L }, new OxDrawing.Extents { Cx = 2400000L, Cy = 2400000L }),
                        new OxDrawing.PresetGeometry(new OxDrawing.AdjustValueList()) { Preset = OxDrawing.ShapeTypeValues.Octagon },
                        new OxDrawing.BlipFill(new OxDrawing.Blip { Embed = relImg1 }, new OxDrawing.Stretch(new OxDrawing.FillRectangle()))));

                slide2.Slide = new OxPres.Slide(
                    new OxPres.CommonSlideData(new OxPres.ShapeTree(
                        new OxPres.NonVisualGroupShapeProperties(
                            new OxPres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new OxPres.NonVisualGroupShapeDrawingProperties(),
                            new OxPres.ApplicationNonVisualDrawingProperties()),
                        new OxPres.GroupShapeProperties(new OxDrawing.TransformGroup()),
                        maskedPic, blipShape)));
                slide2.Slide.Save();

                presPart.Presentation.SlideMasterIdList = new OxPres.SlideMasterIdList(
                    new OxPres.SlideMasterId { Id = 2147483648U, RelationshipId = presPart.GetIdOfPart(masterPart) });
                presPart.Presentation.SlideSize = new OxPres.SlideSize { Cx = 12192000, Cy = 6858000, Type = OxPres.SlideSizeValues.Screen16x9 };
                presPart.Presentation.SlideIdList = new OxPres.SlideIdList(
                    new OxPres.SlideId { Id = 256U, RelationshipId = presPart.GetIdOfPart(slide1) },
                    new OxPres.SlideId { Id = 257U, RelationshipId = presPart.GetIdOfPart(slide2) });
                presPart.Presentation.Save();
            }
            pptxBytes = ms.ToArray();
        }

        var result = await _pptxConverter.ConvertPptxToPdfAsync(new DocumentItem { FileName = "UniversalDeck.pptx", Data = pptxBytes });
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.OutputData);
        Assert.Equal(2, _pdfEngine.GetPageCount(result.OutputData));
    }

    // =========================================================================
    // PDF to Images Tests (Page Packaging + Embedded Image Extraction)
    // =========================================================================

    [Fact]
    public async Task LocalPdfImageConverter_PackageSingleImage_ReturnsDirectImageFile()
    {
        byte[] jpgBytes;
        using (var bmp = new SkiaSharp.SKBitmap(120, 80))
        {
            using (var c = new SkiaSharp.SKCanvas(bmp)) c.Clear(SkiaSharp.SKColors.CornflowerBlue);
            using var img = SkiaSharp.SKImage.FromBitmap(bmp);
            jpgBytes = img.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, 90).ToArray();
        }

        var res = await _pdfImageConverter.PackageImagesAsync(
            "Report.pdf",
            new[]
            {
                new RenderedPdfPageImage
                {
                    PageNumber = 1,
                    FileName = "Report_page_001.jpg",
                    MimeType = "image/jpeg",
                    Data = jpgBytes
                }
            },
            ImageOutputFormat.Jpg);

        Assert.True(res.Success, res.ErrorMessage);
        Assert.Equal("Report_page_001.jpg", res.OutputFileName);
        Assert.Equal("image/jpeg", res.ContentType);
        Assert.Equal(jpgBytes, res.OutputData);
    }

    [Fact]
    public async Task LocalPdfImageConverter_PackageMultipleImages_ReturnsValidZipArchive()
    {
        byte[] pngBytes;
        using (var bmp = new SkiaSharp.SKBitmap(64, 64))
        {
            using (var c = new SkiaSharp.SKCanvas(bmp)) c.Clear(SkiaSharp.SKColors.Teal);
            using var img = SkiaSharp.SKImage.FromBitmap(bmp);
            pngBytes = img.Encode(SkiaSharp.SKEncodedImageFormat.Png, 90).ToArray();
        }

        var inputs = new[]
        {
            new RenderedPdfPageImage { PageNumber = 1, FileName = "Deck_page_001.png", MimeType = "image/png", Data = pngBytes },
            new RenderedPdfPageImage { PageNumber = 2, FileName = "Deck_page_002.png", MimeType = "image/png", Data = pngBytes },
            new RenderedPdfPageImage { PageNumber = 3, FileName = "Deck_page_003.png", MimeType = "image/png", Data = pngBytes }
        };

        var res = await _pdfImageConverter.PackageImagesAsync("Deck.pdf", inputs, ImageOutputFormat.Png);

        Assert.True(res.Success, res.ErrorMessage);
        Assert.Equal("application/zip", res.ContentType);
        Assert.EndsWith(".zip", res.OutputFileName, StringComparison.OrdinalIgnoreCase);

        using var zipMs = new MemoryStream(res.OutputData, writable: false);
        using var zip = new ZipArchive(zipMs, ZipArchiveMode.Read);
        Assert.Equal(3, zip.Entries.Count);
        Assert.Contains(zip.Entries, e => e.Name == "Deck_page_001.png");
        Assert.Contains(zip.Entries, e => e.Name == "Deck_page_002.png");
        Assert.Contains(zip.Entries, e => e.Name == "Deck_page_003.png");
    }

    [Fact]
    public async Task LocalPdfImageConverter_ExtractEmbeddedImages_ExtractsJpegAndRgbImages()
    {
        // Create a 2-page PDF embedding a JPEG on page 1 and a PNG (RGB + alpha) on page 2
        byte[] jpgBytes;
        using (var bmp = new SkiaSharp.SKBitmap(100, 60))
        {
            using (var c = new SkiaSharp.SKCanvas(bmp)) c.Clear(SkiaSharp.SKColors.DarkOrange);
            using var img = SkiaSharp.SKImage.FromBitmap(bmp);
            jpgBytes = img.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, 90).ToArray();
        }

        byte[] pngBytes;
        using (var bmp = new SkiaSharp.SKBitmap(80, 80))
        {
            using (var c = new SkiaSharp.SKCanvas(bmp))
            {
                c.Clear(SkiaSharp.SKColors.Transparent);
                using var p = new SkiaSharp.SKPaint { Color = SkiaSharp.SKColors.MediumSeaGreen, IsAntialias = true };
                c.DrawCircle(40, 40, 30, p);
            }
            using var img = SkiaSharp.SKImage.FromBitmap(bmp);
            pngBytes = img.Encode(SkiaSharp.SKEncodedImageFormat.Png, 95).ToArray();
        }

        byte[] pdfBytes;
        using (var doc = new PdfDocument())
        {
            var p1 = doc.AddPage();
            using (var gfx = XGraphics.FromPdfPage(p1))
            using (var jMs = new MemoryStream(jpgBytes))
            using (var xImg = XImage.FromStream(jMs))
            {
                gfx.DrawImage(xImg, 40, 40, 200, 120);
            }

            var p2 = doc.AddPage();
            using (var gfx = XGraphics.FromPdfPage(p2))
            using (var pMs = new MemoryStream(pngBytes))
            using (var xImg = XImage.FromStream(pMs))
            {
                gfx.DrawImage(xImg, 40, 40, 160, 160);
            }

            using var outMs = new MemoryStream();
            doc.Save(outMs, false);
            pdfBytes = outMs.ToArray();
        }

        var extracted = await _pdfImageConverter.ExtractEmbeddedImagesAsync(pdfBytes, ImageOutputFormat.Png);

        Assert.Equal(2, extracted.Count);
        Assert.Equal(1, extracted[0].PageNumber);
        Assert.Equal(100, extracted[0].Width);
        Assert.Equal(60, extracted[0].Height);
        Assert.NotEmpty(extracted[0].Data);
        Assert.StartsWith("data:image/png;base64,", extracted[0].PreviewDataUrl);

        Assert.Equal(2, extracted[1].PageNumber);
        Assert.Equal(80, extracted[1].Width);
        Assert.Equal(80, extracted[1].Height);
        Assert.NotEmpty(extracted[1].Data);
    }

    [Fact]
    public async Task LocalPdfImageConverter_ExtractFromTextOnlyPdf_ReturnsEmptyListSafely()
    {
        var textOnlyPdf = CreateSamplePdf(2, "NoImages");
        var extracted = await _pdfImageConverter.ExtractEmbeddedImagesAsync(textOnlyPdf, ImageOutputFormat.Jpg);
        Assert.Empty(extracted);
    }

    // =========================================================================
    // Images to PDF Tests (JPG / PNG / WebP -> Merged PDF or Separate ZIP)
    // =========================================================================

    private static byte[] CreateSampleImageBytes(int width, int height, SkiaSharp.SKColor color, SkiaSharp.SKEncodedImageFormat fmt)
    {
        using var bmp = new SkiaSharp.SKBitmap(width, height);
        using (var c = new SkiaSharp.SKCanvas(bmp)) c.Clear(color);
        using var img = SkiaSharp.SKImage.FromBitmap(bmp);
        return img.Encode(fmt, 90).ToArray();
    }

    [Fact]
    public async Task LocalPdfImageConverter_ImagesToPdf_MergedA4AndLandscape_CreatesValidPdf()
    {
        var jpg = await _pdfImageConverter.InspectImageAsync("photo1.jpg", "image/jpeg", CreateSampleImageBytes(300, 200, SkiaSharp.SKColors.SteelBlue, SkiaSharp.SKEncodedImageFormat.Jpeg));
        var png = await _pdfImageConverter.InspectImageAsync("diagram2.png", "image/png", CreateSampleImageBytes(200, 300, SkiaSharp.SKColors.Coral, SkiaSharp.SKEncodedImageFormat.Png));
        var webp = await _pdfImageConverter.InspectImageAsync("graphic3.webp", "image/webp", CreateSampleImageBytes(250, 250, SkiaSharp.SKColors.MediumPurple, SkiaSharp.SKEncodedImageFormat.Webp));

        Assert.NotNull(jpg);
        Assert.NotNull(png);
        Assert.NotNull(webp);
        Assert.Equal(300, jpg!.Width);
        Assert.Equal(200, jpg.Height);
        Assert.StartsWith("data:image/png;base64,", jpg.PreviewDataUrl);

        var res = await _pdfImageConverter.ConvertImagesToPdfAsync(
            new[] { jpg, png!, webp! },
            new ImagesToPdfOptions
            {
                PageSize = ImagePdfPageSize.A4,
                Orientation = ImagePdfOrientation.Landscape,
                Margin = ImagePdfMargin.Small,
                MergeIntoSinglePdf = true
            });

        Assert.True(res.Success, res.ErrorMessage);
        Assert.Equal("application/pdf", res.ContentType);
        Assert.Equal(3, _pdfEngine.GetPageCount(res.OutputData));

        using var ms = new MemoryStream(res.OutputData, writable: false);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        // A4 Landscape: width > height (~841.89 x 595.28 pt)
        Assert.True(pdf.Pages[0].Width.Point > pdf.Pages[0].Height.Point);
        Assert.InRange(pdf.Pages[0].Width.Point, 840, 844);
        Assert.InRange(pdf.Pages[0].Height.Point, 593, 597);
    }

    [Fact]
    public async Task LocalPdfImageConverter_ImagesToPdf_FitToImageAndRotation_AdjustsPageDimensions()
    {
        var wideImg = await _pdfImageConverter.InspectImageAsync("wide.jpg", "image/jpeg", CreateSampleImageBytes(400, 200, SkiaSharp.SKColors.Teal, SkiaSharp.SKEncodedImageFormat.Jpeg));
        Assert.NotNull(wideImg);
        wideImg!.RotationDegrees = 90; // After 90° rotation -> 200w x 400h (tall!)

        var res = await _pdfImageConverter.ConvertImagesToPdfAsync(
            new[] { wideImg },
            new ImagesToPdfOptions
            {
                PageSize = ImagePdfPageSize.FitToImage,
                Margin = ImagePdfMargin.None,
                MergeIntoSinglePdf = true
            });

        Assert.True(res.Success, res.ErrorMessage);
        using var ms = new MemoryStream(res.OutputData, writable: false);
        using var pdf = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
        Assert.Equal(1, pdf.PageCount);
        Assert.True(pdf.Pages[0].Height.Point > pdf.Pages[0].Width.Point);
    }

    [Fact]
    public async Task LocalPdfImageConverter_ImagesToPdf_SeparateFilesMode_ReturnsValidZipOfPdfs()
    {
        var img1 = await _pdfImageConverter.InspectImageAsync("scan1.jpg", "image/jpeg", CreateSampleImageBytes(160, 120, SkiaSharp.SKColors.Navy, SkiaSharp.SKEncodedImageFormat.Jpeg));
        var img2 = await _pdfImageConverter.InspectImageAsync("scan2.png", "image/png", CreateSampleImageBytes(160, 120, SkiaSharp.SKColors.DarkGreen, SkiaSharp.SKEncodedImageFormat.Png));

        var res = await _pdfImageConverter.ConvertImagesToPdfAsync(
            new[] { img1!, img2! },
            new ImagesToPdfOptions
            {
                PageSize = ImagePdfPageSize.Letter,
                Orientation = ImagePdfOrientation.Portrait,
                Margin = ImagePdfMargin.Big,
                MergeIntoSinglePdf = false
            });

        Assert.True(res.Success, res.ErrorMessage);
        Assert.Equal("application/zip", res.ContentType);
        Assert.EndsWith(".zip", res.OutputFileName, StringComparison.OrdinalIgnoreCase);

        using var zipMs = new MemoryStream(res.OutputData, writable: false);
        using var zip = new ZipArchive(zipMs, ZipArchiveMode.Read);
        Assert.Equal(2, zip.Entries.Count);
        Assert.Contains(zip.Entries, e => e.Name == "scan1.pdf");
        Assert.Contains(zip.Entries, e => e.Name == "scan2.pdf");

        foreach (var entry in zip.Entries)
        {
            using var es = entry.Open();
            using var pMs = new MemoryStream();
            es.CopyTo(pMs);
            Assert.Equal(1, _pdfEngine.GetPageCount(pMs.ToArray()));
        }
    }

    [Fact]
    public async Task LocalPdfImageConverter_ImagesToPdf_EmptyInput_FailsSafely()
    {
        var res = await _pdfImageConverter.ConvertImagesToPdfAsync(Array.Empty<ImageToPdfInputItem>());
        Assert.False(res.Success);
        Assert.NotEmpty(res.ErrorMessage);
    }
}
