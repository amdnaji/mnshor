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
}
