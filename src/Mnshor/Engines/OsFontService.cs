using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Mnshor.Models;
using SkiaSharp;
using SkiaSharp.HarfBuzz;

namespace Mnshor.Engines;

public class OsFontService
{
    private static readonly HashSet<string> ExcludedSymbolFamilies = new(StringComparer.OrdinalIgnoreCase)
    {
        "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Marlett",
        "Symbol", "Segoe MDL2 Assets", "Segoe Fluent Icons", "HoloLens MDL2 Assets",
        "MT Extra", "MS Outlook", "MS Reference Specialty", "Bookshelf Symbol 7",
        "Apple Symbols", "Zapf Dingbats"
    };

    // Known cross-platform candidates for Latin Signature / Handwriting / Script
    private static readonly (string FamilyName, string DisplayName, string[] FileHints)[] SignatureCandidates =
    [
        ("Ink Free", "Ink Free (Pen Script)", ["Inkfree.ttf"]),
        ("Segoe Script", "Segoe Script", ["segoesc.ttf", "segoescb.ttf"]),
        ("Brush Script MT", "Brush Script MT", ["BRUSHSCI.TTF", "Brush Script.ttf"]),
        ("Lucida Handwriting", "Lucida Handwriting", ["LHANDW.TTF"]),
        ("Edwardian Script ITC", "Edwardian Script", ["ITCEDSCR.TTF"]),
        ("Kunstler Script", "Kunstler Script", ["KUNSTLER.TTF"]),
        ("Vladimir Script", "Vladimir Script", ["VLADIMIR.TTF"]),
        ("Freestyle Script", "Freestyle Script", ["FREESCPT.TTF"]),
        ("Pristina", "Pristina", ["PRISTINA.TTF"]),
        ("Mistral", "Mistral", ["MISTRAL.TTF"]),
        ("Monotype Corsiva", "Monotype Corsiva", ["MTCORSVA.TTF"]),
        ("Bradley Hand ITC", "Bradley Hand", ["BRADHITC.TTF", "Bradley Hand Bold.ttf"]),
        ("Segoe Print", "Segoe Print", ["segoepr.ttf", "segoeprb.ttf"]),
        ("French Script MT", "French Script", ["FRSCRIPT.TTF"]),
        ("Rage Italic", "Rage Italic", ["RAGE.TTF"]),
        ("Viner Hand ITC", "Viner Hand", ["VINERITC.TTF"]),
        ("Gabriola", "Gabriola", ["Gabriola.ttf"]),
        // macOS signature script fonts
        ("Snell Roundhand", "Snell Roundhand", ["SnellRoundhand.ttc"]),
        ("Apple Chancery", "Apple Chancery", ["Apple Chancery.ttf"]),
        ("Zapfino", "Zapfino", ["Zapfino.ttf"]),
        ("Noteworthy", "Noteworthy", ["Noteworthy.ttc"]),
        ("Savoye LET", "Savoye Script", []),
        ("SignPainter", "SignPainter", []),
        // Linux script fonts
        ("URW Chancery L", "URW Chancery", []),
        ("Z003", "Z003 Chancery", []),
        ("Purisa", "Purisa Handwritten", []),
        ("Dancing Script", "Dancing Script", []),
        ("Great Vibes", "Great Vibes", []),
        ("Pacifico", "Pacifico", [])
    ];

    // Known cross-platform candidates for Arabic Calligraphy & Executive Fonts
    private static readonly (string FamilyName, string DisplayName, string[] FileHints)[] ArabicCandidates =
    [
        ("Aldhabi", "Aldhabi — الديواني الملكي", ["aldhabi.ttf"]),
        ("Arabic Typesetting", "Arabic Typesetting — الخط العربي الأصيل", ["arabtype.ttf"]),
        ("Urdu Typesetting", "Urdu Typesetting — خط النستعليق", ["UrdType.ttf", "UrdTypeb.ttf"]),
        ("Sakkal Majalla", "Sakkal Majalla — سقال مجلة", ["majalla.ttf", "majallab.ttf"]),
        ("Dubai", "Dubai — خط دبي التنفيذي", ["DUBAI-REGULAR.TTF", "DUBAI-MEDIUM.TTF", "DUBAI-BOLD.TTF"]),
        ("Andalus", "Andalus — الخط الأندلسي", ["andlso.ttf"]),
        ("Traditional Arabic", "Traditional Arabic — الخط التقليدي", ["trado.ttf", "tradbdo.ttf"]),
        ("Simplified Arabic", "Simplified Arabic — الخط المبسط", ["simpo.ttf", "simpbdo.ttf"]),
        // macOS Arabic fonts
        ("Diwan Thuluth", "Diwan Thuluth — خط الثلث", []),
        ("Diwan Kufi", "Diwan Kufi — الخط الكوفي", []),
        ("Geeza Pro", "Geeza Pro — جيزة برو", []),
        ("Damascus", "Damascus — خط دمشق", []),
        ("Al Nile", "Al Nile — خط النيل", []),
        ("Baghdad", "Baghdad — خط بغداد", []),
        ("Farah", "Farah — خط فرح", []),
        // Linux Arabic fonts
        ("Amiri", "Amiri — الخط الأميري", []),
        ("Scheherazade New", "Scheherazade — شهرزاد", []),
        ("Noto Naskh Arabic", "Noto Naskh — نسخ نوتو", []),
        ("Cairo", "Cairo — خط القاهرة", []),
        ("Tajawal", "Tajawal — خط تجول", []),
        // Universal Arabic-supporting system fonts
        ("Segoe UI", "Segoe UI — خط النظام القياسي", ["segoeui.ttf", "segoeuib.ttf"]),
        ("Tahoma", "Tahoma — تاهوما", ["tahoma.ttf", "tahomabd.ttf"]),
        ("Arial", "Arial — أريال", ["arial.ttf", "arialbd.ttf"])
    ];

    private readonly Lazy<List<string>> _installedFamilies;
    private readonly Lazy<List<OsFontDescriptor>> _curatedSignatureFonts;
    private readonly Lazy<List<OsFontDescriptor>> _curatedArabicFonts;
    private readonly Lazy<List<string>> _osFontDirectories;
    private readonly ConcurrentDictionary<string, byte[]> _previewCache = new(StringComparer.Ordinal);

    public OsFontService()
    {
        _osFontDirectories = new Lazy<List<string>>(DiscoverOsFontDirectories);
        _installedFamilies = new Lazy<List<string>>(DiscoverInstalledFontFamilies);
        _curatedSignatureFonts = new Lazy<List<OsFontDescriptor>>(BuildCuratedSignatureFonts);
        _curatedArabicFonts = new Lazy<List<OsFontDescriptor>>(BuildCuratedArabicFonts);
    }

    /// <summary>
    /// Returns all installed OS font families across Windows, macOS, and Linux, sorted alphabetically.
    /// </summary>
    public IReadOnlyList<string> GetInstalledFontFamilies() => _installedFamilies.Value;

    /// <summary>
    /// Returns curated signature & handwriting fonts available on the current OS.
    /// </summary>
    public IReadOnlyList<OsFontDescriptor> GetSignatureFonts() => _curatedSignatureFonts.Value;

    /// <summary>
    /// Returns curated Arabic calligraphy & executive fonts available on the current OS.
    /// </summary>
    public IReadOnlyList<OsFontDescriptor> GetArabicFonts() => _curatedArabicFonts.Value;

    private static List<string> DiscoverOsFontDirectories()
    {
        var dirs = new List<string>();

        try
        {
            string specialFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            if (!string.IsNullOrWhiteSpace(specialFonts) && Directory.Exists(specialFonts))
            {
                dirs.Add(specialFonts);
            }
        }
        catch
        {
            // Ignore restricted path lookups
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrWhiteSpace(winDir))
            {
                string sysFonts = Path.Combine(winDir, "Fonts");
                if (Directory.Exists(sysFonts) && !dirs.Contains(sysFonts, StringComparer.OrdinalIgnoreCase))
                {
                    dirs.Add(sysFonts);
                }
            }

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrWhiteSpace(localAppData))
            {
                string userFonts = Path.Combine(localAppData, "Microsoft", "Windows", "Fonts");
                if (Directory.Exists(userFonts) && !dirs.Contains(userFonts, StringComparer.OrdinalIgnoreCase))
                {
                    dirs.Add(userFonts);
                }
            }
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string[] macDirs =
            [
                "/System/Library/Fonts",
                "/Library/Fonts",
                !string.IsNullOrWhiteSpace(home) ? Path.Combine(home, "Library", "Fonts") : string.Empty
            ];

            foreach (var d in macDirs)
            {
                if (!string.IsNullOrWhiteSpace(d) && Directory.Exists(d))
                {
                    dirs.Add(d);
                }
            }
        }
        else
        {
            // Linux / Unix
            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string[] linuxDirs =
            [
                "/usr/share/fonts",
                "/usr/local/share/fonts",
                !string.IsNullOrWhiteSpace(home) ? Path.Combine(home, ".local", "share", "fonts") : string.Empty,
                !string.IsNullOrWhiteSpace(home) ? Path.Combine(home, ".fonts") : string.Empty
            ];

            foreach (var d in linuxDirs)
            {
                if (!string.IsNullOrWhiteSpace(d) && Directory.Exists(d))
                {
                    dirs.Add(d);
                }
            }
        }

        return dirs;
    }

    private List<string> DiscoverInstalledFontFamilies()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var families = SKFontManager.Default.GetFontFamilies();
            foreach (var fam in families)
            {
                if (string.IsNullOrWhiteSpace(fam) || fam.StartsWith('@'))
                    continue;
                if (ExcludedSymbolFamilies.Contains(fam))
                    continue;
                set.Add(fam.Trim());
            }
        }
        catch
        {
            // Fallback if SKFontManager enumeration fails
        }

        // Also ensure any known candidate whose font file exists in OS font dirs is registered
        foreach (var (fam, _, hints) in SignatureCandidates.Concat(ArabicCandidates))
        {
            if (!set.Contains(fam) && TryFindFontFileInOsDirs(hints) != null)
            {
                set.Add(fam);
            }
        }

        if (set.Count == 0)
        {
            set.Add("Segoe UI");
            set.Add("Arial");
            set.Add("Times New Roman");
        }

        return set.OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private List<OsFontDescriptor> BuildCuratedSignatureFonts()
    {
        var installed = new HashSet<string>(_installedFamilies.Value, StringComparer.OrdinalIgnoreCase);
        var list = new List<OsFontDescriptor>();

        foreach (var (family, display, hints) in SignatureCandidates)
        {
            if (installed.Contains(family) || TryFindFontFileInOsDirs(hints) != null)
            {
                list.Add(new OsFontDescriptor
                {
                    FamilyName = family,
                    DisplayName = display,
                    Category = OsFontCategory.SignatureScript
                });
            }
        }

        // Ensure at least a few standard fonts if running on a minimal container OS
        if (list.Count < 4)
        {
            foreach (var fallback in new[] { "Georgia", "Palatino Linotype", "Times New Roman", "Calibri", "Segoe UI", "Arial", "DejaVu Serif", "Liberation Serif" })
            {
                if (installed.Contains(fallback) && list.All(x => !x.FamilyName.Equals(fallback, StringComparison.OrdinalIgnoreCase)))
                {
                    list.Add(new OsFontDescriptor
                    {
                        FamilyName = fallback,
                        DisplayName = fallback,
                        Category = OsFontCategory.SignatureScript
                    });
                }
            }
        }

        return list;
    }

    private List<OsFontDescriptor> BuildCuratedArabicFonts()
    {
        var installed = new HashSet<string>(_installedFamilies.Value, StringComparer.OrdinalIgnoreCase);
        var list = new List<OsFontDescriptor>();

        foreach (var (family, display, hints) in ArabicCandidates)
        {
            if (installed.Contains(family) || TryFindFontFileInOsDirs(hints) != null)
            {
                list.Add(new OsFontDescriptor
                {
                    FamilyName = family,
                    DisplayName = display,
                    Category = OsFontCategory.ArabicCalligraphy
                });
            }
        }

        if (list.Count == 0)
        {
            list.Add(new OsFontDescriptor
            {
                FamilyName = _installedFamilies.Value.FirstOrDefault() ?? "Arial",
                DisplayName = "System Default",
                Category = OsFontCategory.ArabicCalligraphy
            });
        }

        return list;
    }

    private string? TryFindFontFileInOsDirs(string[] fileHints)
    {
        if (fileHints == null || fileHints.Length == 0)
            return null;

        foreach (var dir in _osFontDirectories.Value)
        {
            foreach (var file in fileHints)
            {
                string candidate = Path.Combine(dir, file);
                if (File.Exists(candidate))
                    return candidate;
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves an SKTypeface for the requested OS font family, automatically falling back to an Arabic-capable
    /// OS font if the text contains Arabic script and the chosen font lacks Arabic glyphs.
    /// </summary>
    public SKTypeface ResolveOsTypeface(string? fontFamily, bool isBold = false, bool isItalic = false, bool requiresArabicGlyphs = false)
    {
        string requested = string.IsNullOrWhiteSpace(fontFamily) ? "Segoe Script" : fontFamily.Trim();

        var style = new SKFontStyle(
            isBold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            isItalic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);

        // 1. Check direct file hints for curated fonts first (ensures exact typeface loading)
        var matchedCandidate = SignatureCandidates.Concat(ArabicCandidates)
            .FirstOrDefault(c => c.FamilyName.Equals(requested, StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrEmpty(matchedCandidate.FamilyName) && matchedCandidate.FileHints.Length > 0)
        {
            string? filePath = TryFindFontFileInOsDirs(matchedCandidate.FileHints);
            if (filePath != null)
            {
                var fromFile = SKTypeface.FromFile(filePath);
                if (fromFile != null)
                {
                    if (!requiresArabicGlyphs || SupportsArabicGlyphs(fromFile))
                        return fromFile;
                    fromFile.Dispose();
                }
            }
        }

        // 2. Query OS Font Manager (works across Windows, macOS, and Linux)
        var matched = SKFontManager.Default.MatchFamily(requested, style)
                      ?? SKTypeface.FromFamilyName(requested, style);

        if (matched != null)
        {
            if (!requiresArabicGlyphs || SupportsArabicGlyphs(matched))
                return matched;
        }

        // 3. If Arabic glyphs are required and the chosen Latin script font doesn't have Arabic glyphs,
        // pick the best installed Arabic calligraphic or system font on this OS!
        if (requiresArabicGlyphs)
        {
            foreach (var arCandidate in _curatedArabicFonts.Value)
            {
                var arMatch = ArabicCandidates.FirstOrDefault(c => c.FamilyName.Equals(arCandidate.FamilyName, StringComparison.OrdinalIgnoreCase));
                string? arPath = TryFindFontFileInOsDirs(arMatch.FileHints);
                if (arPath != null)
                {
                    var tf = SKTypeface.FromFile(arPath);
                    if (tf != null && SupportsArabicGlyphs(tf))
                        return tf;
                }

                var sysTf = SKFontManager.Default.MatchFamily(arCandidate.FamilyName, style);
                if (sysTf != null && SupportsArabicGlyphs(sysTf))
                    return sysTf;
            }

            // Use SKFontManager character matching for Arabic 'م'
            var charMatched = SKFontManager.Default.MatchCharacter('\u0645');
            if (charMatched != null)
                return charMatched;
        }

        return matched ?? SKTypeface.Default;
    }

    private static bool SupportsArabicGlyphs(SKTypeface typeface)
    {
        try
        {
            using var font = new SKFont(typeface, 24f);
            ushort glyph = font.GetGlyph('\u0645'); // Arabic letter Meem 'م'
            return glyph != 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Renders a typed signature, initials, date, or custom text into a high-DPI transparent PNG in memory
    /// using the selected OS font, HarfBuzz shaping, and EmojiRenderer.
    /// </summary>
    public (byte[] PngBytes, double AspectRatio) RenderTextSignatureToPng(
        string text,
        string fontFamily,
        string colorHex = "#0F172A",
        float fontSizePt = 44f,
        bool isBold = false,
        bool isItalic = false,
        bool includeSignatureLine = false)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            text = "Signature";
        }

        string trimmed = text.Trim();
        bool hasArabic = ContainsArabic(trimmed);

        string cacheKey = $"{trimmed}|{fontFamily}|{colorHex}|{fontSizePt:F1}|{isBold}|{isItalic}|{includeSignatureLine}";
        if (_previewCache.TryGetValue(cacheKey, out var cachedPng))
        {
            using var cachedBmp = SKBitmap.Decode(cachedPng);
            double cachedRatio = cachedBmp != null && cachedBmp.Height > 0
                ? (double)cachedBmp.Width / cachedBmp.Height
                : 3.0;
            return (cachedPng, cachedRatio);
        }

        const float renderScale = 2.6f;
        float scaledFontSize = Math.Clamp(fontSizePt, 12f, 120f) * renderScale;

        using var typeface = ResolveOsTypeface(fontFamily, isBold, isItalic, requiresArabicGlyphs: hasArabic);
        using var font = new SKFont(typeface, scaledFontSize)
        {
            Edging = SKFontEdging.SubpixelAntialias,
            Subpixel = true,
            Embolden = isBold && !typeface.IsBold,
            SkewX = (isItalic && !typeface.IsItalic) ? -0.18f : 0f
        };
        using var shaper = new SKShaper(typeface);

        if (!SKColor.TryParse(colorHex, out var inkColor))
        {
            inkColor = new SKColor(15, 23, 42);
        }

        using var paint = new SKPaint
        {
            IsAntialias = true,
            Color = inkColor
        };

        var words = trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        float spaceW = font.MeasureText(" ");
        float emojiAdvance = scaledFontSize * 1.15f;

        var measuredWords = new List<(string Word, float Width)>();
        float totalTextWidth = 0f;

        foreach (var word in words)
        {
            var tokens = EmojiRenderer.TokenizeByEmoji(word);
            float w = 0f;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    w += emojiAdvance;
                }
                else
                {
                    var shaped = shaper.Shape(tokenText, font);
                    w += shaped.Width;
                }
            }
            measuredWords.Add((word, w));
            totalTextWidth += (measuredWords.Count == 1 ? w : spaceW + w);
        }

        // Script/calligraphic fonts often have generous swashes/descenders; give comfortable padding
        float padX = scaledFontSize * 0.42f;
        float padY = scaledFontSize * 0.48f;
        int bmpW = Math.Max(48, (int)Math.Ceiling(totalTextWidth + (padX * 2f)));
        int bmpH = Math.Max(36, (int)Math.Ceiling((scaledFontSize * 1.48f) + (padY * 2f)));

        var info = new SKImageInfo(bmpW, bmpH, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var surface = SKSurface.Create(info);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.Transparent);

        float baselineY = padY + (scaledFontSize * 0.88f);
        float startX = (bmpW - totalTextWidth) / 2f;

        if (hasArabic)
        {
            float cursorRight = startX + totalTextWidth;
            foreach (var (word, wordW) in measuredWords)
            {
                float wordLeft = cursorRight - wordW;
                DrawShapedWordWithEmojis(canvas, word, shaper, font, paint, scaledFontSize, isRtl: true, wordLeft, wordW, baselineY);
                cursorRight = wordLeft - spaceW;
            }
        }
        else
        {
            float cursorLeft = startX;
            foreach (var (word, wordW) in measuredWords)
            {
                DrawShapedWordWithEmojis(canvas, word, shaper, font, paint, scaledFontSize, isRtl: false, cursorLeft, wordW, baselineY);
                cursorLeft += wordW + spaceW;
            }
        }

        if (includeSignatureLine)
        {
            using var flourishPaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = inkColor.WithAlpha(185),
                StrokeWidth = Math.Max(2f, scaledFontSize * 0.042f),
                StrokeCap = SKStrokeCap.Round,
                IsAntialias = true
            };

            float lineY = baselineY + (scaledFontSize * 0.24f);
            using var path = new SKPath();
            path.MoveTo(startX + (totalTextWidth * 0.05f), lineY + 4f);
            path.QuadTo(
                startX + (totalTextWidth * 0.5f),
                lineY - (scaledFontSize * 0.08f),
                startX + (totalTextWidth * 0.96f),
                lineY + 2f);
            canvas.DrawPath(path, flourishPaint);
        }

        using var snapshot = surface.Snapshot();
        using var encoded = snapshot.Encode(SKEncodedImageFormat.Png, 100);
        var bytes = encoded.ToArray();

        if (_previewCache.Count > 400)
        {
            _previewCache.Clear();
        }
        _previewCache[cacheKey] = bytes;

        double aspectRatio = bmpH > 0 ? (double)bmpW / bmpH : 3.0;
        return (bytes, aspectRatio);
    }

    /// <summary>
    /// Generates an executive double-bordered official stamp/seal PNG in memory using an OS font.
    /// </summary>
    public (byte[] PngBytes, double AspectRatio) RenderOfficialStampToPng(
        string primaryTitle,
        string secondarySubtitle,
        string fontFamily,
        string colorHex = "#1D4ED8")
    {
        if (string.IsNullOrWhiteSpace(primaryTitle))
        {
            primaryTitle = "APPROVED";
        }

        bool hasArabic = ContainsArabic(primaryTitle) || ContainsArabic(secondarySubtitle);
        using var boldTf = ResolveOsTypeface(fontFamily, isBold: true, isItalic: false, requiresArabicGlyphs: hasArabic);
        using var regTf = ResolveOsTypeface(fontFamily, isBold: false, isItalic: false, requiresArabicGlyphs: hasArabic);

        if (!SKColor.TryParse(colorHex, out var stampColor))
        {
            stampColor = new SKColor(29, 78, 216);
        }

        using var mainFont = new SKFont(boldTf, 54f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
        using var subFont = new SKFont(regTf, 28f) { Subpixel = true, Edging = SKFontEdging.SubpixelAntialias };
        using var mainShaper = new SKShaper(boldTf);
        using var subShaper = new SKShaper(regTf);

        float mainW = mainShaper.Shape(primaryTitle.Trim(), mainFont).Width;
        float subW = !string.IsNullOrWhiteSpace(secondarySubtitle)
            ? subShaper.Shape(secondarySubtitle.Trim(), subFont).Width
            : 0f;

        int bmpW = Math.Max(320, (int)Math.Ceiling(Math.Max(mainW, subW) + 96f));
        int bmpH = !string.IsNullOrWhiteSpace(secondarySubtitle) ? 156 : 116;

        var info = new SKImageInfo(bmpW, bmpH, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var surface = SKSurface.Create(info);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.Transparent);

        using var outerBorder = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            Color = stampColor,
            StrokeWidth = 6f,
            IsAntialias = true
        };
        using var innerBorder = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            Color = stampColor.WithAlpha(190),
            StrokeWidth = 2.2f,
            IsAntialias = true
        };
        using var fillTint = new SKPaint
        {
            Style = SKPaintStyle.Fill,
            Color = stampColor.WithAlpha(18),
            IsAntialias = true
        };

        var outerRect = new SKRect(8, 8, bmpW - 8, bmpH - 8);
        var innerRect = new SKRect(16, 16, bmpW - 16, bmpH - 16);

        canvas.DrawRoundRect(outerRect, 14f, 14f, fillTint);
        canvas.DrawRoundRect(outerRect, 14f, 14f, outerBorder);
        canvas.DrawRoundRect(innerRect, 9f, 9f, innerBorder);

        using var textPaint = new SKPaint
        {
            IsAntialias = true,
            Color = stampColor
        };

        float mainX = (bmpW - mainW) / 2f;
        float mainY = !string.IsNullOrWhiteSpace(secondarySubtitle) ? 74f : 76f;
        canvas.DrawShapedText(mainShaper, primaryTitle.Trim(), mainX, mainY, SKTextAlign.Left, mainFont, textPaint);

        if (!string.IsNullOrWhiteSpace(secondarySubtitle))
        {
            float subX = (bmpW - subW) / 2f;
            canvas.DrawShapedText(subShaper, secondarySubtitle.Trim(), subX, 118f, SKTextAlign.Left, subFont, textPaint);
        }

        using var snapshot = surface.Snapshot();
        using var encoded = snapshot.Encode(SKEncodedImageFormat.Png, 100);
        return (encoded.ToArray(), (double)bmpW / bmpH);
    }

    /// <summary>
    /// Removes white/paper background from an uploaded signature or stamp scan in memory,
    /// with smooth alpha feathering and optional ink re-coloring.
    /// </summary>
    public (byte[]? PngBytes, double AspectRatio) ProcessUploadedSignatureImage(
        byte[] imageBytes,
        bool removeWhiteBackground = true,
        string? tintColorHex = null,
        int luminanceThreshold = 232)
    {
        if (imageBytes == null || imageBytes.Length == 0)
            return (null, 3.0);

        using var srcBitmap = SKBitmap.Decode(imageBytes);
        if (srcBitmap == null || srcBitmap.Width <= 0 || srcBitmap.Height <= 0)
            return (null, 3.0);

        double ratio = (double)srcBitmap.Width / srcBitmap.Height;
        if (!removeWhiteBackground && string.IsNullOrWhiteSpace(tintColorHex))
        {
            using var img = SKImage.FromBitmap(srcBitmap);
            using var enc = img.Encode(SKEncodedImageFormat.Png, 100);
            return (enc.ToArray(), ratio);
        }

        SKColor? targetTint = null;
        if (!string.IsNullOrWhiteSpace(tintColorHex) && SKColor.TryParse(tintColorHex, out var parsedTint))
        {
            targetTint = parsedTint;
        }

        using var dstBitmap = new SKBitmap(srcBitmap.Width, srcBitmap.Height, SKColorType.Rgba8888, SKAlphaType.Premul);

        int width = srcBitmap.Width;
        int height = srcBitmap.Height;
        int featherStart = Math.Max(160, luminanceThreshold - 35);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                var c = srcBitmap.GetPixel(x, y);
                if (c.Alpha == 0)
                {
                    dstBitmap.SetPixel(x, y, SKColors.Transparent);
                    continue;
                }

                // Perceived luminance (0..255)
                int lum = (c.Red * 299 + c.Green * 587 + c.Blue * 114) / 1000;
                byte outAlpha = c.Alpha;

                if (removeWhiteBackground)
                {
                    if (lum >= luminanceThreshold)
                    {
                        dstBitmap.SetPixel(x, y, SKColors.Transparent);
                        continue;
                    }
                    else if (lum > featherStart)
                    {
                        float factor = (float)(luminanceThreshold - lum) / (luminanceThreshold - featherStart);
                        outAlpha = (byte)Math.Clamp((int)(c.Alpha * factor), 0, 255);
                    }
                }

                if (targetTint.HasValue)
                {
                    var t = targetTint.Value;
                    dstBitmap.SetPixel(x, y, new SKColor(t.Red, t.Green, t.Blue, outAlpha));
                }
                else
                {
                    dstBitmap.SetPixel(x, y, new SKColor(c.Red, c.Green, c.Blue, outAlpha));
                }
            }
        }

        using var finalImg = SKImage.FromBitmap(dstBitmap);
        using var finalEnc = finalImg.Encode(SKEncodedImageFormat.Png, 100);
        return (finalEnc.ToArray(), ratio);
    }

    private static void DrawShapedWordWithEmojis(
        SKCanvas canvas,
        string word,
        SKShaper shaper,
        SKFont font,
        SKPaint paint,
        float scaledFontSize,
        bool isRtl,
        float wordLeft,
        float wordW,
        float baselineY)
    {
        var tokens = EmojiRenderer.TokenizeByEmoji(word);
        float emojiBoxSize = scaledFontSize * 1.1f;
        float emojiAdvance = scaledFontSize * 1.15f;

        if (isRtl && tokens.Count > 1)
        {
            float tokenRight = wordLeft + wordW;
            foreach (var (tokenText, isEmoji) in tokens)
            {
                if (isEmoji)
                {
                    float x = tokenRight - emojiAdvance;
                    DrawEmojiToken(canvas, tokenText, x, baselineY - (scaledFontSize * 0.85f), emojiBoxSize);
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
                    DrawEmojiToken(canvas, tokenText, tokenLeft, baselineY - (scaledFontSize * 0.85f), emojiBoxSize);
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

    private static void DrawEmojiToken(SKCanvas canvas, string emojiCluster, float x, float y, float size)
    {
        var pngBytes = EmojiRenderer.RenderEmojiToPng(emojiCluster, 96);
        if (pngBytes == null || pngBytes.Length == 0) return;

        using var skImage = SKImage.FromEncodedData(pngBytes);
        if (skImage == null) return;

        var destRect = new SKRect(x, y, x + size, y + size);
        canvas.DrawImage(skImage, destRect, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
    }

    public static bool ContainsArabic(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (char c in text)
        {
            if ((c >= '\u0600' && c <= '\u06FF') ||
                (c >= '\u0750' && c <= '\u077F') ||
                (c >= '\uFB50' && c <= '\uFDFF') ||
                (c >= '\uFE70' && c <= '\uFEFF'))
            {
                return true;
            }
        }
        return false;
    }
}
