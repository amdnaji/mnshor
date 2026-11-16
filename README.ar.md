<div dir="rtl" align="center">

<img src="src/Mnshor/wwwroot/favicon.svg" alt="شعار مَنشُور" width="88" height="88" />

# مَنشُور — Mnshor

**محطة معالجة المستندات وملفات الـ PDF فائقة السرعة والمحلية 100% داخل الذاكرة العشوائية**

[![CI](https://github.com/amdnaji/mnshor/actions/workflows/ci.yml/badge.svg)](https://github.com/amdnaji/mnshor/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/amdnaji/mnshor?include_prereleases&label=Release)](https://github.com/amdnaji/mnshor/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![.NET 11.0](https://img.shields.io/badge/.NET-11.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Cross--Platform%20Web-0078D6)](#طريقة-التشغيل-والبناء-من-المصدر)
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20In--Memory%20%7C%20Zero%20Cloud-059669)](#لماذا-منشور)

**[English](README.md)** · **[العربية (Arabic)](README.ar.md)**

</div>

---

<div dir="rtl">

## نبذة عن المشروع

**مَنشُور (Mnshor)** هو منصة مفتوحة المصدر وعالية الأداء لمعالجة ملفات الـ PDF والمستندات المكتبية، صُممت لتعمل **محلياً بنسبة 100% داخل جهازك**. تتم جميع العمليات—من دمج وتقسيم وتدوير وإضافة العلامات المائية وتنظيف البيانات الوصفية، إلى تحويل ملفات Word (`.docx`) و Excel (`.xlsx`) و Markdown (`.md`) إلى PDF—**بشكل كامل داخل الذاكرة العشوائية (`MemoryStream` / RAM)** دون رفع أي ملف إلى السحابة ودون كتابة أي ملفات مؤقتة على القرص الصلب.

على نظام Windows، يمكن تشغيل **مَنشُور** كخدمة خلفية خفيفة (`MnshorService` تعمل حصرياً على العنوان المحلي `http://127.0.0.1:5157`) مع مشغّل فوري لسطح المكتب (`Mnshor.Launcher.exe`)، أو تشغيله مباشرةً كتطبيق ويب محلي مبني على `.NET 11`.

---

## لماذا مَنشُور؟

- **معالجة 100% في الذاكرة العشوائية (Zero Disk I/O)**: مستنداتك الحساسة (القانونية، المالية، والإدارية) لا تغادر جهازك إطلاقاً ولا تُحفظ في مجلدات مؤقتة على القرص.
- **دعم متقدم للغة العربية (`RTL`) وتشكيل الحروف**: يعتمد على محرك `SkiaSharp.HarfBuzz` لربط الحروف العربية بدقة عالية، مع دعم كامل لاتجاه النص الثنائي (`RTL` / `LTR`) والرموز التعبيرية (Emoji) والخطوط النظامية.
- **واجهة ثنائية اللغة (عربي / إنجليزي)**: تبديل فوري بين العربية والإنجليزية بنقرة واحدة، مع دعم الوضع الليلي (Dark) والنهاري (Light).
- **مُحسَّن للمستندات الضخمة (100+ صفحة)**: تقسيم صفحات المعاينة، الانتقال السريع لصفحة محددة، وأدوات تحكم دفعي بالنطاقات لمعالجة الملفات الكبيرة بسلاسة.
- **مؤشر زمن التنفيذ اللحظي**: تعرض كل عملية الزمن الفعلي المستغرق داخل الذاكرة بالملي ثانية (`ms`).

---

## الأدوات والمزايا المتوفرة

| الأداة | الوصف | أبرز المميزات |
| :--- | :--- | :--- |
| **1. دمج ملفات PDF** | دمج عدة ملفات PDF في مستند واحد عالي الدقة داخل الذاكرة. | سحب وإفلات لعدة ملفات، إعادة ترتيب الملفات قبل الدمج، عرض عدد الصفحات، وتنزيل فوري. |
| **2. تقسيم واستخراج PDF** | استخراج صفحات محددة أو فصل المستند بالكامل إلى ملفات مستقلة. | دعم كتابة النطاقات المخصصة (مثل `1-3, 5, 8-10`) أو **وضع الفصل الكامل (Burst)** الذي يضع كل صفحة في ملف PDF مستقل داخل أرشيف `.zip` في الذاكرة. |
| **3. تدوير صفحات PDF** | تدوير جميع الصفحات أو صفحات محددة بزاوية `90°` يميناً/يساراً أو `180°`. | معاينة مصغّرة للصفحات، تدوير فردي لكل صفحة أو تدوير دفعي (الفردية/الزوجية/نطاق مخصص)، مع تصفح مريح للملفات فوق 100 صفحة. |
| **4. حذف صفحات من PDF** | حذف صفحات غير مرغوب فيها بصرياً أو عبر تحديد أرقام الصفحات. | النقر المباشر على الصفحات لحذفها أو استعادتها، تحديد الفردية/الزوجية/عكس التحديد، وحماية تمنع حذف كامل صفحات المستند. |
| **5. إعادة ترتيب وعكس PDF** | إعادة ترتيب تسلسل الصفحات بصرياً أو عبر صيغة تسلسل مخصصة. | تحريك الصفحات للأول/الأخير/أعلى/أسفل، نقل لموضع محدد، **عكس الترتيب بالكامل (`N..1`)** بنقرة واحدة، أو إدخال تسلسل مخصص (`3, 1, 2, 4-6`). |
| **6. محرر البيانات الوصفية ومنظف الخصوصية** | عرض وتعديل البيانات الوصفية لملف PDF أو مسحها بالكامل لحماية الخصوصية. | تعديل (العنوان، المؤلف، الموضوع، الكلمات المفتاحية، برنامج الإنشاء) أو استخدام **منظف الخصوصية الفوري** لحذف كافة البيانات المخفية وسجلات XMP بنقرة واحدة قبل مشاركة الملف. |
| **7. إضافة علامة مائية إلى PDF** | ختم المستند بنص شفاف (مثل: `سري للغاية` / `مسودة`) أو شعار مصور. | دعم النصوص العربية والإنجليزية والصور (`PNG`/`JPG`/`WebP`)، معاينة حية، تحديد الموضع عبر شبكة 9 نقاط، **وضع الفسيفساء (التكرار الشبكي)**، التحكم بالشفافية والزاوية والطبقة (فوق/خلف المحتوى). |
| **8. تحويل Word و Excel و Markdown إلى PDF** | تحويل ملفات `.docx` و `.xlsx` و `.md` إلى مستندات PDF منسقة وجاهزة للطباعة. | **Word (`.docx`)**: العناوين، الجداول، القوائم، التنسيقات، الصور، ودعم العربية.<br>**Excel (`.xlsx`)**: استعراض أوراق العمل، معاينة فورية، اتجاه أفقي/عمودي، أحجام ورق (`A4`/`A3`/`Letter`)، اتجاه أعمدة (`RTL`/`LTR`)، تكرار الترويسة، وتلوين الصفوف بالتناوب.<br>**Markdown (`.md`)**: محرر مباشر مع نماذج جاهزة بالعربية والإنجليزية، جداول، كتل برمجية، واقتباسات. |

---

## البنية المعمارية للمشروع

</div>

```text
mnshor/
├── Mnshor.sln                                # ملف الحل البرمجي (.NET 11)
├── src/
│   ├── Mnshor/                               # تطبيق الويب المحلي (Blazor Interactive Server) وخدمة Windows
│   │   ├── Components/Pages/                 # واجهات الأدوات الثمان التفاعلية
│   │   ├── Engines/                          # محركات المعالجة في الذاكرة (PDF, Word, Excel, Markdown, Fonts)
│   │   ├── Models/                           # نماذج البيانات والخيارات ونتائج التنفيذ
│   │   ├── Services/                         # نظام الترجمة الفوري (عربي / إنجليزي) وإدارة الحالة
│   │   └── wwwroot/                          # التصميم (Prism CSS) ومحرك المعاينة PDF.js
│   └── Mnshor.Launcher/                      # المشغّل المكتبي لنظام Windows (يبدأ الخدمة ويفتح المتصفح)
├── installer/
│   ├── build-msi.ps1                         # سكربت أتمتة الاختبار والنشر وبناء حزمة التثبيت MSI
│   └── Mnshor.Installer/                     # إعدادات المثبّت المبني على WiX v4+
└── tests/
    └── Mnshor.Tests/                         # اختبارات الوحدة الشاملة (xUnit) لجميع المحركات
```

<div dir="rtl">

---

## طريقة التشغيل والبناء من المصدر

### المتطلبات الأساسية

- حزمة تطوير **[.NET 11.0 SDK](https://dotnet.microsoft.com/download/dotnet/11.0)** (مطلوبة لتشغيل المشروع وبنائه من المصدر).
- أداة **[WiX Toolset v4+](https://wixtoolset.org/)** (اختيارية — مطلوبة فقط في حال رغبتك ببناء ملف التثبيت `.msi` لنظام Windows).

### 1. استنساخ المستودع والتشغيل المباشر

</div>

```powershell
git clone https://github.com/amdnaji/mnshor.git
cd mnshor

dotnet run --project src/Mnshor/Mnshor.csproj
```

<div dir="rtl">

بعد التشغيل، افتح المتصفح على العنوان: **`http://localhost:5157`**.

### 2. تشغيل الاختبارات البرمجية (Unit Tests)

يحتوي المشروع على حزمة اختبارات `xUnit` للتحقق من دقة جميع محركات الدمج، التقسيم، التدوير، الحذف، إعادة الترتيب، تنظيف البيانات الوصفية، العلامة المائية، وتحويل ملفات Word و Excel و Markdown:

</div>

```powershell
dotnet test Mnshor.sln -c Release
```

<div dir="rtl">

### 3. بناء حزمة التثبيت لنظام Windows (`.msi`)

لبناء نسخة ذاتية الاحتواء (`self-contained win-x64`) داخل حزمة تثبيت `.msi`:

</div>

```powershell
# تثبيت أداة WiX وإضافة واجهة المستخدم (لمرة واحدة فقط):
dotnet tool install --global wix
wix extension add -g WixToolset.UI.wixext

# بناء ملف الـ MSI داخل مجلد ./dist/
.\installer\build-msi.ps1 -Version 0.1.0 -Stage alpha
```

<div dir="rtl">

---

## شكر خاص للمشاريع مفتوحة المصدر (Special Thanks)

بُني **مَنشُور** بالاعتماد على نخبة من المشاريع والمكتبات مفتوحة المصدر. نتوجه بجزيل الشكر والتقدير لجميع المطورين والمساهمين في هذه المشاريع الرائدة التي مكّنتنا من تقديم تجربة معالجة مستندات محلية وآمنة 100%:

| المشروع | دوره في مَنشُور | الترخيص |
| :--- | :--- | :--- |
| **[.NET 11 & ASP.NET Core Blazor](https://github.com/dotnet/aspnetcore)** | بيئة التشغيل عالية الأداء، مكونات الواجهة التفاعلية، واستضافة خدمة Windows (`Microsoft.Extensions.Hosting.WindowsServices`). | [MIT](https://github.com/dotnet/aspnetcore/blob/main/LICENSE.txt) |
| **[PDFsharp](https://github.com/empira/PDFsharp)** | المحرك الأساسي لقراءة وتعديل ملفات الـ PDF في الذاكرة (الدمج، التقسيم، التدوير، البيانات الوصفية، والرسم المتجهي). | [MIT](https://github.com/empira/PDFsharp/blob/master/LICENSE) |
| **[SkiaSharp](https://github.com/mono/SkiaSharp) & [HarfBuzz](https://github.com/harfbuzz/harfbuzz)** | معالجة الرسوميات ثنائية الأبعاد، دمج شعارات العلامة المائية، وتشكيل الحروف العربية المتصلة (`SkiaSharp.HarfBuzz`). | [MIT](https://github.com/mono/SkiaSharp/blob/main/LICENSE.md) |
| **[Open-XML-SDK (`DocumentFormat.OpenXml`)](https://github.com/dotnet/Open-XML-SDK)** | قراءة وتحليل مستندات Word (`.docx`) وجداول بيانات Excel (`.xlsx`) مباشرةً في الذاكرة دون الحاجة لتثبيت Microsoft Office. | [MIT](https://github.com/dotnet/Open-XML-SDK/blob/main/LICENSE) |
| **[Mozilla PDF.js](https://github.com/mozilla/pdf.js)** | عرض المعاينات المصغّرة لصفحات الـ PDF بسرعة عالية داخل المتصفح. | [Apache-2.0](https://github.com/mozilla/pdf.js/blob/master/LICENSE) |
| **[WiX Toolset](https://github.com/wixtoolset/wix)** | بناء حزمة التثبيت الرسمية `.msi` لنظام Windows وتسجيل الخدمة المحلية. | [MS-RL](https://wixtoolset.org/about/license/) |
| **[xUnit.net](https://github.com/xunit/xunit) & [Coverlet](https://github.com/coverlet-coverage/coverlet)** | إطار الاختبارات البرمجية المؤتمتة وقياس تغطية الشيفرة المصدرية. | [Apache-2.0](https://github.com/xunit/xunit/blob/main/LICENSE) / [MIT](https://github.com/coverlet-coverage/coverlet/blob/master/LICENSE) |
| **[IBM Plex Sans Arabic](https://github.com/IBM/plex) & [Plus Jakarta Sans](https://github.com/tokotype/PlusJakartaSans)** | الخطوط المفتوحة المصدر المستخدمة لظهور النصوص العربية والإنجليزية بأعلى وضوح وأناقة. | [OFL-1.1](https://openfontlicense.org/) |

---

## المساهمة في المشروع

نرحب بجميع المساهمات، بلاغات الأخطاء، واقتراحات المزايا الجديدة!

1. قم بعمل Fork للمستودع.
2. أنشئ فرعاً جديداً لميزتك (`git checkout -b feat/your-feature`).
3. تأكد من نجاح جميع الاختبارات (`dotnet test Mnshor.sln -c Release`).
4. أرسل طلب دمج (Pull Request).

---

## الترخيص (License)

هذا المشروع مفتوح المصدر ومرخّص بموجب **[رخصة MIT](LICENSE)**.

</div>
