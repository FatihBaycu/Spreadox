using System.Globalization;

namespace ReportBuilderMvc.Services;

/// <summary>
/// Metadata for a supported language in the system.
/// Adding a new language only requires adding an entry to SupportedLanguages list and dictionary keys.
/// </summary>
public sealed class SupportedLanguage
{
    public required string Code { get; init; }
    public required string TwoLetterCode { get; init; }
    public required string Name { get; init; }
    public required string Flag { get; init; }
    public required string Subtitle { get; init; }
}

public interface ILocalizationService
{
    string this[string key] { get; }
    string this[string key, params object[] args] { get; }
    IReadOnlyList<SupportedLanguage> SupportedLanguages { get; }
    SupportedLanguage CurrentLanguage { get; }
}

public sealed class LocalizationService : ILocalizationService
{
    public static readonly IReadOnlyList<SupportedLanguage> AvailableLanguages = new List<SupportedLanguage>
    {
        new() { Code = "en-US", TwoLetterCode = "en", Name = "English", Flag = "🇬🇧", Subtitle = "United States (Default)" },
        new() { Code = "tr-TR", TwoLetterCode = "tr", Name = "Türkçe", Flag = "🇹🇷", Subtitle = "Türkiye" },
        // To add a 3rd, 4th or 100th language, just add an entry here and corresponding translations in the dictionary below!
    };

    public IReadOnlyList<SupportedLanguage> SupportedLanguages => AvailableLanguages;

    public SupportedLanguage CurrentLanguage
    {
        get
        {
            string currentTwoLetter = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            return AvailableLanguages.FirstOrDefault(l => l.TwoLetterCode.Equals(currentTwoLetter, StringComparison.OrdinalIgnoreCase))
                ?? AvailableLanguages[0]; // fallback to default (en-US)
        }
    }

    public string this[string key] => GetString(key);

    public string this[string key, params object[] args]
    {
        get
        {
            string format = GetString(key);
            if (args == null || args.Length == 0)
            {
                return format;
            }
            try
            {
                return string.Format(CultureInfo.CurrentCulture, format, args);
            }
            catch
            {
                return format;
            }
        }
    }

    private string GetString(string key)
    {
        string langCode = CurrentLanguage.TwoLetterCode;
        if (Translations.TryGetValue(key, out var dict))
        {
            if (dict.TryGetValue(langCode, out var val))
            {
                return val;
            }
            if (dict.TryGetValue("en", out var defVal))
            {
                return defVal;
            }
        }
        return key;
    }

    private static readonly Dictionary<string, Dictionary<string, string>> Translations = new(StringComparer.OrdinalIgnoreCase)
    {
        // Navbar & Brand
        ["AppTitle"] = new() { ["en"] = "Spreadox Report Studio", ["tr"] = "Spreadox Rapor Stüdyosu" },
        ["Subtitle"] = new() { ["en"] = "High-Performance Excel & Vector PDF Engine", ["tr"] = "Yüksek Performanslı Excel ve Vektörel PDF Motoru" },
        ["FooterCopyright"] = new() { ["en"] = "© 2026 Conq.Noble Spreadox - Excel & PDF Chart Reporting Engine", ["tr"] = "© 2026 Conq.Noble Spreadox - Excel ve PDF Grafik Raporlama Motoru" },
        ["Language"] = new() { ["en"] = "Language", ["tr"] = "Dil" },
        ["SelectLanguageTitle"] = new() { ["en"] = "SELECT LANGUAGE", ["tr"] = "DİL SEÇİMİ" },

        // Top Action Bar
        ["ReportTitle"] = new() { ["en"] = "Report Title", ["tr"] = "Rapor Başlığı" },
        ["ReportTitlePlaceholder"] = new() { ["en"] = "Enter report title...", ["tr"] = "Rapor başlığı girin..." },
        ["PageLayout"] = new() { ["en"] = "Page Layout", ["tr"] = "Sayfa Düzeni" },
        ["SeparatePages"] = new() { ["en"] = "Separate Pages", ["tr"] = "Ayrı Sayfalarda" },
        ["SinglePage"] = new() { ["en"] = "Single Page", ["tr"] = "Aynı Sayfada" },
        ["SeriesCount"] = new() { ["en"] = "Series Count", ["tr"] = "Başlık Sayısı" },
        ["GenerateRandom"] = new() { ["en"] = "Generate", ["tr"] = "Üret" },
        ["Generating"] = new() { ["en"] = "Generating...", ["tr"] = "Üretiliyor..." },
        ["ExportPdf"] = new() { ["en"] = "Export PDF", ["tr"] = "PDF İndir" },
        ["ExportExcel"] = new() { ["en"] = "Export Excel", ["tr"] = "Excel İndir" },

        // Default Model Values
        ["DefaultReportTitle"] = new() { ["en"] = "Quarterly Business Performance", ["tr"] = "Çeyreklik İş Performansı Raporu" },
        ["DefaultHeaderTitle"] = new() { ["en"] = "Enterprise Analytics Report", ["tr"] = "Kurumsal Analitik Raporu" },
        ["DefaultHeaderText"] = new() { ["en"] = "Automated data extract generated with Spreadox", ["tr"] = "Spreadox ile otomatik oluşturulmuş veri çıktısı" },
        ["DefaultFooterText"] = new() { ["en"] = "Confidential - Spreadox Reporting Engine", ["tr"] = "Gizli - Spreadox Raporlama Motoru" },
        ["DefaultPageNumberFormat"] = new() { ["en"] = "Page {0} of {1}", ["tr"] = "Sayfa {0} / {1}" },
        ["DefaultCoverCompany"] = new() { ["en"] = "Spreadox Analytics Corp.", ["tr"] = "Spreadox Analitik A.Ş." },
        ["DefaultCoverTitle"] = new() { ["en"] = "Executive Data Summary", ["tr"] = "Yönetici Veri Özeti" },
        ["DefaultCoverProvider"] = new() { ["en"] = "Spreadox MVC Sample Suite", ["tr"] = "Spreadox MVC Örnek Uygulaması" },

        // Tabs
        ["TabDataCharts"] = new() { ["en"] = "1. Data & Charts", ["tr"] = "1. Veri & Grafikler" },
        ["TabHeader"] = new() { ["en"] = "2. Header Setup", ["tr"] = "2. Header (Üstbilgi)" },
        ["TabFooter"] = new() { ["en"] = "3. Footer Setup", ["tr"] = "3. Footer (Altbilgi)" },
        ["TabCover"] = new() { ["en"] = "4. Cover Page", ["tr"] = "4. Kapak Sayfası" },
        ["TabFonts"] = new() { ["en"] = "5. Font Registry (.TTF)", ["tr"] = "5. Font Desteği (.TTF)" },

        // Tab 1: Data & Charts
        ["GeneratedDataTitle"] = new() { ["en"] = "Generated Series & Chart Types", ["tr"] = "Üretilen Veriler ve Grafik Tipleri" },
        ["AddNewSeries"] = new() { ["en"] = "Add New Series", ["tr"] = "Yeni Başlık Ekle" },
        ["SeriesTitle"] = new() { ["en"] = "Series Title", ["tr"] = "Başlık Adı" },
        ["ChartType"] = new() { ["en"] = "Chart Type", ["tr"] = "Grafik Tipi" },
        ["SeriesColor"] = new() { ["en"] = "Series Color", ["tr"] = "Grafik Rengi" },
        ["Delete"] = new() { ["en"] = "Delete", ["tr"] = "Sil" },
        ["DeleteSeriesTooltip"] = new() { ["en"] = "Delete this series", ["tr"] = "Bu başlığı sil" },
        ["CategoriesLabel"] = new() { ["en"] = "Categories (Comma-separated)", ["tr"] = "Kategoriler (Virgülle ayrılmış)" },
        ["ValuesLabel"] = new() { ["en"] = "Values (Comma-separated - Edit directly)", ["tr"] = "Değerler (Virgülle ayrılmış - Doğrudan değiştirebilirsiniz)" },

        // Tab 2: Header
        ["HeaderTitle"] = new() { ["en"] = "Header Configuration", ["tr"] = "Header (Üstbilgi) Düzenleme" },
        ["HeaderDesc"] = new() { ["en"] = "Customize title, subtitle, logo image, font, and colors rendered at the top of each PDF page.", ["tr"] = "Her PDF sayfasının üstünde görünecek başlık, alt metin, logo ve font ayarları." },
        ["HeaderMainTitle"] = new() { ["en"] = "Header Main Title", ["tr"] = "Header Başlığı (Title)" },
        ["HeaderMainTitlePlaceholder"] = new() { ["en"] = "e.g. Monthly Executive Analytics", ["tr"] = "Örn: Aylık Performans Raporu" },
        ["HeaderSubtitle"] = new() { ["en"] = "Subtitle / Custom Text", ["tr"] = "Alt Metin / İstenilen Text (Subtitle)" },
        ["HeaderSubtitlePlaceholder"] = new() { ["en"] = "e.g. Automated report generated with Spreadox", ["tr"] = "Örn: Conq.Noble Spreadox ile otomatik oluşturulmuştur" },
        ["HeaderLogo"] = new() { ["en"] = "Header Logo Image (PNG, JPEG)", ["tr"] = "Header Logosu / Resim (PNG, JPEG)" },
        ["HeaderLogoDesc"] = new() { ["en"] = "Placed aspect-fit in the top-left corner.", ["tr"] = "Sayfanın sol üst köşesine orantılı olarak yerleştirilir." },
        ["FontFamily"] = new() { ["en"] = "Font Family", ["tr"] = "Yazı Tipi (Font Family)" },
        ["DefaultFont"] = new() { ["en"] = "Default (Liberation Sans)", ["tr"] = "Varsayılan (Liberation Sans)" },
        ["TitleColor"] = new() { ["en"] = "Title Color", ["tr"] = "Başlık Rengi" },
        ["FontSize"] = new() { ["en"] = "Font Size", ["tr"] = "Font Boyutu" },
        ["BoldTitle"] = new() { ["en"] = "Bold Title", ["tr"] = "Kalın (Bold) Başlık" },

        // Tab 3: Footer
        ["FooterConfigTitle"] = new() { ["en"] = "Footer Configuration", ["tr"] = "Footer (Altbilgi) Düzenleme" },
        ["FooterDesc"] = new() { ["en"] = "Customize footer text, dynamic page numbers, and optional logo.", ["tr"] = "Sayfaların altında çizgi ile ayrılmış altbilgi, sayfa numarası ve logo." },
        ["FooterText"] = new() { ["en"] = "Footer Text", ["tr"] = "Footer Metni (Text)" },
        ["FooterTextPlaceholder"] = new() { ["en"] = "e.g. Confidential - Internal Use Only", ["tr"] = "Örn: Gizli - Şirket İçi Kullanım İçindir" },
        ["PageNumberFormat"] = new() { ["en"] = "Page Number Format", ["tr"] = "Sayfa Numarası Formatı" },
        ["PageNumberFormatDesc"] = new() { ["en"] = "{0} = Current Page, {1} = Total Pages", ["tr"] = "{0} = Geçerli Sayfa, {1} = Toplam Sayfa" },
        ["FooterLogo"] = new() { ["en"] = "Footer Logo Image (PNG, JPEG)", ["tr"] = "Footer Logosu / Resim (PNG, JPEG)" },
        ["FooterLogoDesc"] = new() { ["en"] = "Rendered next to footer text.", ["tr"] = "Footer sol tarafında metin yanında küçük logo olarak görüntülenir." },

        // Tab 4: Cover
        ["CoverConfigTitle"] = new() { ["en"] = "PDF Cover Page Configuration", ["tr"] = "PDF Kapak Sayfası (Cover Page)" },
        ["CoverDesc"] = new() { ["en"] = "First page: company, title, date, provider and summary details.", ["tr"] = "PDF raporunun ilk sayfasındaki kurumsal başlık, şirket ve tarih bilgileri." },
        ["CompanyName"] = new() { ["en"] = "Company / Organization Name", ["tr"] = "Şirket / Kurum Adı" },
        ["CompanyNamePlaceholder"] = new() { ["en"] = "e.g. Conq.Noble Tech", ["tr"] = "Örn: Conq.Noble Tech" },
        ["CoverReportTitle"] = new() { ["en"] = "Cover Title", ["tr"] = "Kapak Başlığı" },
        ["CoverReportTitlePlaceholder"] = new() { ["en"] = "e.g. Annual Analytics & Performance Review", ["tr"] = "Örn: Yıllık Faaliyet ve Veri Raporu" },
        ["Provider"] = new() { ["en"] = "Report Provider / Author", ["tr"] = "Rapor Sağlayıcı (Provider)" },
        ["ProviderPlaceholder"] = new() { ["en"] = "e.g. Spreadox Reporting System", ["tr"] = "Örn: Spreadox Reporting System" },
        ["IncludeDate"] = new() { ["en"] = "Include Generation Date/Time", ["tr"] = "Oluşturulma Tarihini Ekle" },

        // Tab 5: Font Registry
        ["FontRegistryTitle"] = new() { ["en"] = "Custom Font Registration & Management", ["tr"] = "Özel Font Yükleme & Yönetimi" },
        ["FontRegistryDesc"] = new() { ["en"] = "Dynamically register TrueType (.ttf) font families into the Spreadox vector PDF engine at runtime.", ["tr"] = "Spreadox PDF motoruna dinamik olarak .ttf TrueType yazı tipi ekleyin." },
        ["FontFamilyName"] = new() { ["en"] = "Font Family Name", ["tr"] = "Font Ailesi Adı (Family Name)" },
        ["FontFamilyNamePlaceholder"] = new() { ["en"] = "e.g. Roboto, Montserrat, Open Sans", ["tr"] = "Örn: Roboto, Open Sans, Montserrat" },
        ["FontFileLabel"] = new() { ["en"] = "TrueType Font File (.ttf)", ["tr"] = "TrueType Font Dosyası (.ttf)" },
        ["UploadFontButton"] = new() { ["en"] = "Register Font", ["tr"] = "Fontu Kaydet" },
        ["UploadingFont"] = new() { ["en"] = "Registering...", ["tr"] = "Yükleniyor..." },
        ["SelectFontAlert"] = new() { ["en"] = "Please select a valid .ttf font file.", ["tr"] = "Lütfen bir .ttf font dosyası seçin." },
        ["OnlyTtfSupported"] = new() { ["en"] = "Only TrueType font (.ttf) files are supported.", ["tr"] = "Yalnızca TrueType font (.ttf) dosyaları desteklenir." },
        ["FontRegisteredSuccess"] = new() { ["en"] = "Font '{0}' registered successfully!", ["tr"] = "'{0}' fontu başarıyla sisteme kaydedildi!" },
        ["FontRegistrationFailed"] = new() { ["en"] = "Failed to register font: {0}", ["tr"] = "Font kaydedilemedi: {0}" },
        ["InstalledFonts"] = new() { ["en"] = "Registered System Fonts", ["tr"] = "Sistemde Yüklü Fontlar" },
        ["ActiveBadge"] = new() { ["en"] = "Active", ["tr"] = "Aktif" },
        ["NewlyAddedBadge"] = new() { ["en"] = "Newly Registered", ["tr"] = "Yeni Eklendi" },
        ["DebugOnlyBadge"] = new() { ["en"] = "Debug Only", ["tr"] = "Sadece Debug" },
        ["FontUploadDisabledTitle"] = new() { ["en"] = "Font Upload Disabled (Production Mode)", ["tr"] = "Font Yükleme Devre Dışı (Canlı Ortam)" },
        ["FontUploadDisabledNotice"] = new() { ["en"] = "For security reasons, dynamic font uploads are disabled in production mode. Custom font registration is only available during local debugging.", ["tr"] = "Güvenlik gerekçesiyle canlı (production) ortamda font yükleme devre dışıdır. Özel font yükleme özelliği yalnızca yerel debug modunda kullanılabilir." },
        ["FontUploadDisabledAlert"] = new() { ["en"] = "Font upload is disabled in production.", ["tr"] = "Canlı ortamda font yüklemeye izin verilmemektedir." },
        ["FileSizeExceeded"] = new() { ["en"] = "Font file size must be less than {0} MB.", ["tr"] = "Font dosyası boyutu {0} MB'dan küçük olmalıdır." },
        ["InvalidFontFile"] = new() { ["en"] = "The file does not appear to be a valid TrueType (.ttf) font.", ["tr"] = "Dosya geçerli bir TrueType (.ttf) fontu gibi görünmüyor." },

        // Dynamic JS templates
        ["NewSeriesDefaultTitle"] = new() { ["en"] = "New Series", ["tr"] = "Yeni Başlık" },
        ["SeriesPrefix"] = new() { ["en"] = "Series", ["tr"] = "Başlık" },
    };
}
