using System.Diagnostics;
using Conq.Noble.Spreadox;
using Microsoft.AspNetCore.Localization;
using Microsoft.AspNetCore.Mvc;
using ReportBuilderMvc.Models;
using ReportBuilderMvc.Services;

namespace ReportBuilderMvc.Controllers;

public class HomeController : Controller
{
    private readonly ReportGeneratorService _reportService;
    private readonly ILocalizationService _loc;
    private readonly IWebHostEnvironment _env;

    // Track loaded fonts for UI selection
    private static readonly List<string> AvailableFonts = new() { "Liberation Sans" };

    public static bool IsFontUploadAllowed =>
#if DEBUG
        true;
#else
        false;
#endif

    public HomeController(ReportGeneratorService reportService, ILocalizationService loc, IWebHostEnvironment env)
    {
        _reportService = reportService;
        _loc = loc;
        _env = env;
    }

    [HttpGet]
    public IActionResult Index()
    {
        var model = new ReportConfigViewModel
        {
            ReportTitle = _loc["DefaultReportTitle"],
            DataSeriesCount = 10,
            Layout = ChartLayout.SeparatePages,
            Format = OutputFormat.Pdf,
            Header = new HeaderConfig
            {
                Title = _loc["DefaultHeaderTitle"],
                Text = _loc["DefaultHeaderText"],
                FontSize = 12,
                Bold = true,
                TitleColor = "#005293"
            },
            Footer = new FooterConfig
            {
                Text = _loc["DefaultFooterText"],
                PageNumberFormat = _loc["DefaultPageNumberFormat"]
            },
            Cover = new CoverConfig
            {
                Company = _loc["DefaultCoverCompany"],
                Title = _loc["DefaultCoverTitle"],
                Provider = _loc["DefaultCoverProvider"],
                IncludeDate = true
            }
        };

        model.Items = _reportService.GenerateRandomData(model.DataSeriesCount);
        ViewBag.AvailableFonts = AvailableFonts;
        ViewBag.IsFontUploadEnabled = IsFontUploadAllowed;
        return View(model);
    }

    [HttpPost]
    public IActionResult SetLanguage(string culture, string returnUrl = "/")
    {
        Response.Cookies.Append(
            CookieRequestCultureProvider.DefaultCookieName,
            CookieRequestCultureProvider.MakeCookieValue(new RequestCulture(culture)),
            new CookieOptions
            {
                Expires = DateTimeOffset.UtcNow.AddYears(1),
                IsEssential = true,
                SameSite = SameSiteMode.Lax
            }
        );

        return LocalRedirect(string.IsNullOrEmpty(returnUrl) ? "/" : returnUrl);
    }

    [HttpPost]
    public IActionResult GenerateRandomData([FromBody] RandomDataRequest request)
    {
        int count = Math.Clamp(request.Count, 1, 30);
        var items = _reportService.GenerateRandomData(count);
        return Json(new { success = true, items });
    }

    [HttpPost]
#if DEBUG
    public async Task<IActionResult> UploadFont(IFormFile? fontFile, string? fontName)
    {
        if (fontFile == null || fontFile.Length == 0)
        {
            return Json(new { success = false, message = _loc["SelectFontAlert"] });
        }

        // Limit font file size (max 5 MB)
        const long maxSizeBytes = 5 * 1024 * 1024;
        if (fontFile.Length > maxSizeBytes)
        {
            return Json(new { success = false, message = _loc["FileSizeExceeded", 5] });
        }

        if (!fontFile.FileName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase))
        {
            return Json(new { success = false, message = _loc["OnlyTtfSupported"] });
        }

        string familyName = string.IsNullOrWhiteSpace(fontName)
            ? Path.GetFileNameWithoutExtension(fontFile.FileName)
            : fontName.Trim();

        // Sanitize family name to safe characters
        if (familyName.Length > 50 || familyName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            familyName = new string(familyName.Where(c => char.IsLetterOrDigit(c) || c == ' ' || c == '-' || c == '_').Take(50).ToArray()).Trim();
        }

        if (string.IsNullOrWhiteSpace(familyName))
        {
            familyName = "CustomFont";
        }

        using var memoryStream = new MemoryStream();
        await fontFile.CopyToAsync(memoryStream);
        byte[] fontBytes = memoryStream.ToArray();

        // Validate TTF magic bytes
        if (!IsValidFontHeader(fontBytes))
        {
            return Json(new { success = false, message = _loc["InvalidFontFile"] });
        }

        try
        {
            PdfFontRegistry.Register(familyName, fontBytes);

            lock (AvailableFonts)
            {
                if (!AvailableFonts.Contains(familyName, StringComparer.OrdinalIgnoreCase))
                {
                    AvailableFonts.Add(familyName);
                }
            }

            return Json(new { success = true, familyName, message = _loc["FontRegisteredSuccess", familyName] });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = _loc["FontRegistrationFailed", ex.Message] });
        }
    }

    private static bool IsValidFontHeader(byte[] bytes)
    {
        if (bytes == null || bytes.Length < 4) return false;
        // Standard TrueType: 00 01 00 00
        if (bytes[0] == 0x00 && bytes[1] == 0x01 && bytes[2] == 0x00 && bytes[3] == 0x00) return true;
        // 'true': 74 72 75 65
        if (bytes[0] == 0x74 && bytes[1] == 0x72 && bytes[2] == 0x75 && bytes[3] == 0x65) return true;
        // 'OTTO': 4F 54 54 4F
        if (bytes[0] == 0x4F && bytes[1] == 0x54 && bytes[2] == 0x54 && bytes[3] == 0x4F) return true;
        // 'typ1': 74 79 70 31
        if (bytes[0] == 0x74 && bytes[1] == 0x79 && bytes[2] == 0x70 && bytes[3] == 0x31) return true;
        return false;
    }
#else
    public IActionResult UploadFont(IFormFile? fontFile, string? fontName)
    {
        return StatusCode(StatusCodes.Status403Forbidden, new { success = false, message = _loc["FontUploadDisabledAlert"] });
    }
#endif

    [HttpPost]
    public async Task<IActionResult> Export(
        [FromForm] ReportConfigViewModel model,
        IFormFile? headerLogo,
        IFormFile? footerLogo)
    {
        const long maxLogoSizeBytes = 2 * 1024 * 1024; // 2 MB limit for logos
        byte[]? headerLogoBytes = null;
        if (headerLogo != null && headerLogo.Length > 0 && headerLogo.Length <= maxLogoSizeBytes)
        {
            using var ms = new MemoryStream();
            await headerLogo.CopyToAsync(ms);
            headerLogoBytes = ms.ToArray();
        }

        byte[]? footerLogoBytes = null;
        if (footerLogo != null && footerLogo.Length > 0 && footerLogo.Length <= maxLogoSizeBytes)
        {
            using var ms = new MemoryStream();
            await footerLogo.CopyToAsync(ms);
            footerLogoBytes = ms.ToArray();
        }

        string sanitizedTitle = string.IsNullOrWhiteSpace(model.ReportTitle)
            ? "Report"
            : string.Concat(model.ReportTitle.Split(Path.GetInvalidFileNameChars()));

        if (model.Format == OutputFormat.Excel)
        {
            byte[] excelBytes = _reportService.GenerateExcel(model);
            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{sanitizedTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        else
        {
            byte[] pdfBytes = _reportService.GeneratePdf(model, headerLogoBytes, footerLogoBytes);
            return File(
                pdfBytes,
                "application/pdf",
                $"{sanitizedTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
        }
    }

    [HttpGet]
    public IActionResult GetFonts()
    {
        return Json(AvailableFonts);
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}

public sealed class RandomDataRequest
{
    public int Count { get; set; } = 10;
}

public sealed class ErrorViewModel
{
    public string? RequestId { get; set; }
    public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
}
