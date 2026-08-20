using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ExcelProcessor.Services;
using ExcelProcessor.Models;

namespace ExcelProcessor.Pages;

public class IndexModel : PageModel
{
    private readonly ILogger<IndexModel> _logger;
    private readonly IExcelProcessingService _excelService;
    private readonly IWebHostEnvironment _environment;

    public IndexModel(
        ILogger<IndexModel> logger, 
        IExcelProcessingService excelService,
        IWebHostEnvironment environment)
    {
        _logger = logger;
        _excelService = excelService;
        _environment = environment;
    }

    public void OnGet()
    {
    }

    /// <summary>
    /// قراءة ملف Excel وعرض معاينة
    /// </summary>
    public async Task<IActionResult> OnPostPreviewAsync(IFormFile excelFile)
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { error = "لم يتم رفع أي ملف" });
            }

            await using var stream = excelFile.OpenReadStream();
            var (columns, rows, rowCount) = await _excelService.ReadExcelAsync(stream);

            return new OkObjectResult(new 
            { 
                columns,
                rows = rows.Take(10).ToList(),
                rowCount 
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error previewing file");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// الحصول على القيم الافتراضية من ملف النموذج
    /// </summary>
    public async Task<IActionResult> OnGetDefaultValuesAsync()
    {
        try
        {
            var goodSamplePath = Path.Combine(_environment.ContentRootPath, "..", "Good_sample_data.xlsx");
            
            if (!System.IO.File.Exists(goodSamplePath))
            {
                return NotFound(new { error = "ملف النموذج الصحيح غير موجود" });
            }

            await using var stream = System.IO.File.OpenRead(goodSamplePath);
            var defaults = _excelService.GetDefaultValuesFromSample(stream);

            return new OkObjectResult(defaults);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting default values");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// معالجة الملف الرئيسي
    /// </summary>
    public async Task<IActionResult> OnPostProcessAsync(
        IFormFile excelFile,
        IFormFile? goodSampleFile,
        IFormFile? natsFile,
        int bookingStart = 1,
        long identityStart = 1000000000,
        string defaultPhone = "966500000000")
    {
        try
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { error = "لم يتم رفع ملف Excel" });
            }

            // تحميل الملفات المرجعية
            var goodSamplePath = goodSampleFile?.Length > 0 
                ? null 
                : Path.Combine(_environment.ContentRootPath, "..", "Good_sample_data.xlsx");
            
            var natsPath = natsFile?.Length > 0 
                ? null 
                : Path.Combine(_environment.ContentRootPath, "..", "Nats.xlsx");

            if (!string.IsNullOrEmpty(goodSamplePath) && !System.IO.File.Exists(goodSamplePath))
            {
                return BadRequest(new { error = "ملف النموذج الصحيح غير موجود" });
            }

            if (!string.IsNullOrEmpty(natsPath) && !System.IO.File.Exists(natsPath))
            {
                return BadRequest(new { error = "ملف الجنسيات غير موجود" });
            }

            // إعدادات المعالجة
            var settings = new ProcessingSettings
            {
                BookingNumberStart = bookingStart,
                IdentityNumberStart = (int)identityStart,
                DefaultPhoneNumber = defaultPhone,
                DefaultValues = new Dictionary<string, string>()
            };

            // قراءة القيم الافتراضية من النموذج إذا لم يتم رفع ملف مخصص
            if (goodSampleFile?.Length == 0 && !string.IsNullOrEmpty(goodSamplePath))
            {
                await using var sampleStream = System.IO.File.OpenRead(goodSamplePath);
                var defaults = _excelService.GetDefaultValuesFromSample(sampleStream);
                foreach (var kvp in defaults)
                {
                    settings.DefaultValues[kvp.Key] = kvp.Value;
                }
            }

            // فتح streams للملفات
            await using var inputStream = excelFile.OpenReadStream();
            
            Stream goodSampleStream;
            if (goodSampleFile?.Length > 0)
            {
                goodSampleStream = goodSampleFile.OpenReadStream();
            }
            else if (!string.IsNullOrEmpty(goodSamplePath))
            {
                goodSampleStream = System.IO.File.OpenRead(goodSamplePath);
            }
            else
            {
                return BadRequest(new { error = "يجب توفير ملف النموذج الصحيح" });
            }

            Stream natsStream;
            if (natsFile?.Length > 0)
            {
                natsStream = natsFile.OpenReadStream();
            }
            else if (!string.IsNullOrEmpty(natsPath))
            {
                natsStream = System.IO.File.OpenRead(natsPath);
            }
            else
            {
                return BadRequest(new { error = "يجب توفير ملف الجنسيات" });
            }

            // معالجة الملف
            var result = await _excelService.ProcessExcelAsync(inputStream, settings, goodSampleStream, natsStream);

            if (result.Errors.Count > 0)
            {
                return StatusCode(500, new { errors = result.Errors });
            }

            // حفظ الملف مؤقتاً للتحميل
            var tempFile = Path.Combine(Path.GetTempPath(), result.FileName);
            await System.IO.File.WriteAllBytesAsync(tempFile, result.ProcessedFile);

            // تخزين مسار الملف المؤقت في الجلسة للتحميل لاحقاً
            HttpContext.Session.SetString("LastProcessedFile", tempFile);
            HttpContext.Session.SetString("LastProcessedFileName", result.FileName);

            return new OkObjectResult(new
            {
                result.TotalRows,
                result.TotalColumns,
                result.CorrectedValues,
                result.InvalidPhoneNumbers,
                result.MatchedNationalities,
                result.AddedColumns,
                result.ExtraColumnsKept,
                fileName = result.FileName,
                message = "تمت المعالجة بنجاح"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing file");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// تحميل الملف المعالج
    /// </summary>
    public IActionResult OnGetDownload()
    {
        try
        {
            var tempFile = HttpContext.Session.GetString("LastProcessedFile");
            var fileName = HttpContext.Session.GetString("LastProcessedFileName") ?? "Processed.xlsx";

            if (string.IsNullOrEmpty(tempFile) || !System.IO.File.Exists(tempFile))
            {
                return NotFound(new { error = "الملف غير موجود" });
            }

            var fileBytes = System.IO.File.ReadAllBytes(tempFile);
            
            // تنظيف الملف المؤقت بعد التحميل
            try
            {
                System.IO.File.Delete(tempFile);
            }
            catch
            {
                // تجاهل أخطاء الحذف
            }

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error downloading file");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
