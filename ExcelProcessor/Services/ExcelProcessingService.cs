using System.Globalization;
using ClosedXML.Excel;
using ExcelProcessor.Models;

namespace ExcelProcessor.Services;

/// <summary>
/// خدمة معالجة ملفات Excel
/// </summary>
public interface IExcelProcessingService
{
    Task<(List<string> columns, List<Dictionary<string, string?>> rows, int rowCount)> ReadExcelAsync(Stream stream);
    Task<ProcessingResult> ProcessExcelAsync(Stream inputStream, ProcessingSettings settings, Stream goodSampleStream, Stream natsStream);
    Dictionary<string, string> GetDefaultValuesFromSample(Stream goodSampleStream);
}

/// <summary>
/// تطبيق خدمة معالجة Excel
/// </summary>
public class ExcelProcessingService : IExcelProcessingService
{
    private readonly ILogger<ExcelProcessingService> _logger;
    private readonly Dictionary<string, string> _nationalityMap = new();

    public ExcelProcessingService(ILogger<ExcelProcessingService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// قراءة ملف Excel وإرجاع الأعمدة والصفوف
    /// </summary>
    public Task<(List<string> columns, List<Dictionary<string, string?>> rows, int rowCount)> ReadExcelAsync(Stream stream)
    {
        return Task.Run(() =>
        {
            var columns = new List<string>();
            var rows = new List<Dictionary<string, string?>>();

            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);

            // قراءة الأعمدة من الصف الأول
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (int col = 1; col <= lastColumn; col++)
            {
                var cellValue = worksheet.Cell(1, col).GetValue<string>();
                columns.Add(cellValue ?? $"Column{col}");
            }

            // قراءة البيانات من الصف الثاني فما بعد
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            for (int row = 2; row <= lastRow; row++)
            {
                var rowData = new Dictionary<string, string?>();
                for (int col = 1; col <= columns.Count; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    rowData[columns[col - 1]] = GetCellValue(cell);
                }
                rows.Add(rowData);
            }

            return (columns, rows, rows.Count);
        });
    }

    /// <summary>
    /// الحصول على القيم الافتراضية من ملف النموذج الصحيح
    /// </summary>
    public Dictionary<string, string> GetDefaultValuesFromSample(Stream goodSampleStream)
    {
        var defaults = new Dictionary<string, string>();
        
        try
        {
            goodSampleStream.Position = 0;
            using var workbook = new XLWorkbook(goodSampleStream);
            var worksheet = workbook.Worksheet(1);

            // قراءة الأعمدة والقيم من أول صف بيانات
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            var columnNames = new List<string>();
            
            for (int col = 1; col <= lastColumn; col++)
            {
                var colName = worksheet.Cell(1, col).GetValue<string>();
                columnNames.Add(colName ?? $"Column{col}");
            }

            // قراءة قيم الصف الأول كبيانات افتراضية
            if (worksheet.RowCount() > 1)
            {
                for (int col = 1; col <= lastColumn; col++)
                {
                    var cell = worksheet.Cell(2, col);
                    var value = GetCellValue(cell);
                    if (!string.IsNullOrEmpty(value))
                    {
                        defaults[columnNames[col - 1]] = value;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error reading default values from sample");
        }

        return defaults;
    }

    /// <summary>
    /// تحميل خريطة الجنسيات من الملف المرجعي
    /// </summary>
    private void LoadNationalities(Stream natsStream)
    {
        _nationalityMap.Clear();
        natsStream.Position = 0;
        
        using var workbook = new XLWorkbook(natsStream);
        var worksheet = workbook.Worksheet(1);

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        for (int row = 2; row <= lastRow; row++)
        {
            var nat = worksheet.Cell(row, 1).GetValue<string>()?.Trim();
            var country = worksheet.Cell(row, 2).GetValue<string>()?.Trim();
            
            if (!string.IsNullOrEmpty(nat) && !string.IsNullOrEmpty(country))
            {
                // إضافة أشكال مختلفة من الجنسية للمطابقة المرنة
                _nationalityMap[nat] = country;
                
                // إضافة أشكال مرنة أخرى
                var normalizedNat = NormalizeText(nat);
                if (!string.IsNullOrEmpty(normalizedNat))
                {
                    _nationalityMap[normalizedNat] = country;
                }
                
                // إضافة صيغة المؤنث والمذكر
                if (nat.EndsWith("ي", StringComparison.Ordinal))
                {
                    var feminine = nat.Substring(0, nat.Length - 1) + "ية";
                    if (!_nationalityMap.ContainsKey(feminine))
                        _nationalityMap[feminine] = country;
                }
            }
        }
    }

    /// <summary>
    /// معالجة ملف Excel الرئيسي
    /// </summary>
    public async Task<ProcessingResult> ProcessExcelAsync(
        Stream inputStream, 
        ProcessingSettings settings, 
        Stream goodSampleStream, 
        Stream natsStream)
    {
        var result = new ProcessingResult();
        
        try
        {
            // تحميل الجنسيات
            LoadNationalities(natsStream);
            
            // الحصول على القيم الافتراضية من النموذج
            var defaultValues = GetDefaultValuesFromSample(goodSampleStream);
            
            // دمج مع القيم المخصصة من المستخدم
            foreach (var kvp in settings.DefaultValues)
            {
                if (!string.IsNullOrEmpty(kvp.Value))
                {
                    defaultValues[kvp.Key] = kvp.Value;
                }
            }

            // قراءة الملف المدخل
            inputStream.Position = 0;
            using var inputWorkbook = new XLWorkbook(inputStream);
            var inputWorksheet = inputWorkbook.Worksheet(1);

            // إنشاء ملف جديد للنتيجة
            using var outputWorkbook = new XLWorkbook();
            var outputWorksheet = outputWorkbook.Worksheets.Add("Processed Data");

            // قراءة أعمدة الملف المدخل
            var inputColumns = new List<string>();
            var lastInputColumn = inputWorksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (int col = 1; col <= lastInputColumn; col++)
            {
                var colName = inputWorksheet.Cell(1, col).GetValue<string>();
                inputColumns.Add(colName ?? $"Column{col}");
            }

            // قراءة أعمدة النموذج الصحيح
            goodSampleStream.Position = 0;
            using var sampleWorkbook = new XLWorkbook(goodSampleStream);
            var sampleWorksheet = sampleWorkbook.Worksheet(1);
            
            var requiredColumns = new List<string>();
            var lastSampleColumn = sampleWorksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (int col = 1; col <= lastSampleColumn; col++)
            {
                var colName = sampleWorksheet.Cell(1, col).GetValue<string>();
                if (!string.IsNullOrEmpty(colName))
                {
                    requiredColumns.Add(colName);
                }
            }

            // تحديد الأعمدة النهائية: المطلوبة أولاً، ثم الإضافية
            var finalColumns = new List<string>();
            var originalColumnsMap = new Dictionary<string, string>(); // Maps required column -> original column name
            var addedColumnsCount = 0;
            
            // إضافة الأعمدة المطلوبة أولاً
            foreach (var reqCol in requiredColumns)
            {
                finalColumns.Add(reqCol);
                
                // البحث عن عمود مطابق في الملف الأصلي (بما في ذلك المرادفات)
                string? matchingInputCol = null;
                if (inputColumns.Contains(reqCol))
                {
                    matchingInputCol = reqCol;
                }
                else
                {
                    // بحث عن مرادفات للأعمدة
                    matchingInputCol = FindMatchingColumn(reqCol, inputColumns);
                }
                
                if (matchingInputCol != null)
                {
                    originalColumnsMap[reqCol] = matchingInputCol;
                }
                else
                {
                    addedColumnsCount++;
                }
            }

            // إضافة الأعمدة الإضافية من الملف الأصلي (غير المطلوبة)
            var extraColumnsCount = 0;
            foreach (var inputCol in inputColumns)
            {
                // التحقق مما إذا كان هذا العمود قد تم استخدامه كمرادف لعمود مطلوب
                bool isUsedAsMapping = originalColumnsMap.Values.Contains(inputCol);
                bool isRequired = requiredColumns.Contains(inputCol);
                
                if (!isRequired && !isUsedAsMapping)
                {
                    finalColumns.Add(inputCol);
                    extraColumnsCount++;
                }
            }

            result.AddedColumns = addedColumnsCount;
            result.ExtraColumnsKept = extraColumnsCount;

            // كتابة رؤوس الأعمدة في الملف الناتج
            for (int col = 0; col < finalColumns.Count; col++)
            {
                outputWorksheet.Cell(1, col + 1).Value = finalColumns[col];
            }

            // معالجة الصفوف
            var lastInputRow = inputWorksheet.LastRowUsed()?.RowNumber() ?? 1;
            var bookingCounter = settings.BookingNumberStart;
            var identityCounter = settings.IdentityNumberStart;

            for (int row = 2; row <= lastInputRow; row++)
            {
                var outputRow = row;
                var correctionsCount = 0;

                // قراءة البيانات الأصلية
                var originalData = new Dictionary<string, string?>();
                foreach (var col in inputColumns)
                {
                    var colIndex = inputColumns.IndexOf(col) + 1;
                    originalData[col] = GetCellValue(inputWorksheet.Cell(row, colIndex));
                }

                // معالجة كل عمود نهائي
                foreach (var finalCol in finalColumns)
                {
                    var colIndex = finalColumns.IndexOf(finalCol) + 1;
                    string? value = null;
                    string? originalValue = null;
                    bool wasCorrected = false;

                    // البحث عن العمود المصدر (إما نفس الاسم أو عمودMapped)
                    string? sourceColumn = null;
                    if (originalData.ContainsKey(finalCol))
                    {
                        sourceColumn = finalCol;
                    }
                    else if (originalColumnsMap.TryGetValue(finalCol, out var mappedCol) && originalData.ContainsKey(mappedCol))
                    {
                        sourceColumn = mappedCol;
                    }
                    
                    if (sourceColumn != null)
                    {
                        value = originalData[sourceColumn];
                        
                        // تطبيق التصحيحات حسب نوع العمود
                        if (finalCol == "الجنسية" || finalCol == "Nat")
                        {
                            var correctedNat = CorrectNationality(value);
                            if (correctedNat != value)
                            {
                                originalValue = value;
                                value = correctedNat;
                                wasCorrected = true;
                                result.MatchedNationalities++;
                            }
                        }
                        else if (finalCol == "رقم الجوال" || finalCol.Contains("جوال", StringComparison.OrdinalIgnoreCase))
                        {
                            var correctedPhone = CorrectPhoneNumber(value, settings.DefaultPhoneNumber, out bool wasInvalid);
                            if (wasInvalid)
                            {
                                result.InvalidPhoneNumbers++;
                            }
                            if (correctedPhone != value)
                            {
                                originalValue = value;
                                value = correctedPhone;
                                wasCorrected = true;
                            }
                        }
                        else if (finalCol == "رقم الحجز")
                        {
                            if (string.IsNullOrEmpty(value) || value == "0")
                            {
                                value = bookingCounter.ToString();
                                bookingCounter++;
                            }
                        }
                        else if (finalCol == "رقم الهوية")
                        {
                            if (string.IsNullOrEmpty(value) || value == "0")
                            {
                                value = identityCounter.ToString();
                                identityCounter++;
                            }
                        }
                        else if (finalCol.Contains("تاريخ", StringComparison.OrdinalIgnoreCase))
                        {
                            var correctedDate = CorrectDate(value);
                            if (correctedDate != value)
                            {
                                originalValue = value;
                                value = correctedDate;
                                wasCorrected = true;
                            }
                        }
                    }
                    else
                    {
                        // العمود غير موجود، استخدام القيمة الافتراضية
                        if (defaultValues.TryGetValue(finalCol, out var defaultValue))
                        {
                            value = defaultValue;
                            
                            // للأرقام التسلسلية
                            if (finalCol == "رقم الحجز")
                            {
                                value = bookingCounter.ToString();
                                bookingCounter++;
                            }
                            else if (finalCol == "رقم الهوية")
                            {
                                value = identityCounter.ToString();
                                identityCounter++;
                            }
                        }
                        else
                        {
                            value = string.Empty;
                        }
                    }

                    // كتابة القيمة المصححة
                    outputWorksheet.Cell(outputRow, colIndex).Value = value ?? string.Empty;

                    // إذا تم التصحيح، إضافة عمود Original بجانبه
                    if (wasCorrected && !string.IsNullOrEmpty(originalValue))
                    {
                        var originalColName = $"Original_{finalCol}";
                        // التحقق مما إذا كان عمود Original موجوداً بالفعل
                        var existingOriginalCol = finalColumns.FirstOrDefault(c => c == originalColName);
                        if (existingOriginalCol == null)
                        {
                            // إضافة العمود في نهاية القائمة
                            var newColIndex = finalColumns.Count + 1;
                            finalColumns.Add(originalColName);
                            outputWorksheet.Cell(1, newColIndex).Value = originalColName;
                            result.TotalColumns++;
                        }
                        
                        var originalColIndex = finalColumns.IndexOf(originalColName) + 1;
                        outputWorksheet.Cell(outputRow, originalColIndex).Value = originalValue;
                        correctionsCount++;
                    }
                }

                result.CorrectedValues += correctionsCount;
                result.TotalRows++;
            }

            result.TotalColumns = finalColumns.Count;

            // حفظ الملف الناتج
            using var ms = new MemoryStream();
            outputWorkbook.SaveAs(ms);
            result.ProcessedFile = ms.ToArray();
            result.FileName = $"Processed_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing Excel file");
            result.Errors.Add(ex.Message);
        }

        return result;
    }

    /// <summary>
    /// تصحيح رقم الجوال
    /// </summary>
    private string CorrectPhoneNumber(string? value, string defaultPhone, out bool wasInvalid)
    {
        wasInvalid = false;
        
        if (string.IsNullOrEmpty(value))
        {
            wasInvalid = true;
            return defaultPhone;
        }

        // تنظيف الرقم من أي أحرف غير رقمية ومسافات
        var cleaned = new string(value.Where(char.IsDigit).ToArray());
        
        if (string.IsNullOrEmpty(cleaned))
        {
            wasInvalid = true;
            return defaultPhone;
        }
        
        // التعامل مع الصيغة الدولية 966XXXXXXXXX (12 رقم)
        if (cleaned.StartsWith("966") && cleaned.Length == 12)
        {
            // استخدام آخر 9 أرقام فقط كما في النموذج
            var last9 = cleaned.Substring(3);
            // التحقق من أن الأرقام الـ 9 تبدأ بـ 5 (للسعودية)
            if (last9.StartsWith("5"))
            {
                return last9;
            }
            // إذا لم تبدأ بـ 5، نستخدمها كما هي (قد تكون لدولة أخرى)
            return last9;
        }
        else if (cleaned.StartsWith("966") && cleaned.Length > 12)
        {
            // رقم طويل جداً، استخدام آخر 9 أرقام بعد 966
            if (cleaned.Length >= 15)
            {
                return cleaned.Substring(3, 9);
            }
            wasInvalid = true;
            return defaultPhone;
        }
        else if (cleaned.Length == 9)
        {
            // رقم محلي صحيح (9 أرقام)
            return cleaned;
        }
        else if (cleaned.StartsWith("05") && cleaned.Length == 10)
        {
            // رقم يبدأ بـ 05، تحويله إلى 9 أرقام بحذف الصفر
            return cleaned.Substring(1);
        }
        else if (cleaned.StartsWith("5") && cleaned.Length == 9)
        {
            // رقم يبدأ بـ 5 وطوله 9 أرقام
            return cleaned;
        }
        else if (cleaned.Length >= 9)
        {
            // محاولة استخراج 9 أرقام من النهاية
            return cleaned.Substring(cleaned.Length - 9);
        }
        else
        {
            // رقم غير صالح
            wasInvalid = true;
            return defaultPhone;
        }
    }

    /// <summary>
    /// تصحيح وتصحيح الجنسية باستخدام المطابقة المرنة
    /// </summary>
    private string CorrectNationality(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var normalized = NormalizeText(value);
        
        // البحث المباشر
        if (_nationalityMap.TryGetValue(value.Trim(), out var directMatch))
        {
            return directMatch;
        }

        // البحث مع النص الطبيعي
        if (_nationalityMap.TryGetValue(normalized, out var normalizedMatch))
        {
            return normalizedMatch;
        }

        // بحث مرن إضافي
        foreach (var kvp in _nationalityMap)
        {
            var keyNormalized = NormalizeText(kvp.Key);
            if (keyNormalized == normalized)
            {
                return kvp.Value;
            }
            
            // التحقق من التضمين
            if (normalized.Contains(keyNormalized) || keyNormalized.Contains(normalized))
            {
                return kvp.Value;
            }
        }

        // إرجاع القيمة الأصلية إذا لم يتم العثور على مطابقة
        return value;
    }

    /// <summary>
    /// تطبيع النص للمقارنة المرنة
    /// </summary>
    private string NormalizeText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return text
            .Trim()
            .Replace(" ", "")
            .Replace("أ", "ا")
            .Replace("إ", "ا")
            .Replace("آ", "ا")
            .Replace("ة", "ه")
            .Replace("ى", "ي")
            .ToLowerInvariant();
    }

    /// <summary>
    /// تصحيح التاريخ
    /// </summary>
    private string CorrectDate(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        // محاولة تحليل التاريخ بصيغ مختلفة
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
        {
            return parsedDate.ToString("yyyy-MM-dd");
        }

        // محاولة تحليل بصيغة DD/MM/YYYY أو MM/DD/YYYY
        var parts = value.Split('/', '-', '.');
        if (parts.Length == 3 && parts.All(p => int.TryParse(p, out _)))
        {
            int p1 = int.Parse(parts[0]);
            int p2 = int.Parse(parts[1]);
            int p3 = int.Parse(parts[2]);

            // سنة مكونة من خانتين
            if (p3 < 100)
            {
                p3 += p3 < 50 ? 2000 : 1900;
            }

            // تحديد ما إذا كان DD/MM/YYYY أو MM/DD/YYYY
            if (p1 > 12)
            {
                // بالتأكيد يوم/شهر/سنة
                return $"{p3}-{p2:D2}-{p1:D2}";
            }
            else if (p2 > 12)
            {
                // بالتأكيد شهر/يوم/سنة
                return $"{p3}-{p1:D2}-{p2:D2}";
            }
            else
            {
                // افتراض DD/MM/YYYY كصيغة عربية قياسية
                return $"{p3}-{p2:D2}-{p1:D2}";
            }
        }

        return value;
    }

    /// <summary>
    /// البحث عن عمود مطابق في القائمة باستخدام مرادفات معروفة
    /// </summary>
    private string? FindMatchingColumn(string requiredColumn, List<string> inputColumns)
    {
        // قائمة المرادفات للأعمدة الشائعة
        var columnSynonyms = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["رقم الحجز"] = new List<string> { "م", "الرقم التسلسلي", "تسلسل", "#", "ترتيب" },
            ["رقم الهوية"] = new List<string> { "الهوية", "رقم الهوية", "ID", "Identity" },
            ["رقم الجوال"] = new List<string> { "الهاتف", "جوال", "موبايل", "Phone", "Mobile", "Tel", "Contact" },
            ["اسم الحاج"] = new List<string> { "الاسم", "اسم", "Name", "المعتمر" },
            ["الجنس"] = new List<string> { "النوع", "Gender", "Sex" },
            ["تاريخ الميلاد"] = new List<string> { "الميلاد", "تاريخ", "DOB", "BirthDate", "DateOfBirth", "العمر" },
            ["الجنسية"] = new List<string> { "Nat", "Nationality", "دولة", "بلد", "القومية" },
            ["نوع الباقة"] = new List<string> { "الباقة", "Package", "ServiceType", "نوع الخدمة" },
            ["نوع المواصلات"] = new List<string> { "المواصلات", "Transport", "VehicleType", "النقل" },
            ["اسم الشركه"] = new List<string> { "الشركة", "Company", "Organization", "جهة", "المؤسسة" },
            ["المدينة"] = new List<string> { "مدينة", "City", "Location", "مقر التواجد", "العنوان" }
        };

        // البحث عن مرادف مطابق
        if (columnSynonyms.TryGetValue(requiredColumn, out var synonyms))
        {
            foreach (var synonym in synonyms)
            {
                var match = inputColumns.FirstOrDefault(c => 
                    c.Equals(synonym, StringComparison.OrdinalIgnoreCase) ||
                    c.Contains(synonym, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    return match;
                }
            }
        }

        // بحث جزئي إذا لم يتم العثور على مرادف
        foreach (var inputCol in inputColumns)
        {
            if (requiredColumn.Contains(inputCol, StringComparison.OrdinalIgnoreCase) ||
                inputCol.Contains(requiredColumn, StringComparison.OrdinalIgnoreCase))
            {
                return inputCol;
            }
        }

        return null;
    }

    /// <summary>
    /// الحصول على قيمة الخلية كنص
    /// </summary>
    private string GetCellValue(IXLCell cell)
    {
        try
        {
            var dataType = cell.DataType;
            
            switch (dataType)
            {
                case XLDataType.DateTime:
                    var dateValue = cell.GetValue<DateTime>();
                    return dateValue.ToString("yyyy-MM-dd");
                case XLDataType.Number:
                    var numValue = cell.GetValue<double>();
                    // إذا كان الرقم صحيحاً كبيراً (مثل رقم هوية)، إرجاعه بدون فواصل عشرية
                    if (numValue == Math.Floor(numValue) && numValue > 1000000)
                    {
                        return ((long)numValue).ToString();
                    }
                    return numValue.ToString(CultureInfo.InvariantCulture);
                case XLDataType.Boolean:
                    return cell.GetValue<bool>().ToString();
                case XLDataType.Text:
                case XLDataType.Blank:
                default:
                    return cell.GetValue<string>() ?? string.Empty;
            }
        }
        catch
        {
            return string.Empty;
        }
    }
}
