namespace ExcelProcessor.Models;

/// <summary>
/// نموذج يمثل إعدادات المعالجة التي يحددها المستخدم
/// </summary>
public class ProcessingSettings
{
    public int BookingNumberStart { get; set; } = 1;
    public int IdentityNumberStart { get; set; } = 1000000000;
    public string DefaultPhoneNumber { get; set; } = "966500000000";
    public Dictionary<string, string> DefaultValues { get; set; } = new();
}

/// <summary>
/// نتيجة معالجة الملف
/// </summary>
public class ProcessingResult
{
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public int CorrectedValues { get; set; }
    public int InvalidPhoneNumbers { get; set; }
    public int MatchedNationalities { get; set; }
    public int AddedColumns { get; set; }
    public int ExtraColumnsKept { get; set; }
    public byte[] ProcessedFile { get; set; } = Array.Empty<byte>();
    public string FileName { get; set; } = string.Empty;
    public List<string> Errors { get; set; } = new();
}

/// <summary>
/// نموذج لعرض البيانات في المعاينة
/// </summary>
public class PreviewRow
{
    public Dictionary<string, string?> Values { get; set; } = new();
}

/// <summary>
/// معلومات عن عمود مع/original
/// </summary>
public class ColumnInfo
{
    public string Name { get; set; } = string.Empty;
    public Type DataType { get; set; } = typeof(string);
    public string? DefaultValue { get; set; }
    public bool IsRequired { get; set; }
    public bool IsSequential { get; set; }
}
