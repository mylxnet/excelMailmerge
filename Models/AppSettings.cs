using ExcelMailMerge.Helpers;

namespace ExcelMailMerge.Models;

/// <summary>
/// 高级设置：数据写入类型策略
/// </summary>
public class AdvancedWriteSettings : ObservableObject
{
    private bool _preserveDateTimeType = true;
    /// <summary>
    /// 日期写入类型：true=保持DateTime类型（套用模板格式），false=转字符串
    /// </summary>
    public bool PreserveDateTimeType { get => _preserveDateTimeType; set => SetProperty(ref _preserveDateTimeType, value); }

    private bool _preserveNumericType = true;
    /// <summary>
    /// 数字写入类型：true=保持数值类型，false=转字符串
    /// </summary>
    public bool PreserveNumericType { get => _preserveNumericType; set => SetProperty(ref _preserveNumericType, value); }

    private bool _autoLongNumberAsText = true;
    /// <summary>
    /// 长数字自动转文本（防科学计数法）：true=纯数字>11位强制文本格式
    /// </summary>
    public bool AutoLongNumberAsText { get => _autoLongNumberAsText; set => SetProperty(ref _autoLongNumberAsText, value); }

    private bool _writeDbNullForEmpty = true;
    /// <summary>
    /// 空值写入：true=DBNull（真正清空），false=空字符串
    /// </summary>
    public bool WriteDbNullForEmpty { get => _writeDbNullForEmpty; set => SetProperty(ref _writeDbNullForEmpty, value); }
}

/// <summary>
/// 输出模式
/// </summary>
public enum OutputMode
{
    /// <summary>
    /// 模式A：所有行→同一文件（每行复制全套模板Sheet）
    /// </summary>
    SingleFileMultiSheet = 0,

    /// <summary>
    /// 模式B：每行→独立文件（每行复制完整模板）
    /// </summary>
    MultiFilePerRow = 1
}

/// <summary>
/// 应用配置（持久化到settings.json）
/// </summary>
public class AppSettings
{
    public string? DataSourceFilePath { get; set; }
    public string? DataSourceSheetName { get; set; }
    public int TitleRowIndex { get; set; } = 1;
    public string? TemplateFilePath { get; set; }
    public OutputMode OutputMode { get; set; } = OutputMode.MultiFilePerRow;
    public string? PrimaryNamingColumn { get; set; }
    public string? SecondaryNamingColumn { get; set; }
    public string NamingSeparator { get; set; } = "_";
    public string? CustomOutputFolder { get; set; }
    public AdvancedWriteSettings WriteSettings { get; set; } = new();
}

