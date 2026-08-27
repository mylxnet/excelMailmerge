namespace ExcelMailMerge.Models;

/// <summary>
/// 数据源Sheet信息
/// </summary>
public class DataSourceSheetInfo
{
    public string SheetName { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public bool HasData { get; set; }
}

/// <summary>
/// 列名信息
/// </summary>
public class ColumnInfo
{
    /// <summary>
    /// 原始列名（未Trim，用于匹配）
    /// </summary>
    public string OriginalName { get; set; } = string.Empty;

    /// <summary>
    /// Trim后的显示名
    /// </summary>
    public string DisplayName => OriginalName.Trim();

    /// <summary>
    /// 列索引（1-based）
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 占位符文本：{DisplayName}
    /// </summary>
    public string PlaceholderText => $"{{{DisplayName}}}";

    /// <summary>
    /// 是否在模板中被使用
    /// </summary>
    public bool IsUsedInTemplate { get; set; }
}

/// <summary>
/// 一行数据
/// </summary>
public class DataRow
{
    /// <summary>
    /// Excel中的原始行号（1-based）
    /// </summary>
    public int OriginalRowIndex { get; set; }

    /// <summary>
    /// 序号（显示用，从1开始）
    /// </summary>
    public int DisplayIndex { get; set; }

    /// <summary>
    /// 列名→单元格值
    /// </summary>
    public Dictionary<string, object?> Values { get; set; } = new();

    /// <summary>
    /// 索引器：按列名获取单元格值（用于 DataGrid 动态列绑定 Binding="[列名]"）
    /// </summary>
    public object? this[string columnName] => Values.TryGetValue(columnName, out var v) ? v : null;

    /// <summary>
    /// 原始行号显示（如"第3行"）
    /// </summary>
    public string RowLabel => $"第{OriginalRowIndex}行";
}

/// <summary>
/// 解析后的数据源
/// </summary>
public class ParsedDataSource
{
    public string FilePath { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public int TitleRowIndex { get; set; }
    public int DataStartRowIndex { get; set; }
    public List<ColumnInfo> Columns { get; set; } = new();
    public List<DataRow> Rows { get; set; } = new();

    /// <summary>
    /// 根据Trim后的列名查找列
    /// </summary>
    public ColumnInfo? FindColumnByDisplayName(string displayName)
    {
        return Columns.FirstOrDefault(c =>
            string.Equals(c.DisplayName, displayName.Trim(), StringComparison.Ordinal));
    }

    /// <summary>
    /// Trim后的列名集合
    /// </summary>
    public HashSet<string> DisplayNameSet => new(Columns.Select(c => c.DisplayName));
}
