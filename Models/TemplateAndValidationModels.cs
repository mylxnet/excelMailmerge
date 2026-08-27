namespace ExcelMailMerge.Models;

/// <summary>
/// 模板内容预览行（展示模板中所有单元格的原始内容，包括占位符和已有数据）
/// </summary>
public class TemplateContentRow
{
    /// <summary>所在Sheet名</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Excel行号（1-based）</summary>
    public int RowIndex { get; set; }

    /// <summary>列名→单元格值（列名为Excel列标如A、B、C，或表头行文本）</summary>
    public Dictionary<string, string> Values { get; set; } = new();

    /// <summary>按列名获取单元格值（用于 DataGrid 动态列绑定）</summary>
    public string this[string columnName] => Values.TryGetValue(columnName, out var v) ? v : string.Empty;

    /// <summary>显示用行标签</summary>
    public string RowLabel => $"{SheetName}!{RowIndex}";
}

/// <summary>
/// 模板中的占位符实例
/// </summary>
public class TemplatePlaceholder
{
    /// <summary>
    /// Trim后的占位符内部文本（如"姓名"，不含大括号）
    /// </summary>
    public string Key { get; set; } = string.Empty;

    /// <summary>
    /// 完整占位符文本（如"{姓名}"）
    /// </summary>
    public string FullText => $"{{{Key}}}";

    /// <summary>
    /// 出现次数
    /// </summary>
    public int OccurrenceCount { get; set; }

    /// <summary>
    /// 是否在数据源中有匹配的列
    /// </summary>
    public bool IsMatched { get; set; }

    /// <summary>
    /// 匹配状态文字
    /// </summary>
    public string MatchStatusText => IsMatched ? "✅ 已匹配" : "⚠️ 数据源无此列（将留空）";
}

/// <summary>
/// 扫描后的模板信息
/// </summary>
public class ScannedTemplate
{
    public string FilePath { get; set; } = string.Empty;
    public List<string> SheetNames { get; set; } = new();
    public int SheetCount => SheetNames.Count;
    public List<TemplatePlaceholder> Placeholders { get; set; } = new();

    /// <summary>
    /// 有多少个占位符没匹配到列
    /// </summary>
    public int UnmatchedCount => Placeholders.Count(p => !p.IsMatched);

    /// <summary>
    /// 是否与模式A冲突（Sheet>1且选模式A时提示）
    /// </summary>
    public bool HasConflictWithSingleFileMode => SheetCount > 1;
}

/// <summary>
/// 单个校验问题
/// </summary>
public class ValidationIssue
{
    /// <summary>
    /// 严重级别
    /// </summary>
    public ValidationLevel Level { get; set; }

    /// <summary>
    /// 关联的数据行号（0=不关联具体行）
    /// </summary>
    public int RelatedRowIndex { get; set; }

    /// <summary>
    /// 关联的列名（空=不关联具体列）
    /// </summary>
    public string RelatedColumnName { get; set; } = string.Empty;

    /// <summary>
    /// 问题描述
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// 建议修复方式
    /// </summary>
    public string Suggestion { get; set; } = string.Empty;

    public string LevelText => Level switch
    {
        ValidationLevel.Error => "❌ 错误",
        ValidationLevel.Warning => "⚠️ 警告",
        _ => "ℹ️ 提示"
    };

    public string LocationText
    {
        get
        {
            if (RelatedRowIndex <= 0 && string.IsNullOrEmpty(RelatedColumnName)) return "—";
            var parts = new List<string>();
            if (RelatedRowIndex > 0) parts.Add($"第{RelatedRowIndex}行");
            if (!string.IsNullOrEmpty(RelatedColumnName)) parts.Add($"[{RelatedColumnName}]列");
            return string.Join("，", parts);
        }
    }
}

public enum ValidationLevel
{
    Info = 0,
    Warning = 1,
    Error = 2
}

/// <summary>
/// 校验结果汇总
/// </summary>
public class ValidationResult
{
    public bool IsSuccess => Issues.Count(i => i.Level == ValidationLevel.Error) == 0;
    public List<ValidationIssue> Issues { get; set; } = new();
    public int ErrorCount => Issues.Count(i => i.Level == ValidationLevel.Error);
    public int WarningCount => Issues.Count(i => i.Level == ValidationLevel.Warning);
    public int InfoCount => Issues.Count(i => i.Level == ValidationLevel.Info);

    public string SummaryText => IsSuccess
        ? $"✅ 校验通过！共{Issues.Count}项提示（含警告{WarningCount}项）"
        : $"❌ 校验未通过：错误{ErrorCount}项，警告{WarningCount}项，请先修复所有错误";
}
