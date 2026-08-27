namespace ExcelMailMerge.Models;

/// <summary>
/// 生成进度报告
/// </summary>
public class GenerationProgress
{
    public int TotalRows { get; set; }
    public int CurrentRowIndex { get; set; }
    public int SuccessCount { get; set; }
    public int FailCount { get; set; }
    public int SkipCount { get; set; }
    public bool IsCompleted { get; set; }
    public bool IsCanceled { get; set; }
    public string OutputFolder { get; set; } = string.Empty;
    public List<string> Logs { get; set; } = new();
    public List<(int rowIndex, string fileName, string error)> Failures { get; set; } = new();

    public double ProgressPercent => TotalRows == 0 ? 0 : Math.Round((double)CurrentRowIndex / TotalRows * 100, 1);

    public string StatusText
    {
        get
        {
            if (IsCanceled) return $"⛔ 已取消：成功{SuccessCount}个，处理到第{CurrentRowIndex}行";
            if (IsCompleted) return FailCount == 0
                ? $"✅ 全部完成：成功{SuccessCount}个，耗时见日志"
                : $"⚠️ 完成但有失败：成功{SuccessCount}，失败{FailCount}";
            return $"⏳ 处理中：{CurrentRowIndex}/{TotalRows}（{ProgressPercent}%） 成功{SuccessCount} 失败{FailCount}";
        }
    }
}

/// <summary>
/// 命名组合结果
/// </summary>
public class NamingResult
{
    public bool IsValid { get; set; }
    public string Name { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
}

/// <summary>
/// 试生成（预览）结果
/// </summary>
public class PreviewResult
{
    /// <summary>是否成功</summary>
    public bool IsSuccess { get; set; }

    /// <summary>临时预览文件路径</summary>
    public string TempFilePath { get; set; } = string.Empty;

    /// <summary>使用的第几行数据（原始行号）</summary>
    public int SourceRowIndex { get; set; }

    /// <summary>临时文件名</summary>
    public string TempFileName { get; set; } = string.Empty;

    /// <summary>包含 Sheet 数</summary>
    public int SheetCount { get; set; }

    /// <summary>已替换占位符数</summary>
    public int PlaceholderReplacedCount { get; set; }

    /// <summary>未匹配占位符数</summary>
    public int PlaceholderUnmatchedCount { get; set; }

    /// <summary>公式转值数</summary>
    public int FormulaConvertedCount { get; set; }

    /// <summary>填充了数据的模板内容预览行（用于UI显示）</summary>
    public List<TemplateContentRow> FilledContentRows { get; set; } = new();

    /// <summary>填充了数据的模板内容预览列（Excel列标 A/B/C...）</summary>
    public List<string> FilledContentColumns { get; set; } = new();

    /// <summary>错误信息</summary>
    public string ErrorMessage { get; set; } = string.Empty;
}

